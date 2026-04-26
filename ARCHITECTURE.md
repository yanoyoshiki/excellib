# アーキテクチャ

## 全体構造

```
#include "excellib/excellib.hpp"
           │
           ▼
┌──────────────────────────────────────────────────┐
│                  公開 API 層                      │
│  WorkbookFactory  Sheet  CellValue  BatchPrinter  │
├────────────────────┬─────────────────────────────┤
│   XLS パーサー      │   XLSX パーサー              │
│                   │                              │
│  OLE2 解析        │  ZIP 展開（zlib）            │
│  BIFF8 レコード解析 │  XML 解析（自前）            │
│  SST / XF / RK   │  SharedStrings / Styles      │
├────────────────────┴─────────────────────────────┤
│              印刷エンジン                         │
│  PageSetup → XLSX XML 注入                       │
│  VBScript 生成 → Excel COM → 印刷 / PDF          │
└──────────────────────────────────────────────────┘
             ↓ 外部依存
          zlib のみ（印刷時は Excel COM も使用）
```

---

## フォーマット検出

ファイルを開くとき、**拡張子は一切参照しない**。先頭 8 バイトのマジックバイトで判定する。

```
D0 CF 11 E0 A1 B1 1A E1  →  OLE2 Compound File  →  XLS パーサー
50 4B 03 04               →  ZIP ファイル         →  XLSX パーサー
それ以外                  →  FormatError を投げる
```

---

## XLS パーサー（BIFF8 形式）

`.xls` ファイルの実体は OLE2 複合ドキュメント。その中の "Workbook" ストリームに BIFF8 レコードが並んでいる。

```
XLS ファイル
└── OLE2 コンテナ
    └── "Workbook" ストリーム
        ├── BOF            ブック開始
        ├── SST            共有文字列テーブル（文字列の実体）
        ├── XF 群          セル書式情報
        ├── FORMAT 群      数値フォーマット文字列
        ├── BOUNDSHEET 群  シート名と開始オフセット
        ├── BOF            シート1 開始
        │   ├── DIMENSION  シートの大きさ
        │   ├── NUMBER     数値セル（IEEE 754 double）
        │   ├── LABELSST   文字列セル（SST インデックス）
        │   ├── RK         圧縮数値（Excel 独自形式）
        │   ├── MULRK      同行複数 RK をまとめたレコード
        │   ├── BOOLERR    真偽値・エラー値
        │   └── FORMULA    数式セル
        └── EOF
```

**RK 値**は Excel 独自の数値圧縮形式。32 bit に以下を詰め込む。

```
bit 0: 1 なら 100 で割る
bit 1: 1 なら整数、0 なら IEEE 754 の上位 30 bit
bit 2-31: 値
```

---

## XLSX パーサー（OOXML 形式）

`.xlsx` は ZIP ファイル。展開すると XML が複数入っている。

```
report.xlsx（= ZIP）
├── [Content_Types].xml       ファイル種別定義
├── _rels/.rels               ルート関係定義
└── xl/
    ├── workbook.xml          シート一覧
    ├── _rels/workbook.xml.rels  rId → ファイルパスの対応表
    ├── sharedStrings.xml     文字列の実体（重複排除）
    ├── styles.xml            書式情報（日付フォーマット判定に使用）
    └── worksheets/
        ├── sheet1.xml        シート1 のセルデータ
        └── sheet2.xml
```

ZIP の読み書きは zlib だけで自前実装している。外部 ZIP ライブラリは不使用。

---

## セル値の型システム

```cpp
using CellValue = std::variant<
    BlankValue,    // 空セル
    bool,          // 真偽値
    int64_t,       // 整数（小数点なし・指数なし数値）
    double,        // 浮動小数点
    std::string,   // 文字列
    ErrorValue     // #REF! などのエラー値
>;
```

- **暗黙の型変換はしない**。型が違えば `std::bad_variant_access` を投げる
- `as_double()` だけは `int64_t` → `double` の明示的変換を許す
- `to_string()` は任意の型を文字列に変換するユーティリティ

---

## 印刷の仕組み

印刷は 2 段階で実現している。

### Stage 1: ページ設定を XLSX XML に書き込む

`apply_page_setup()` は Excel を使わず ZIP/XML を直接操作する。

```
apply_page_setup("report.xlsx", "Sheet1", ps)
  ↓
XLSX を ZIP として展開
  ↓
xl/worksheets/sheet1.xml の </sheetData> 直後に注入:
  <printOptions gridLines="1"/>
  <pageMargins left="0.5" right="0.5" .../>
  <pageSetup paperSize="9" orientation="landscape" fitToPage="1" .../>
  <headerFooter><oddHeader>...</oddHeader></headerFooter>
  ↓
xl/workbook.xml に追記（印刷範囲・タイトル行）:
  <definedNames>
    <definedName name="_xlnm.Print_Area">'Sheet1'!$A$1:$H$50</definedName>
    <definedName name="_xlnm.Print_Titles">'Sheet1'!$1:$1</definedName>
  </definedNames>
  ↓
ZIP に再梱包してファイルに書き戻す
```

### Stage 2: Excel COM で印刷・PDF 化

VBScript を一時ファイルに書き出して `cscript` で実行する。Excel のフル API にアクセスできる最もシンプルな方法。

```vbscript
Set xl = CreateObject("Excel.Application")
xl.Visible = False
Set wb = xl.Workbooks.Open("C:\report.xlsx")
Set ws = wb.Sheets("Sheet1")
' PageSetup を VBScript でも設定可能（apply_page_setup と二重設定も可）
ws.ExportAsFixedFormat Type:=0, Filename:="C:\report.pdf"
wb.Close False
xl.Quit
```

---

## バッチ印刷の高速化

`BatchPrinter` は Excel を **1 回だけ起動**して全ジョブを処理する。

```
通常（ジョブごとに Excel を起動する場合）:
  Excel 起動(3秒) → 処理(1秒) → 終了(1秒)  × ジョブ数
  → 10 ジョブ = 約 50 秒

BatchPrinter:
  Excel 起動(3秒) → 処理×10(10秒) → 終了(1秒)  × 1 回
  → 10 ジョブ = 約 14 秒（3〜4 倍速）
```

生成される VBScript は 1 つで、その中に全ジョブのループが含まれている。

```vbscript
Set xl = CreateObject("Excel.Application")
xl.ScreenUpdating = False   ' 画面更新オフで追加高速化

' --- Job 1 ---
Set wb = xl.Workbooks.Open("東京_売上.xlsx")
Set ws = wb.Sheets("月次")
ws.PageSetup.Orientation = 2  ' 横向き
ws.PrintOut Copies:=1
wb.Close False

' --- Job 2 ---
Set wb = xl.Workbooks.Open("大阪_売上.xlsx")
' ... 以降同様 ...

xl.Quit
```

---

## ZIP ライター実装と既知バグ（修正済み）

ZIP セントラルディレクトリエントリの正しいバイト構造：

```
offset 0  : sig(4) = PK\x01\x02
offset 4  : version_made(2)
offset 6  : version_needed(2)
offset 8  : flags(2)
offset 10 : compression(2)
offset 12 : mod_time(2)
offset 14 : mod_date(2)      ← これが欠落していたバグ
offset 16 : crc32(4)
offset 20 : comp_size(4)
offset 24 : uncomp_size(4)
offset 28 : filename_len(2)
...
offset 42 : local_header_offset(4)
offset 46 : filename(N)
```

`mod_date` の 2 バイトが欠落していたため、`apply_page_setup` 後に再オープンすると ZIP 破損エラーになっていた。現在は修正済み。
