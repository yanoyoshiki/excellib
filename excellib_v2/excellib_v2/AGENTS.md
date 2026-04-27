# AGENTS.md — excellib AI エージェント向けコンテキスト

このファイルは AI エージェントが excellib を即座に使いこなすためのコンテキストです。
人間向けドキュメント（`docs/`）より情報密度を優先しています。
ChatGPT / Claude / Gemini / Copilot など AI 全般を対象としています。

---

## 一言で言うと

**C++17 製の Excel (.xls/.xlsx) 読み書き + 印刷ライブラリ。外部依存ゼロ。**
- XLS 読み込み専用 / XLSX 読み書き両対応
- 印刷・PDF 出力は Windows + Excel 環境のみ
- macOS / Linux でも読み書きは動く

---

## ビルド（外部依存なし）

```bash
cmake -B build -DCMAKE_BUILD_TYPE=Release
cmake --build build -- -j4
# テスト
./build/test_all
```

ZLIB は不要（DEFLATE/CRC32 は `src/common/deflate.cpp` に内部実装済み）。
Windows では Visual Studio / MSVC でも同様にビルドできる。

---

## インクルード

```cpp
#include "excellib/excellib.hpp"   // これ1枚で全機能使える
using namespace excellib;
```

---

## ファイル構造

```
include/excellib/
  excellib.hpp          ← 統合ヘッダー（これだけ include すれば OK）
  workbook.hpp          ← Workbook / Sheet / CellValue の定義
  print_settings.hpp    ← PageSetup / PrintArea
  excel_printer.hpp     ← ExcelPrinter (単体印刷)
  batch_printer.hpp     ← BatchPrinter (複数ファイル一括印刷)
  xls_parser.hpp        ← XLS専用（直接使う必要はほぼない）
  xlsx_parser.hpp       ← XLSX専用（直接使う必要はほぼない）

src/
  common/common.cpp     ← CellAddress・型変換
  common/factory.cpp    ← フォーマット自動判定・ファイルオープン
  common/deflate.cpp    ← DEFLATE解凍 + CRC32（外部依存なし）
  xls/xls_parser.cpp    ← BIFF8 パーサー
  xlsx/xlsx_parser.cpp  ← ZIP展開 + XML解析 + 読み書き
  print/xlsx_page_setup.cpp ← ページ設定をXML直書き
  print/excel_printer.cpp   ← Excel COM 経由の印刷
  print/batch_printer.cpp   ← バッチ印刷

tests/test_all.cpp      ← 103件のテスト（DEFLATE含む）
examples/               ← 01_read.cpp ～ 05_error_handling.cpp
docs/                   ← API_REFERENCE.md / ARCHITECTURE.md
```

---

## 型システム（最重要）

```cpp
using CellValue = std::variant<
    BlankValue,   // 空セル
    bool,
    int64_t,
    double,
    std::string,
    ErrorValue    // #REF! など
>;

// 判定
is_blank(v)   is_bool(v)   is_int(v)
is_double(v)  is_string(v) is_error(v)  is_numeric(v)

// 取得（型が違えば std::bad_variant_access）
std::get<int64_t>(v)
std::get<std::string>(v)

// 便利変換
double  as_double(v)    // int → double 変換あり、他はエラー
std::string to_string(v) // 全型を文字列に

// CellAddress (0-based)
CellAddress a = CellAddress::from_a1("B3");  // row=2, col=1
a.to_a1()  // "B3"
```

---

## 読み込みパターン

```cpp
// ファイルから（xls/xlsx 自動判定）
auto wb = WorkbookFactory::open("data.xlsx");

// バイト列から（ネットワーク取得データなど）
std::vector<uint8_t> bytes = ...;
auto wb = WorkbookFactory::open(bytes);

// シート取得
Sheet& sh = wb->sheet(0);          // インデックス
Sheet& sh = wb->sheet("売上");     // 名前（なければ RangeError）

// セル読み取り
Cell c = sh.cell("B3");           // 存在しなければ BlankValue
auto opt = sh.try_cell("B3");     // std::optional<Cell>

if (is_string(c.value)) {
    auto s = std::get<std::string>(c.value);
}
if (is_numeric(c.value)) {
    double d = as_double(c.value);
}

// 全セル走査
sh.for_each_cell([](const Cell& c) {
    // c.address, c.value
});

// 行取得
std::vector<Cell> row = sh.row(0);   // 0-based
size_t nrows = sh.row_count();
size_t ncols = sh.col_count();
```

---

## 書き込みパターン

```cpp
// 新規作成
auto wb = WorkbookFactory::create(FileFormat::XLSX);
Sheet& sh = wb->add_sheet("売上");

// セル書き込み（型は明示が必要）
sh.set_cell("A1", std::string{"製品名"});
sh.set_cell("B1", int64_t{100});
sh.set_cell("C1", 3.14);
sh.set_cell("D1", true);
sh.set_formula("E1", "=B1*C1");

// 2D テーブル一括書き込み
std::vector<std::vector<CellValue>> table = {
    {std::string{"品名"}, int64_t{数量}},
    {std::string{"りんご"}, int64_t{50}},
};
sh.set_table("A1", table);

// 保存
wb->save("output.xlsx");

// バイト列として取得（ファイルに書かずに済む）
auto bytes = wb->to_bytes(FileFormat::XLSX);
```

---

## 印刷パターン（Windows + Excel が必要）

```cpp
#include "excellib/excel_printer.hpp"
#include "excellib/print_settings.hpp"

PageSetup ps;
ps.paper_size   = PaperSize::A4;
ps.orientation  = Orientation::Landscape;
ps.fit_to       = FitTo::Width{1};          // 横1ページに収める
ps.print_area   = PrintArea::from_range("A1", "H50");
ps.repeat_rows  = {0, 0};                   // 1行目を全ページで繰り返す
ps.margins      = {0.5, 0.5, 0.7, 0.7};    // 左右上下インチ

ExcelPrinter printer;
printer.apply_page_setup("report.xlsx", "Sheet1", ps);
printer.to_pdf("report.xlsx", "Sheet1", "report.pdf");
printer.print("report.xlsx", "Sheet1");     // プリンター送信
```

---

## バッチ印刷（複数ファイルを高速処理）

```cpp
#include "excellib/batch_printer.hpp"

BatchPrinter batch;
batch.add("東京.xlsx", "月次", ps, "東京月次")
     .add("大阪.xlsx", "月次", ps, "大阪月次")
     .add("福岡.xlsx", "月次", ps);         // ジョブ名省略可

// PDF 個別出力
auto result = batch.to_pdf_individual("/output/");

// PDF 1ファイルに結合
auto result = batch.to_pdf_merged("all.pdf");

// 結果確認
if (!result.all_success()) {
    for (auto* f : result.failures())
        std::cerr << f->file_path << ": " << f->error_message << "\n";
}
```

---

## 例外階層

すべての例外は `ExcelError` の派生クラス。

| 例外 | 発生条件 |
|---|---|
| `ParseError` | ファイル破損・ZIP/XML不正 |
| `FormatError` | 未知形式・型不一致 |
| `IOError` | ファイル読み書き失敗 |
| `RangeError` | シート/セルの範囲外アクセス |
| `WriteError` | シート名重複など書き込み制約違反 |
| `PrintError` | 印刷・PDF エラー（Windowsのみ） |

```cpp
try {
    auto wb = WorkbookFactory::open("missing.xlsx");
} catch (const IOError& e) {
    // ファイルが見つからない
} catch (const ExcelError& e) {
    // その他すべての excellib エラー
}
```

---

## 内部実装の要点（コード変更時に読む）

### フォーマット自動判定（factory.cpp）
拡張子ではなく**先頭8バイトのマジックバイト**で判定。
- `D0 CF 11 E0 ...` → OLE2 (XLS)
- `50 4B 03 04` → ZIP (XLSX)

### DEFLATE 解凍（deflate.cpp）
`src/common/deflate.cpp` に RFC 1951 準拠の実装。
- `deflate_decompress(src, src_len, hint)` — raw DEFLATE（zlib/gzip ヘッダーなし）
- `crc32_compute(data, len)` — IEEE 802.3 CRC32
- zlib には一切依存していない

### ZIP 読み書き（xlsx_parser.cpp / xlsx_page_setup.cpp）
外部 ZIP ライブラリなし。セントラルディレクトリを直接バイト操作。
書き込みは常に **STORE (method=0)** で無圧縮。
読み込みは STORE と DEFLATE (method=8) の両対応。

### XML 解析（xlsx_parser.cpp）
外部 XML ライブラリなし。`xml_attr()` / `xml_text()` という自前の単純な文字列マッチャー。
名前空間プレフィックス付き属性（`r:id="rId1"` など）も一部対応。

### XLSX 書き込みの流れ
1. XML テンプレートを文字列結合で生成
2. 各エントリを ZIP STORE 形式でメモリ上に組み立て
3. `Workbook::save()` / `to_bytes()` でファイルまたはバイト列に出力

### ページ設定の注入（xlsx_page_setup.cpp）
Excel を使わず ZIP/XML を直接書き換える。
`</sheetData>` の直後に `<pageSetup>` ブロックを文字列挿入。
`xl/workbook.xml` の `<definedNames>` に印刷範囲・タイトル行を追記。

---

## よくある落とし穴

```cpp
// NG: 文字列リテラルは自動で std::string にならない
sh.set_cell("A1", "テキスト");            // コンパイルエラー
// OK:
sh.set_cell("A1", std::string{"テキスト"});

// NG: XLS は書き込み不可
auto wb = WorkbookFactory::open("data.xls");
wb->save("out.xls");                      // 実行時エラー

// NG: CellAddress は 0-based、A1 記法と混同しやすい
sh.cell(CellAddress{0, 0});  // A1
sh.cell(CellAddress{1, 1});  // B2

// 注意: 印刷機能は Windows + Excel インストール必須
// macOS/Linux で ExcelPrinter を使うとコンパイルは通るが実行時エラー
```

---

## テスト実行

```bash
./build/test_all
# 期待: 103 passed  0 failed  0 skipped
```

テスト内訳:
- CellAddress / CellValue / Sheet API / XLSX roundtrip
- XLS 内部関数 / 印刷設定 / バッチ印刷
- **DEFLATE解凍** (stored / fixed huffman / dynamic huffman / CRC32) ← zlib削除後に追加

---

## ブランチ戦略

| ブランチ | 用途 |
|---|---|
| `main` | リリース済み安定版 |
| `develop` | 開発統合 |
| `feature/*` `fix/*` `docs/*` | トピックブランチ |

コミットメッセージは **Conventional Commits** 形式かつ**日本語**。
例: `feat: バッチ印刷でシート名省略を可能に`
