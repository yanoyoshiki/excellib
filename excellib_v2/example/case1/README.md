# case1 — エクセル一括印刷 GUI

excellib を使った Win32 ネイティブ GUI アプリのサンプルです。
**マウス主導・大きなボタン・段階的やり直し** を意識した UI 設計になっています。

---

## 何ができるか

1. **入力** — 1 つの可変文字列を入力 (パスの一部として使われる)
2. **マスター読込** — `config.hpp` のテンプレートに文字列を埋め込んで Excel を開き、行/列構造のレコードを抽出
3. **ジョブ調整** — レコード一覧をチェックボックスで取捨、列範囲も個別編集可能
4. **実行 & 結果** — Excel 経由で一括印刷 (プリンター / 個別 PDF / 結合 PDF)、失敗ジョブのみ再実行可

各ステージは独立してやり直し可能 (前のステージへ戻ったり、現ステージだけリセット)。

---

## 環境設定 (config.hpp を編集)

`config.hpp` を自分の運用に合わせて書き換えてください:

```cpp
// パステンプレート (拡張子は .xlsx → .xls の順に自動試行)
constexpr std::string_view PATH_TEMPLATE = R"(C:\reports\{INPUT}\master)";

// マスター開始行 (0-based, 1 = 2 行目)
constexpr uint32_t MASTER_START_ROW = 1;

// マスターの列構成
constexpr uint32_t COL_FILE_PATH  = 0;   // A列
constexpr uint32_t COL_SHEET_NAME = 1;   // B列
constexpr uint32_t COL_PRINT_COLS = 2;   // C列  ("A:D" 等 / 空なら全列)
constexpr uint32_t COL_LABEL      = 3;   // D列
```

列構成が変わってもこの 1 ファイルだけ編集すれば追従できます。

---

## ビルド (Windows)

### MinGW (Git Bash / MSYS2)

```bash
cd example/case1
rm -rf build && mkdir build && cd build
cmake .. -G "MinGW Makefiles" -DCMAKE_BUILD_TYPE=Release
cmake --build .
./case1.exe
```

→ MinGW ランタイム DLL なしで動くスタンドアロン exe が `build/case1.exe` に生成されます。

### Visual Studio (MSVC)

Developer Command Prompt for VS から:

```cmd
cd example\case1
cmake -B build -DCMAKE_BUILD_TYPE=Release
cmake --build build --config Release
.\build\Release\case1.exe
```

---

## 動作要件

- **Windows 7 以降** (Visual Styles 6.0 + DPI Aware)
- **Microsoft Excel インストール必須** (印刷・PDF 出力に COM 経由で利用)

---

## ディレクトリ構成

```
example/case1/
├── CMakeLists.txt
├── config.hpp           ← 環境ごとに書き換え
├── main.cpp             ← WinMain
├── workflow.hpp/.cpp    ← GUI 非依存のロジック
├── ui_theme.hpp/.cpp    ← 配色・フォント・カスタム描画
├── ui_app.hpp/.cpp      ← メインウィンドウ + ステップバー
├── ui_stages.hpp        ← パネル基底クラス (inline 実装)
├── ui_stage1_input.cpp  ← 入力
├── ui_stage2_master.cpp ← マスター読込
├── ui_stage3_jobs.cpp   ← ジョブ調整
├── ui_stage4_run.cpp    ← 実行 & 結果
├── app.manifest         ← Visual Styles + DPI Awareness 宣言
└── app.rc               ← マニフェストを exe に埋め込むリソース
```

---

## マスター Excel のサンプル形式

`config.hpp` のデフォルト設定 (A列開始 / 2行目から) を前提にした例:

| | A (ファイル) | B (シート) | C (列範囲) | D (ラベル) |
|---|---|---|---|---|
| 1 | (ヘッダ) | (ヘッダ) | (ヘッダ) | (ヘッダ) |
| 2 | C:\data\東京.xlsx | 月次レポート | A:H | 東京月次 |
| 3 | C:\data\大阪.xlsx | 月次レポート | A:H | 大阪月次 |
| 4 | C:\data\福岡.xlsx | 月次レポート |  | 福岡月次 (全列) |

A列が空白になった行で読み取りが終了します (`Sheet::read_until_blank`)。
C 列が空のジョブは「全列」(FitTo::Width のみ) で印刷されます。

---

## トラブルシュート

### マスターが開かない
- `config.hpp` の `PATH_TEMPLATE` が実際のパスと一致しているか確認
- `{INPUT}` が含まれていることを確認
- ファイルが `.xlsx` でも `.xls` でも見つからない場合はバナーにエラーが出ます

### 印刷ボタンを押しても何も起きない
- Excel がインストールされているか
- マクロ警告ダイアログを表示する設定になっていないか (バックグラウンド実行のため)

### キャンセルボタンが効かない
- 実行中の VBScript は中断できません。次のジョブに移る前に停止します。
- 本当に強制終了したい場合は タスクマネージャから `EXCEL.EXE` を終了
