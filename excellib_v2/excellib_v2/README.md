# excellib

Excel ファイル（`.xls` / `.xlsx`）の読み書きと印刷を行う C++17 ライブラリ。

- **外部サービス不要** — zlib のみに依存、完全オフライン動作
- **統一 API** — xls と xlsx を同じコードで扱える
- **印刷** — Excel COM 経由でレイアウトを完全再現して印刷・PDF 出力
- **バッチ印刷** — 複数ファイル・複数シートを Excel 1 起動で高速処理

---

## 動作環境

| 項目 | 要件 |
|------|------|
| C++ | C++17 以上 |
| OS（読み書き）| Windows / Linux / macOS |
| OS（印刷・PDF）| Windows（Excel インストール済み） |
| 依存ライブラリ | zlib のみ |

---

## ビルド

```bash
# CMake を使う場合
mkdir build && cd build
cmake .. -DCMAKE_BUILD_TYPE=Release
cmake --build .

# g++ で直接ビルドする場合
g++ -std=c++17 -Iinclude \
    src/common/common.cpp \
    src/common/factory.cpp \
    src/xls/xls_parser.cpp \
    src/xlsx/xlsx_parser.cpp \
    src/print/xlsx_page_setup.cpp \
    src/print/excel_printer.cpp \
    src/print/batch_printer.cpp \
    your_program.cpp \
    -lz -o your_program
```

---

## クイックスタート

```cpp
#include "excellib/excellib.hpp"
using namespace excellib;
```

### Excel ファイルを読む

```cpp
// xls / xlsx を同じ API で読む（拡張子ではなくマジックバイトで自動判定）
auto wb = WorkbookFactory::open("data.xlsx");
auto& sh = wb->sheet("売上");

// セルの読み取り
Cell c = sh.cell("B3");
if (is_string(c.value))  std::cout << get_string(c.value);
if (is_int(c.value))     std::cout << get_int(c.value);
if (is_numeric(c.value)) std::cout << as_double(c.value);

// 全非空セルを走査
sh.for_each_cell([](const Cell& c) {
    std::cout << c.address.to_a1() << ": " << to_string(c.value) << "\n";
});
```

### Excel ファイルを作る

```cpp
auto wb = WorkbookFactory::create(FileFormat::XLSX);
auto& sh = wb->add_sheet("売上");

sh.set_cell("A1", std::string{"製品名"});
sh.set_cell("B1", int64_t{100});
sh.set_cell("C1", 9.99);
sh.set_formula("D1", "B1*C1");

// 2D テーブルを一括書き込み
sh.set_table("A2", {
    {std::string{"WidgetA"}, int64_t{50}, 4.99},
    {std::string{"WidgetB"}, int64_t{30}, 9.99},
});

wb->save("output.xlsx");
```

### 印刷・PDF 出力

```cpp
ExcelPrinter printer;

// ページ設定
PageSetup ps;
ps.orientation = Orientation::Landscape;
ps.fit_to      = FitTo::Width;           // 全列を 1 ページ幅に収める
ps.print_area  = PrintArea::from_range("A1:H50");
ps.header_footer.odd_header = "&C&14売上レポート";
ps.header_footer.odd_footer = "&Lページ &P / &N&R&D";

// PDF 出力
auto r = printer.to_pdf("report.xlsx",
                         {.output_path = "report.pdf"},
                         "売上",   // シート名（空 = 先頭シート）
                         &ps);
if (!r.success) std::cerr << r.error_message;

// プリンター送信
printer.print("report.xlsx", {.printer_name = "Canon MF3010"}, "売上", &ps);
```

### バッチ印刷（複数ファイルを Excel 1 起動で高速処理）

```cpp
BatchPrinter batch;

PageSetup ps;
ps.orientation = Orientation::Landscape;
ps.fit_to      = FitTo::Width;

batch.add("東京_売上.xlsx",   "月次",  ps)
     .add("大阪_売上.xlsx",   "月次",  ps)
     .add("名古屋_売上.xlsx", "月次",  ps)
     .add("経費_Q1.xlsx",    "Sheet1");

// プリンターに一括送信（Excel 1 回起動で全件処理）
auto r = batch.print_all({.printer_name = "Canon MF3010"});

// または全シートを 1 つの PDF に結合
auto r = batch.to_pdf_merged("all_reports.pdf");

std::cout << r.succeeded << "/" << r.total << " 成功\n";
std::cout << r.total_elapsed_ms << "ms\n";
```

---

## ファイル構成

```
excellib/
├── include/excellib/
│   ├── excellib.hpp          ← これだけ #include すれば全機能使える
│   ├── workbook.hpp          ← 読み書き API（Cell / Sheet / Workbook）
│   ├── print_settings.hpp   ← ページ設定型定義
│   ├── excel_printer.hpp    ← 単一ファイル印刷・PDF
│   └── batch_printer.hpp    ← バッチ印刷・PDF
├── src/
│   ├── common/              ← 共通処理（CellAddress / factory）
│   ├── xls/                 ← XLS BIFF8 パーサー
│   ├── xlsx/                ← XLSX ZIP+XML パーサー
│   └── print/               ← 印刷エンジン
├── docs/
│   ├── ARCHITECTURE.md      ← 設計説明
│   ├── API_REFERENCE.md     ← API 全一覧
│   └── CHANGELOG.md         ← 変更履歴
├── examples/                ← サンプルコード
├── tests/                   ← テストスイート（92 件）
└── CMakeLists.txt
```

---

## テスト実行

```bash
cd build
ctest --output-on-failure
# または直接
./test_all
```

---

## 制限事項

- XLS（旧形式）は**読み込みのみ**。書き込みは XLSX 形式で行うこと
- 数式の**計算は非対応**（数式文字列の保持・読み書きは可能）
- 画像・グラフ・ピボットテーブルは**非対応**（セルデータのみ）
- 印刷・PDF は**Windows + Excel 必須**（Linux/macOS では apply_page_setup のみ使用可）
