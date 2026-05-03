# API リファレンス

```cpp
#include "excellib/excellib.hpp"
using namespace excellib;
```

---

## WorkbookFactory

### open — ファイルを開く

```cpp
// ファイルパスから開く（xls / xlsx 自動判定）
std::unique_ptr<Workbook> WorkbookFactory::open(
    const std::string& path,
    const OpenOptions& opts = {}
);

// メモリ上のバイト列から開く
std::unique_ptr<Workbook> WorkbookFactory::open(
    const std::vector<uint8_t>& data,
    const OpenOptions& opts = {}
);
```

### create — 新規作成

```cpp
std::unique_ptr<Workbook> WorkbookFactory::create(
    FileFormat fmt = FileFormat::XLSX
);
```

### OpenOptions

```cpp
struct OpenOptions {
    bool            strict_validation{true};  // 形式違反で例外を投げる
    bool            strict{false};            // true のとき ParseWarning を ParseError に昇格
    WarningCallback on_warning;               // nullptr = 警告を stderr に出す
};
```

### ParseWarning / WarningCallback

```cpp
struct ParseWarning {
    enum class Kind { UnknownXmlElement, UnsupportedRecord, DataDropped, MalformedField };
    Kind        kind;
    std::string location;  // 例: "xl/worksheets/sheet1.xml row 42"
    std::string message;
};
using WarningCallback = std::function<void(const ParseWarning&)>;
```

---

## Workbook

```cpp
// メタ情報
FileFormat   wb.format();       // FileFormat::XLS or FileFormat::XLSX
std::string  wb.format_name();  // "XLS" or "XLSX"
size_t       wb.sheet_count();

// シートアクセス
Sheet& wb.sheet(size_t index);           // 範囲外は RangeError
Sheet& wb.sheet(const std::string& name);// 存在しなければ RangeError
std::vector<std::string> wb.sheet_names();

// シート管理
Sheet& wb.add_sheet(const std::string& name);              // 重複は WriteError
void   wb.remove_sheet(size_t index);
void   wb.rename_sheet(size_t index, const std::string& name);

// 保存
void                 wb.save(const std::string& path, const SaveOptions& opts = {});
std::vector<uint8_t> wb.to_bytes(FileFormat fmt, const SaveOptions& opts = {});
```

### SaveOptions

```cpp
struct SaveOptions {
    FileFormat format{FileFormat::Auto}; // Auto = 元の形式を維持
    // XLS 形式を指定すると WriteError が発生する
};
```

### FileFormat

```cpp
enum class FileFormat { XLS, XLSX, Auto };
```

---

## Sheet

### 読み取り

```cpp
std::string sh.name();
uint32_t    sh.row_count();   // 0-based の最大行 + 1
uint32_t    sh.col_count();   // 0-based の最大列 + 1

// セル取得（空セルでも例外を投げない。空のセルを返す）
Cell sh.cell(uint32_t row, uint32_t col);
Cell sh.cell(const CellAddress& addr);
Cell sh.cell(const std::string& a1);        // "B3" などの A1 記法

// 空セルは nullopt を返す安全版
std::optional<Cell> sh.try_cell(uint32_t row, uint32_t col);

// 行単位取得（列番号順でソート済み）
std::vector<Cell> sh.row(uint32_t row_idx);

// 全非空セル（行優先順）
std::vector<Cell> sh.cells();

// イテレーション
void sh.for_each_cell(std::function<void(const Cell&)> fn);
```

### 書き込み

```cpp
void sh.set_cell(uint32_t row, uint32_t col, const CellValue& v);
void sh.set_cell(const std::string& a1, const CellValue& v);
void sh.set_formula(const std::string& a1, const std::string& formula);

// 2D テーブルを一括書き込み
void sh.set_table(
    const std::string& top_left,              // 開始セル（A1 記法）
    const std::vector<std::vector<CellValue>>& data
);
```

---

## Cell

```cpp
struct Cell {
    CellAddress                address;   // {row, col} 0-based
    CellType                   type;      // Blank/Number/String/Boolean/Error/Formula
    CellValue                  value;
    std::optional<std::string> formula;   // 数式文字列（"SUM(A1:A10)" など）
    CellStyle                  style;     // フォーマット情報
    bool                       merged;    // 結合セルの一部か

    bool is_blank()    const;
    bool has_formula() const;
};
```

---

## CellValue

### 型述語

```cpp
bool is_blank  (const CellValue& v);
bool is_bool   (const CellValue& v);
bool is_int    (const CellValue& v);   // int64_t
bool is_double (const CellValue& v);
bool is_string (const CellValue& v);
bool is_error  (const CellValue& v);
bool is_numeric(const CellValue& v);   // int または double
```

### 厳格な取得（型不一致は std::bad_variant_access）

```cpp
bool        get_bool  (const CellValue& v);
int64_t     get_int   (const CellValue& v);
double      get_double(const CellValue& v);
std::string get_string(const CellValue& v);
ErrorValue  get_error (const CellValue& v);  // .code に "#REF!" など
```

### 変換

```cpp
double      as_double (const CellValue& v);  // int/double のみ受け付ける
std::string to_string (const CellValue& v);  // 全型を文字列に変換
```

### セルへの値設定

```cpp
sh.set_cell("A1", std::string{"テキスト"});   // 文字列
sh.set_cell("B1", int64_t{42});              // 整数
sh.set_cell("C1", 3.14);                     // 浮動小数点
sh.set_cell("D1", true);                     // 真偽値
sh.set_cell("E1", BlankValue{});             // 空（クリア）
sh.set_cell("F1", ErrorValue{"#REF!"});      // エラー値
```

---

## CellAddress

```cpp
struct CellAddress {
    uint32_t row;   // 0-based
    uint32_t col;   // 0-based

    static CellAddress from_a1(const std::string& a1);  // "B3" → {2,1}
    std::string to_a1() const;                           // {2,1} → "B3"
};
```

A1 記法の制限：行 1〜1048576、列 A〜XFD（1〜16384）。範囲外は `FormatError`。

---

## ページ設定

### PageSetup

```cpp
struct PageSetup {
    // 用紙・向き
    PaperSize   paper_size {PaperSize::A4};
    Orientation orientation{Orientation::Portrait};

    // 拡縮（fit_to と scale_percent は排他）
    FitTo    fit_to          {FitTo::None};
    uint16_t fit_to_pages_wide{1};       // FitTo::Width 時のページ数
    uint16_t fit_to_pages_tall{1};       // FitTo::Height 時のページ数
    uint16_t scale_percent   {100};      // FitTo::None 時（10〜400）

    // 余白（インチ）
    PageMargins margins;                 // left/right=0.70, top/bottom=0.75

    // ヘッダー・フッター
    HeaderFooter header_footer;

    // 印刷範囲・タイトル
    std::optional<PrintArea>    print_area;
    std::optional<RepeatTitles> repeat_titles;

    // 印刷オプション
    bool print_gridlines       {false};
    bool print_row_col_headings{false};
    bool black_and_white       {false};
    bool draft_quality         {false};
    bool page_order_over_then_down{false};

    uint16_t first_page_number{0};   // 0 = 自動
    uint16_t copies           {1};   // 印刷部数（print_all / print 時に有効）
};
```

### PaperSize

```cpp
enum class PaperSize : uint16_t {
    Letter = 1,
    A3     = 8,
    A4     = 9,   // デフォルト
    A5     = 11,
    B4     = 12,
    B5     = 13,
    A6     = 70,
    Custom = 0,
};
```

### FitTo

```cpp
enum class FitTo {
    None,           // scale_percent で拡縮
    Width,          // 全列を 1 ページ幅に収める（行は自由）
    Height,         // 全行を 1 ページ高さに収める（列は自由）
    WidthAndHeight, // 全体を 1 ページに収める
};
```

### PrintArea

```cpp
// "A1:H50" 形式の文字列から生成（逆順座標は自動正規化）
PrintArea pa = PrintArea::from_range("A1:H50");

// 文字列に戻す
std::string range = pa.to_range();  // "A1:H50"
```

### RepeatTitles — 全ページに繰り返す行・列

```cpp
RepeatTitles rt;
rt.row_start = 0; rt.row_end = 0;   // 1 行目を全ページに繰り返す
rt.col_start = 0; rt.col_end = 0;   // A 列を全ページに繰り返す
ps.repeat_titles = rt;
```

### HeaderFooter — ヘッダー・フッター

ヘッダー / フッター文字列で使える Excel コード：

| コード | 意味 |
|--------|------|
| `&L` | 左揃え |
| `&C` | 中央揃え |
| `&R` | 右揃え |
| `&P` | ページ番号 |
| `&N` | 総ページ数 |
| `&D` | 日付 |
| `&T` | 時刻 |
| `&F` | ファイル名 |
| `&A` | シート名 |
| `&14` | フォントサイズ 14pt |
| `&"Arial,Bold"` | フォント指定 |

```cpp
ps.header_footer.odd_header = "&C&\"Arial,Bold\"&14売上レポート";
ps.header_footer.odd_footer = "&Lページ &P / &N&R&D";
```

---

## ExcelPrinter

### apply_page_setup — XLSX にページ設定を書き込む

Excel 不要。ZIP/XML を直接操作する。

```cpp
ExcelPrinter printer;
printer.apply_page_setup(
    "report.xlsx",  // 対象ファイル（上書き）
    "Sheet1",       // シート名（空 = 先頭シート）
    ps              // ページ設定
);
```

### to_pdf — PDF 出力

```cpp
PrintResult r = printer.to_pdf(
    "report.xlsx",
    {.output_path = "out.pdf",  // 必須
     .open_after  = false},     // 生成後に開く
    "Sheet1",                   // シート名（省略可、空 = 先頭シート）
    &ps,                        // ページ設定（省略可、nullptr = ファイルの設定を使用）
    [](const std::string& msg) { std::cout << msg << "\n"; }  // 進捗コールバック（省略可）
);

if (!r.success) std::cerr << r.error_message;
std::cout << r.output_path;   // 出力先パス
```

### print — プリンター送信

```cpp
PrintResult r = printer.print(
    "report.xlsx",
    {.printer_name = "Canon MF3010",  // 空 = デフォルトプリンター
     .collate      = true,
     .duplex       = false},
    "Sheet1",
    &ps
);
```

### list_printers / default_printer

```cpp
std::vector<std::string> printers = printer.list_printers();
std::string def = printer.default_printer();
```

### PrintResult

```cpp
struct PrintResult {
    bool        success;
    std::string engine_used;     // "Excel"
    std::string error_message;   // 失敗時のエラー内容
    std::string output_path;     // PDF 出力時のパス
};
```

---

## BatchPrinter

### PrintJob

```cpp
// ファクトリ関数で生成
PrintJob job = PrintJob::make(
    "sales.xlsx",               // ファイルパス
    "Sheet1",                   // シート名（省略可、空 = 先頭シート）
    ps,                         // ページ設定（省略可）
    "Q1_Sales"                  // ラベル（進捗表示・PDF ファイル名に使用）
);

// または直接構築
PrintJob job = {"sales.xlsx", "Sheet1", ps, "Q1_Sales"};
```

### ジョブ登録

```cpp
BatchPrinter batch;

// チェーン記法で連続登録
batch.add("file1.xlsx", "Sheet1", ps, "label1")
     .add("file2.xlsx")          // シート・設定・ラベルは省略可
     .add(job3);                 // PrintJob を直接渡す

// ベクターでまとめて登録
batch.add_all(jobs_vector);

batch.job_count();  // 登録数
batch.clear();      // 全クリア
```

### 実行

```cpp
// プリンターに一括送信（Excel 1 起動で全件処理）
BatchResult r = batch.print_all(
    {.printer_name = "Canon MF3010"},
    [](size_t done, size_t total, const PrintJob& job, const std::string& msg) {
        std::cout << done << "/" << total << " " << job.file_path << "\n";
    }
);

// 全シートを 1 つの PDF に結合
BatchResult r = batch.to_pdf_merged(
    "all_reports.pdf",
    callback  // 省略可
);

// ジョブごとに個別 PDF
BatchResult r = batch.to_pdf_individual(
    "/output/dir",  // 保存先ディレクトリ
    callback        // 省略可
);
```

### BatchResult

```cpp
struct BatchResult {
    std::vector<JobResult> jobs;        // ジョブごとの結果
    size_t total;
    size_t succeeded;
    size_t failed;
    double total_elapsed_ms;            // 全体の処理時間
    std::string output_path;           // PDF 出力先（to_pdf_* 時のみ）

    bool all_success() const;

    // 失敗したジョブのポインター一覧
    std::vector<const JobResult*> failures() const;
};

struct JobResult {
    size_t      job_index;
    std::string file_path;
    std::string sheet_name;
    bool        success;
    std::string error_message;
    double      elapsed_ms;
};
```

---

## 例外

```cpp
// 全例外は ExcelError を基底とする
try {
    auto wb = WorkbookFactory::open("data.xlsx");
} catch (const ParseError& e)  { /* ファイル破損 */ }
  catch (const FormatError& e) { /* 未知形式・型不一致 */ }
  catch (const IOError& e)     { /* ファイル読み書き失敗 */ }
  catch (const RangeError& e)  { /* シート・セルの範囲外 */ }
  catch (const WriteError& e)  { /* シート名重複など */ }
  catch (const PrintError& e)  { /* 印刷・PDF エラー */ }
  catch (const ExcelError& e)  { /* 上記全部をまとめて受け取る */ }
```

---

## スレッド安全性

- `WorkbookFactory` のスタティックメソッドはスレッドセーフ
- `Workbook` / `Sheet` インスタンスはスレッドセーフでない（読み取り専用なら複数スレッドから参照可能）
- `ExcelPrinter` / `BatchPrinter` はスレッドセーフでない（インスタンスをスレッド間で共有しないこと）
