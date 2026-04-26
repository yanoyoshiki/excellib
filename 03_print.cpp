/**
 * examples/03_print.cpp
 *
 * 単一ファイルの印刷・PDF 出力サンプル。
 *
 * 実行:
 *   ./03_print
 */
#include "excellib/excellib.hpp"
#include <iostream>

using namespace excellib;

int main() {
    // ---- テスト用 XLSX 作成 ----
    {
        auto wb = WorkbookFactory::create();
        auto& sh = wb->add_sheet("売上レポート");
        sh.set_table("A1", {
            {std::string{"製品"},  std::string{"Q1"},      std::string{"Q2"},      std::string{"合計"}},
            {std::string{"A"},     int64_t{1200000},    int64_t{1350000},    int64_t{2550000}},
            {std::string{"B"},     int64_t{ 980000},    int64_t{1100000},    int64_t{2080000}},
            {std::string{"C"},     int64_t{ 760000},    int64_t{ 820000},    int64_t{1580000}},
            {std::string{"合計"},  int64_t{2940000},    int64_t{3270000},    int64_t{6210000}},
        });
        wb->save("report.xlsx");
    }

    ExcelPrinter printer;

    // ---- ページ設定 ----
    PageSetup ps;
    ps.paper_size   = PaperSize::A4;
    ps.orientation  = Orientation::Landscape;    // 横向き
    ps.fit_to       = FitTo::Width;              // 全列を 1 ページ幅に収める
    ps.print_area   = PrintArea::from_range("A1:D5");
    ps.margins.left = ps.margins.right = 0.5;

    // ヘッダー・フッター
    // &C=中央  &14=14pt  &P=ページ番号  &N=総ページ数  &D=日付
    ps.header_footer.odd_header = "&C&\"Arial,Bold\"&14売上レポート";
    ps.header_footer.odd_footer = "&Lページ &P / &N&R&D";

    // 1 行目を全ページに繰り返す
    RepeatTitles rt;
    rt.row_start = 0; rt.row_end = 0;
    ps.repeat_titles = rt;

    // ---- Option A: ページ設定をファイルに焼き込む ----
    // Excel で開いても設定が残る。印刷プレビューで確認してから印刷する用途に向く
    printer.apply_page_setup("report.xlsx", "売上レポート", ps);
    std::cout << "ページ設定を report.xlsx に書き込みました\n";

    // ---- Option B: PDF 出力 ----
    auto r = printer.to_pdf(
        "report.xlsx",
        {.output_path = "report.pdf"},
        "売上レポート",
        nullptr,    // nullptr = ファイルに焼き込まれた設定を使用
        [](const std::string& msg) { std::cout << "[PDF] " << msg << "\n"; }
    );
    if (r.success)
        std::cout << "PDF を生成: " << r.output_path << "\n";
    else
        std::cout << "PDF 失敗: " << r.error_message << "\n";

    // ---- Option C: プリンターに直接送信 ----
    // デフォルトプリンターを確認
    std::string def_printer = printer.default_printer();
    std::cout << "デフォルトプリンター: " << def_printer << "\n";

    // プリンター一覧
    for (auto& p : printer.list_printers())
        std::cout << "  - " << p << "\n";

    // 印刷（PageSetup を直接渡す方法）
    auto pr = printer.print(
        "report.xlsx",
        {.printer_name = def_printer},   // 空文字でもデフォルトプリンターに送られる
        "売上レポート",
        &ps   // ページ設定をその場で指定（ファイルを書き換えない）
    );
    if (pr.success)
        std::cout << "印刷を送信しました\n";
    else
        std::cout << "印刷失敗: " << pr.error_message << "\n";

    return 0;
}
