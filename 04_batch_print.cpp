/**
 * examples/04_batch_print.cpp
 *
 * バッチ印刷サンプル。
 * 複数ファイル・複数シートを Excel 1 起動で高速処理する。
 *
 * 実行:
 *   ./04_batch_print
 */
#include "excellib/excellib.hpp"
#include <iostream>
#include <iomanip>

using namespace excellib;

// テスト用 XLSX を作成するヘルパー
static void make_test_file(const std::string& path,
                            const std::string& sheet,
                            const std::string& title,
                            int base_value) {
    auto wb = WorkbookFactory::create();
    auto& sh = wb->add_sheet(sheet);
    sh.set_table("A1", {
        {std::string{title}, std::string{"Q1"}, std::string{"Q2"}, std::string{"Q3"}},
        {std::string{"製品A"}, int64_t{base_value},     int64_t{base_value+100}, int64_t{base_value+200}},
        {std::string{"製品B"}, int64_t{base_value+50},  int64_t{base_value+150}, int64_t{base_value+250}},
        {std::string{"製品C"}, int64_t{base_value+100}, int64_t{base_value+200}, int64_t{base_value+300}},
    });
    wb->save(path);
}

int main() {
    // ---- テスト用ファイルを準備 ----
    make_test_file("tokyo.xlsx",   "月次", "東京",   1000);
    make_test_file("osaka.xlsx",   "月次", "大阪",   800);
    make_test_file("nagoya.xlsx",  "月次", "名古屋", 600);
    make_test_file("fukuoka.xlsx", "月次", "福岡",   400);
    make_test_file("sapporo.xlsx", "月次", "札幌",   300);
    std::cout << "テストファイルを生成しました\n\n";

    // ---- 共通ページ設定 ----
    PageSetup ps;
    ps.paper_size  = PaperSize::A4;
    ps.orientation = Orientation::Landscape;
    ps.fit_to      = FitTo::Width;    // 全列を 1 ページ幅に収める
    ps.print_gridlines = true;
    ps.header_footer.odd_footer = "&Lページ &P / &N&R&D";

    // ---- BatchPrinter にジョブを登録 ----
    BatchPrinter batch;

    // ファイルごとに異なるヘッダーを設定したい場合は個別にコピーして変更
    auto make_ps = [&](const std::string& title) {
        PageSetup p = ps;
        p.header_footer.odd_header = "&C&14" + title + " 売上レポート";
        return p;
    };

    batch.add("tokyo.xlsx",   "月次", make_ps("東京"),   "東京_月次")
         .add("osaka.xlsx",   "月次", make_ps("大阪"),   "大阪_月次")
         .add("nagoya.xlsx",  "月次", make_ps("名古屋"), "名古屋_月次")
         .add("fukuoka.xlsx", "月次", make_ps("福岡"),   "福岡_月次")
         .add("sapporo.xlsx", "月次", make_ps("札幌"),   "札幌_月次");

    std::cout << "ジョブ登録数: " << batch.job_count() << "\n\n";

    // ---- 進捗コールバック ----
    auto progress = [](size_t done, size_t total,
                       const PrintJob& job, const std::string& msg) {
        std::cout << std::setw(2) << done << "/" << total
                  << " [" << job.label << "] " << msg << "\n";
    };

    // ============================================================
    // ケース 1: 全シートを 1 つの PDF に結合
    // ============================================================
    std::cout << "=== PDF 結合出力 ===\n";
    {
        BatchPrinter b;
        b.add_all({
            PrintJob::make("tokyo.xlsx",   "月次", make_ps("東京"),   "東京_月次"),
            PrintJob::make("osaka.xlsx",   "月次", make_ps("大阪"),   "大阪_月次"),
            PrintJob::make("nagoya.xlsx",  "月次", make_ps("名古屋"), "名古屋_月次"),
            PrintJob::make("fukuoka.xlsx", "月次", make_ps("福岡"),   "福岡_月次"),
            PrintJob::make("sapporo.xlsx", "月次", make_ps("札幌"),   "札幌_月次"),
        });

        auto r = b.to_pdf_merged("all_regions.pdf", progress);

        std::cout << "\n結果: " << r.succeeded << "/" << r.total << " 成功\n";
        std::cout << "処理時間: " << std::fixed << std::setprecision(0)
                  << r.total_elapsed_ms << " ms\n";
        if (!r.output_path.empty())
            std::cout << "出力先: " << r.output_path << "\n";
        if (!r.all_success()) {
            for (auto* f : r.failures())
                std::cerr << "  失敗: " << f->file_path << " - " << f->error_message << "\n";
        }
    }

    // ============================================================
    // ケース 2: ジョブごとに個別 PDF
    // ============================================================
    std::cout << "\n=== 個別 PDF 出力 ===\n";
    {
        BatchPrinter b;
        b.add_all({
            PrintJob::make("tokyo.xlsx",   "月次", make_ps("東京"),   "東京_月次"),
            PrintJob::make("osaka.xlsx",   "月次", make_ps("大阪"),   "大阪_月次"),
            PrintJob::make("nagoya.xlsx",  "月次", make_ps("名古屋"), "名古屋_月次"),
        });

        auto r = b.to_pdf_individual(".", progress);  // カレントディレクトリに出力

        std::cout << "\n結果: " << r.succeeded << "/" << r.total << " 成功\n";
        std::cout << "処理時間: " << r.total_elapsed_ms << " ms\n";
    }

    // ============================================================
    // ケース 3: プリンターに一括送信
    // ============================================================
    std::cout << "\n=== プリンター送信 ===\n";
    {
        // 利用可能なプリンターを確認
        ExcelPrinter ep;
        auto printers = ep.list_printers();
        if (printers.empty()) {
            std::cout << "プリンターが見つかりません（Windows 環境で実行してください）\n";
        } else {
            std::cout << "利用可能なプリンター:\n";
            for (auto& p : printers) std::cout << "  - " << p << "\n";

            BatchPrinter b;
            b.add("tokyo.xlsx",  "月次", make_ps("東京"))
             .add("osaka.xlsx",  "月次", make_ps("大阪"))
             .add("nagoya.xlsx", "月次", make_ps("名古屋"));

            // デフォルトプリンターに送信（printer_name 省略）
            auto r = b.print_all({}, progress);

            std::cout << "\n結果: " << r.succeeded << "/" << r.total << " 成功\n";
            std::cout << "処理時間: " << r.total_elapsed_ms << " ms\n";
        }
    }

    return 0;
}
