/**
 * examples/04_batch_print.cpp
 *
 * バッチ印刷サンプル。
 * 複数ファイル・複数シートを Excel 1 起動でまとめて処理する方法を示す。
 * Excel.Application を 1 度だけ起動してから全ジョブを処理するため、
 * 個別に起動するより高速に動作する。
 *
 * 実行:
 *   ./04_batch_print
 */
// excellib.hpp: WorkbookFactory, BatchPrinter, ExcelPrinter など全クラスをインクルード
#include "excellib/excellib.hpp"
// std::cout: 標準出力
#include <iostream>
// std::setw / std::setprecision / std::fixed: 出力書式の指定
#include <iomanip>

// excellib 名前空間を使用する宣言
using namespace excellib;

// ============================================================
//  テスト用 XLSX ファイルを作成するヘルパー関数
// ============================================================
// path: 保存先ファイルパス (例: "tokyo.xlsx")
// sheet: シート名 (例: "月次")
// title: セル A1 に入れるタイトル文字列 (例: "東京")
// base_value: 数値の基準値 (拠点ごとに異なる値を設定)
static void make_test_file(const std::string& path,
                            const std::string& sheet,
                            const std::string& title,
                            int base_value) {
    auto wb = WorkbookFactory::create();
    auto& sh = wb->add_sheet(sheet);
    // set_table() で 4 行 4 列のデータを書き込む
    // base_value から始まる連続した数値を各行に設定する
    sh.set_table("A1", {
        // ヘッダー行: タイトルと四半期名
        {std::string{title}, std::string{"Q1"}, std::string{"Q2"}, std::string{"Q3"}},
        // 製品 A のデータ: base_value, base_value+100, base_value+200
        {std::string{"製品A"}, int64_t{base_value},     int64_t{base_value+100}, int64_t{base_value+200}},
        {std::string{"製品B"}, int64_t{base_value+50},  int64_t{base_value+150}, int64_t{base_value+250}},
        {std::string{"製品C"}, int64_t{base_value+100}, int64_t{base_value+200}, int64_t{base_value+300}},
    });
    wb->save(path);
}

int main() {
    // ============================================================
    //  テスト用ファイルを 5 拠点分作成する
    // ============================================================
    make_test_file("tokyo.xlsx",   "月次", "東京",   1000);
    make_test_file("osaka.xlsx",   "月次", "大阪",   800);
    make_test_file("nagoya.xlsx",  "月次", "名古屋", 600);
    make_test_file("fukuoka.xlsx", "月次", "福岡",   400);
    make_test_file("sapporo.xlsx", "月次", "札幌",   300);
    std::cout << "テストファイルを生成しました\n\n";

    // ============================================================
    //  共通ページ設定を定義する
    // ============================================================
    PageSetup ps;
    ps.paper_size  = PaperSize::A4;                // A4 用紙
    ps.orientation = Orientation::Landscape;        // 横向き
    ps.fit_to      = FitTo::Width;                 // 全列を 1 ページ幅に収める
    ps.print_gridlines = true;                      // グリッド線を印刷する
    // フッター: 左端にページ番号、右端に日付
    ps.header_footer.odd_footer = "&Lページ &P / &N&R&D";

    // ============================================================
    //  BatchPrinter にジョブを登録する
    // ============================================================
    // BatchPrinter: 複数のジョブを管理して一括処理するクラス
    BatchPrinter batch;

    // 拠点ごとにヘッダーを変えた PageSetup を生成するラムダ
    // ps (共通設定) をコピーしてヘッダーだけ上書きする
    auto make_ps = [&](const std::string& title) {
        PageSetup p = ps;   // ps のコピーを作る (全フィールドがコピーされる)
        // ヘッダー中央に "東京 売上レポート" のようなタイトルを設定する
        // &C = 中央寄せ、&14 = 14pt フォントサイズ
        p.header_footer.odd_header = "&C&14" + title + " 売上レポート";
        return p;
    };

    // add() のメソッドチェーン: 引数は (ファイルパス, シート名, PageSetup, ラベル)
    // ラベルは個別 PDF のファイル名に使われる
    batch.add("tokyo.xlsx",   "月次", make_ps("東京"),   "東京_月次")
         .add("osaka.xlsx",   "月次", make_ps("大阪"),   "大阪_月次")
         .add("nagoya.xlsx",  "月次", make_ps("名古屋"), "名古屋_月次")
         .add("fukuoka.xlsx", "月次", make_ps("福岡"),   "福岡_月次")
         .add("sapporo.xlsx", "月次", make_ps("札幌"),   "札幌_月次");

    // job_count(): 登録されたジョブ数を返す
    std::cout << "ジョブ登録数: " << batch.job_count() << "\n\n";

    // ============================================================
    //  進捗コールバックを定義する
    // ============================================================
    // BatchProgressCallback: (完了数, 総数, 現在のジョブ, メッセージ) を受け取るラムダ
    auto progress = [](size_t done, size_t total,
                       const PrintJob& job, const std::string& msg) {
        // 進捗を "2/5 [東京_月次] Starting..." のように出力する
        std::cout << std::setw(2) << done << "/" << total
                  << " [" << job.label << "] " << msg << "\n";
    };

    // ============================================================
    //  ケース 1: 全シートを 1 つの PDF にまとめて出力する
    // ============================================================
    std::cout << "=== PDF 結合出力 ===\n";
    {
        BatchPrinter b;   // 新しい BatchPrinter を作成 (ジョブリストは空)
        // add_all(): PrintJob のベクターからまとめてジョブを追加する
        // PrintJob::make(ファイルパス, シート名, PageSetup, ラベル): PrintJob を生成する
        b.add_all({
            PrintJob::make("tokyo.xlsx",   "月次", make_ps("東京"),   "東京_月次"),
            PrintJob::make("osaka.xlsx",   "月次", make_ps("大阪"),   "大阪_月次"),
            PrintJob::make("nagoya.xlsx",  "月次", make_ps("名古屋"), "名古屋_月次"),
            PrintJob::make("fukuoka.xlsx", "月次", make_ps("福岡"),   "福岡_月次"),
            PrintJob::make("sapporo.xlsx", "月次", make_ps("札幌"),   "札幌_月次"),
        });

        // to_pdf_merged(): 全ジョブのシートを 1 つの作業ブックにコピーしてから
        // 一括で PDF に変換する (Excel が 1 度だけ起動される)
        // 引数: (出力 PDF ファイルパス, プログレスコールバック)
        auto r = b.to_pdf_merged("all_regions.pdf", progress);

        // BatchResult: 全ジョブの処理結果をまとめた構造体
        std::cout << "\n結果: " << r.succeeded << "/" << r.total << " 成功\n";
        // total_elapsed_ms: 全ジョブの合計処理時間 (ミリ秒)
        std::cout << "処理時間: " << std::fixed << std::setprecision(0)
                  << r.total_elapsed_ms << " ms\n";
        // output_path: to_pdf_merged で指定した出力先パス
        if (!r.output_path.empty())
            std::cout << "出力先: " << r.output_path << "\n";
        // all_success(): 全ジョブが成功なら true
        if (!r.all_success()) {
            // failures(): 失敗したジョブへのポインターのリストを返す
            for (auto* f : r.failures())
                std::cerr << "  失敗: " << f->file_path << " - " << f->error_message << "\n";
        }
    }

    // ============================================================
    //  ケース 2: ジョブごとに個別の PDF ファイルを出力する
    // ============================================================
    std::cout << "\n=== 個別 PDF 出力 ===\n";
    {
        BatchPrinter b;
        b.add_all({
            PrintJob::make("tokyo.xlsx",   "月次", make_ps("東京"),   "東京_月次"),
            PrintJob::make("osaka.xlsx",   "月次", make_ps("大阪"),   "大阪_月次"),
            PrintJob::make("nagoya.xlsx",  "月次", make_ps("名古屋"), "名古屋_月次"),
        });

        // to_pdf_individual(): 各ジョブを個別の PDF ファイルとして書き出す
        // 引数: (出力ディレクトリ, プログレスコールバック)
        // 出力ファイル名は "[ラベル]_[連番].pdf" の形式で自動生成される
        // 例: "東京_月次_1.pdf", "大阪_月次_2.pdf", ...
        auto r = b.to_pdf_individual(".", progress);  // "." = カレントディレクトリ

        std::cout << "\n結果: " << r.succeeded << "/" << r.total << " 成功\n";
        std::cout << "処理時間: " << r.total_elapsed_ms << " ms\n";
    }

    // ============================================================
    //  ケース 3: プリンターに一括送信する
    // ============================================================
    std::cout << "\n=== プリンター送信 ===\n";
    {
        // ExcelPrinter: プリンター一覧の取得に使う
        ExcelPrinter ep;
        auto printers = ep.list_printers();   // 利用可能なプリンターのリスト
        if (printers.empty()) {
            // list_printers() が空を返すのは非 Windows 環境か、プリンターがない場合
            std::cout << "プリンターが見つかりません（Windows 環境で実行してください）\n";
        } else {
            std::cout << "利用可能なプリンター:\n";
            for (auto& p : printers) std::cout << "  - " << p << "\n";

            BatchPrinter b;
            b.add("tokyo.xlsx",  "月次", make_ps("東京"))
             .add("osaka.xlsx",  "月次", make_ps("大阪"))
             .add("nagoya.xlsx", "月次", make_ps("名古屋"));

            // print_all(): 全ジョブを 1 つの Excel インスタンスでまとめてプリンターに送信する
            // PrinterOptions: printer_name (空ならデフォルトプリンター), collate (部単位印刷) を指定
            // {} は PrinterOptions のデフォルト値 (= デフォルトプリンター)
            auto r = b.print_all({}, progress);

            std::cout << "\n結果: " << r.succeeded << "/" << r.total << " 成功\n";
            std::cout << "処理時間: " << r.total_elapsed_ms << " ms\n";
        }
    }

    return 0;
}
