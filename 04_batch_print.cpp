/**
 * examples/04_batch_print.cpp
 *
 * バッチ印刷サンプル。
 * 複数ファイル・複数シートを Excel 1 起動で高速処理する。
 *
 * 実行:
 *   ./04_batch_print
 */

// excellib の全機能をインクルード
#include "excellib/excellib.hpp"
// std::cout など標準入出力を使うため
#include <iostream>
// std::setw, std::fixed, std::setprecision など出力フォーマット用
#include <iomanip>

// excellib:: を省略できるようにする
using namespace excellib;

// テスト用 XLSX を作成するヘルパー関数
// static はこのファイル内だけで使える関数にするキーワード (外部リンクを持たない)
// const std::string& は文字列をコピーせず参照で受け取る (効率的)
// int base_value はデータの基準値 (各製品の売上の基底となる数値)
static void make_test_file(const std::string& path,
                            const std::string& sheet,
                            const std::string& title,
                            int base_value) {
    // 新規ワークブックを作成
    auto wb = WorkbookFactory::create();
    // 指定した名前のシートを追加し参照を取得
    auto& sh = wb->add_sheet(sheet);
    // A1 起点に 2次元データを一括書き込み
    sh.set_table("A1", {
        // ヘッダー行: タイトル文字列と Q1〜Q3 の列名
        {std::string{title}, std::string{"Q1"}, std::string{"Q2"}, std::string{"Q3"}},
        // 製品A: base_value を起点に 100 ずつ増える値
        {std::string{"製品A"}, int64_t{base_value},     int64_t{base_value+100}, int64_t{base_value+200}},
        // 製品B: base_value+50 を起点に
        {std::string{"製品B"}, int64_t{base_value+50},  int64_t{base_value+150}, int64_t{base_value+250}},
        // 製品C: base_value+100 を起点に
        {std::string{"製品C"}, int64_t{base_value+100}, int64_t{base_value+200}, int64_t{base_value+300}},
    });
    // 指定したパスにファイルを保存
    wb->save(path);
}

int main() {
    // ---- テスト用ファイルを準備 ----
    // 各地域のテスト用 XLSX ファイルを生成
    // "月次" シートに各地域のデータを base_value を変えて作成
    make_test_file("tokyo.xlsx",   "月次", "東京",   1000);
    make_test_file("osaka.xlsx",   "月次", "大阪",   800);
    make_test_file("nagoya.xlsx",  "月次", "名古屋", 600);
    make_test_file("fukuoka.xlsx", "月次", "福岡",   400);
    make_test_file("sapporo.xlsx", "月次", "札幌",   300);
    std::cout << "テストファイルを生成しました\n\n";

    // ---- 共通ページ設定 ----
    // 全ジョブで使い回す PageSetup を作成
    PageSetup ps;
    // 用紙サイズ: A4
    ps.paper_size  = PaperSize::A4;
    // 印刷の向き: 横向き
    ps.orientation = Orientation::Landscape;
    // 全列を 1 ページ幅に収める
    ps.fit_to      = FitTo::Width;    // 全列を 1 ページ幅に収める
    // グリッド線を印刷する
    ps.print_gridlines = true;
    // フッターにページ番号と日付を設定 (&L=左, &P=ページ番号, &N=総数, &R=右, &D=日付)
    ps.header_footer.odd_footer = "&Lページ &P / &N&R&D";

    // ---- BatchPrinter にジョブを登録 ----
    // BatchPrinter は複数の印刷ジョブをまとめて処理するクラス
    BatchPrinter batch;

    // ファイルごとに異なるヘッダーを設定したい場合は個別にコピーして変更
    // make_ps はラムダ関数 (その場で定義できる小さな関数)
    // [&] で外側の ps 変数をキャプチャ (参照で借りる)
    // const std::string& title で地域名を受け取る
    auto make_ps = [&](const std::string& title) {
        // ps をコピーして個別のページ設定を作る (ps そのものは変更しない)
        PageSetup p = ps;
        // ヘッダー中央に "&14"(14pt) + 地域名 + " 売上レポート" を設定
        p.header_footer.odd_header = "&C&14" + title + " 売上レポート";
        // 設定済みのページ設定を返す
        return p;
    };

    // add() でジョブを1件ずつ登録し、戻り値 (BatchPrinter&) を使ってメソッドチェーン
    // 各引数: ファイルパス, シート名, PageSetup, ラベル (進捗表示・PDF名に使われる)
    batch.add("tokyo.xlsx",   "月次", make_ps("東京"),   "東京_月次")
         .add("osaka.xlsx",   "月次", make_ps("大阪"),   "大阪_月次")
         .add("nagoya.xlsx",  "月次", make_ps("名古屋"), "名古屋_月次")
         .add("fukuoka.xlsx", "月次", make_ps("福岡"),   "福岡_月次")
         .add("sapporo.xlsx", "月次", make_ps("札幌"),   "札幌_月次");

    // 登録されたジョブ数を表示
    std::cout << "ジョブ登録数: " << batch.job_count() << "\n\n";

    // ---- 進捗コールバック ----
    // BatchProgressCallback の型に合わせたラムダ関数を定義
    // done: 完了済みジョブ数, total: 全ジョブ数, job: 現在のジョブ, msg: 進捗メッセージ
    auto progress = [](size_t done, size_t total,
                       const PrintJob& job, const std::string& msg) {
        // setw(2) で2桁に揃えて done/total を表示
        std::cout << std::setw(2) << done << "/" << total
                  << " [" << job.label << "] " << msg << "\n";
    };

    // ============================================================
    // ケース 1: 全シートを 1 つの PDF に結合
    // ============================================================
    std::cout << "=== PDF 結合出力 ===\n";
    {
        // 新しい BatchPrinter を作成 (b は「このスコープだけ」の変数)
        BatchPrinter b;
        // add_all() でジョブのベクタを一括登録
        // PrintJob::make() でジョブオブジェクトを生成
        b.add_all({
            PrintJob::make("tokyo.xlsx",   "月次", make_ps("東京"),   "東京_月次"),
            PrintJob::make("osaka.xlsx",   "月次", make_ps("大阪"),   "大阪_月次"),
            PrintJob::make("nagoya.xlsx",  "月次", make_ps("名古屋"), "名古屋_月次"),
            PrintJob::make("fukuoka.xlsx", "月次", make_ps("福岡"),   "福岡_月次"),
            PrintJob::make("sapporo.xlsx", "月次", make_ps("札幌"),   "札幌_月次"),
        });

        // to_pdf_merged() で全ジョブを1つの PDF に結合して出力
        // Excel を1回起動してシートをコピーし、一括で PDF に変換する
        auto r = b.to_pdf_merged("all_regions.pdf", progress);

        // 結果のサマリーを表示
        // r.succeeded: 成功件数, r.total: 全件数
        std::cout << "\n結果: " << r.succeeded << "/" << r.total << " 成功\n";
        // std::fixed で小数点固定表示, setprecision(0) で小数点以下0桁 (整数表示)
        std::cout << "処理時間: " << std::fixed << std::setprecision(0)
                  << r.total_elapsed_ms << " ms\n";
        // output_path が空でなければ出力先ファイルパスを表示
        if (!r.output_path.empty())
            std::cout << "出力先: " << r.output_path << "\n";
        // all_success() が false (=失敗があった) の場合
        if (!r.all_success()) {
            // failures() で失敗したジョブのポインタ一覧を取得
            // auto* f でポインタとして受け取る
            for (auto* f : r.failures())
                // エラーの詳細をファイルパスとエラーメッセージとともに表示
                std::cerr << "  失敗: " << f->file_path << " - " << f->error_message << "\n";
        }
    }

    // ============================================================
    // ケース 2: ジョブごとに個別 PDF
    // ============================================================
    std::cout << "\n=== 個別 PDF 出力 ===\n";
    {
        // 3件だけ登録する例
        BatchPrinter b;
        b.add_all({
            PrintJob::make("tokyo.xlsx",   "月次", make_ps("東京"),   "東京_月次"),
            PrintJob::make("osaka.xlsx",   "月次", make_ps("大阪"),   "大阪_月次"),
            PrintJob::make("nagoya.xlsx",  "月次", make_ps("名古屋"), "名古屋_月次"),
        });

        // to_pdf_individual() でジョブごとに別ファイルの PDF を出力
        // "." はカレントディレクトリを出力先として指定
        // 出力ファイル名は "{ラベル}_{連番}.pdf" の形式になる
        auto r = b.to_pdf_individual(".", progress);  // カレントディレクトリに出力

        std::cout << "\n結果: " << r.succeeded << "/" << r.total << " 成功\n";
        std::cout << "処理時間: " << r.total_elapsed_ms << " ms\n";
    }

    // ============================================================
    // ケース 3: プリンターに一括送信
    // ============================================================
    std::cout << "\n=== プリンター送信 ===\n";
    {
        // ExcelPrinter でプリンター一覧を取得 (Windows 限定)
        ExcelPrinter ep;
        // list_printers() でシステムにインストール済みのプリンター名一覧を取得
        auto printers = ep.list_printers();
        // プリンターが見つからない場合はメッセージを表示してスキップ
        if (printers.empty()) {
            std::cout << "プリンターが見つかりません（Windows 環境で実行してください）\n";
        } else {
            // プリンター一覧を表示
            std::cout << "利用可能なプリンター:\n";
            for (auto& p : printers) std::cout << "  - " << p << "\n";

            // 3件を印刷する例
            BatchPrinter b;
            b.add("tokyo.xlsx",  "月次", make_ps("東京"))
             .add("osaka.xlsx",  "月次", make_ps("大阪"))
             .add("nagoya.xlsx", "月次", make_ps("名古屋"));

            // print_all() でプリンターに一括送信
            // {} は PrinterOptions のデフォルト値 (printer_name が空 → デフォルトプリンター)
            // 省略可能な printer_name を空にすることでシステムの既定プリンターに送られる
            auto r = b.print_all({}, progress);

            std::cout << "\n結果: " << r.succeeded << "/" << r.total << " 成功\n";
            std::cout << "処理時間: " << r.total_elapsed_ms << " ms\n";
        }
    }

    // 正常終了
    return 0;
}
