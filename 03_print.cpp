/**
 * examples/03_print.cpp
 *
 * 単一ファイルの印刷・PDF 出力サンプル。
 *
 * 実行:
 *   ./03_print
 */

// excellib の全機能をインクルード
#include "excellib/excellib.hpp"
// std::cout など標準入出力を使うため
#include <iostream>

// excellib:: を省略できるようにする
using namespace excellib;

int main() {
    // ---- テスト用 XLSX 作成 ----
    // {} で囲まれたブロックスコープ: ここで作った変数はブロック終了時に破棄される
    // wb を早めに破棄してファイルへの書き込みを確実に終わらせるための工夫
    {
        // 新規ワークブックを作成 (引数なし → デフォルトは XLSX)
        auto wb = WorkbookFactory::create();
        // "売上レポート" という名前のシートを追加し参照を取得
        auto& sh = wb->add_sheet("売上レポート");
        // set_table() でヘッダー + データを A1 から一括書き込み
        sh.set_table("A1", {
            // ヘッダー行
            {std::string{"製品"},  std::string{"Q1"},      std::string{"Q2"},      std::string{"合計"}},
            // データ行: int64_t で整数値を指定
            {std::string{"A"},     int64_t{1200000},    int64_t{1350000},    int64_t{2550000}},
            {std::string{"B"},     int64_t{ 980000},    int64_t{1100000},    int64_t{2080000}},
            {std::string{"C"},     int64_t{ 760000},    int64_t{ 820000},    int64_t{1580000}},
            {std::string{"合計"},  int64_t{2940000},    int64_t{3270000},    int64_t{6210000}},
        });
        // report.xlsx として保存
        wb->save("report.xlsx");
    // ブロックを抜けると wb が破棄され、ファイルへの書き込みが完了する
    }

    // ExcelPrinter オブジェクトを作成
    // このオブジェクトが PDF 出力・印刷・ページ設定の機能を提供する
    ExcelPrinter printer;

    // ---- ページ設定 ----
    // PageSetup 構造体にページの各種設定を詰め込む
    PageSetup ps;
    // 用紙サイズを A4 に設定 (PaperSize::A4 は内部的に数値 9 に対応)
    ps.paper_size   = PaperSize::A4;
    // 印刷の向きを横向きに設定
    ps.orientation  = Orientation::Landscape;    // 横向き
    // 全列を 1 ページ幅に収めるよう自動縮小
    ps.fit_to       = FitTo::Width;              // 全列を 1 ページ幅に収める
    // 印刷範囲を A1:D5 に限定 (それ以外は印刷されない)
    ps.print_area   = PrintArea::from_range("A1:D5");
    // 左右の余白を 0.5 インチに設定
    // ps.margins.left と ps.margins.right に同じ値を一度に代入
    ps.margins.left = ps.margins.right = 0.5;

    // ヘッダー・フッター
    // Excel のヘッダー/フッターコード説明:
    // &C=中央配置  &14=フォントサイズ14pt  &P=現在のページ番号  &N=総ページ数  &D=今日の日付
    // &L=左配置  &R=右配置
    // ヘッダー中央に "売上レポート" を太字・14pt で表示
    ps.header_footer.odd_header = "&C&\"Arial,Bold\"&14売上レポート";
    // フッター: 左にページ番号、右に日付
    ps.header_footer.odd_footer = "&Lページ &P / &N&R&D";

    // 1 行目を全ページに繰り返す (ヘッダー行を各ページの先頭に印刷する)
    RepeatTitles rt;
    // 0-indexed なので row_start=0, row_end=0 は Excel の「1行目」を意味する
    rt.row_start = 0; rt.row_end = 0;
    // PageSetup に RepeatTitles を設定
    ps.repeat_titles = rt;

    // ---- Option A: ページ設定をファイルに焼き込む ----
    // Excel で開いても設定が残る。印刷プレビューで確認してから印刷する用途に向く
    // apply_page_setup() は XLSX の ZIP 内の XML を直接書き換える (Excel 不要)
    printer.apply_page_setup("report.xlsx", "売上レポート", ps);
    std::cout << "ページ設定を report.xlsx に書き込みました\n";

    // ---- Option B: PDF 出力 ----
    // to_pdf() は Excel COM オートメーション経由で PDF を生成する (Windows のみ)
    // 第2引数の {.output_path = "report.pdf"} は C++20 の指示初期化子構文
    // 第4引数の nullptr は「ファイルに焼き込まれた設定を使用」を意味する
    // 第5引数はラムダ関数 (進捗メッセージを受け取るコールバック)
    auto r = printer.to_pdf(
        "report.xlsx",
        {.output_path = "report.pdf"},
        "売上レポート",
        nullptr,    // nullptr = ファイルに焼き込まれた設定を使用
        [](const std::string& msg) { std::cout << "[PDF] " << msg << "\n"; }
    );
    // r.success が true なら PDF 生成成功
    if (r.success)
        std::cout << "PDF を生成: " << r.output_path << "\n";
    else
        // r.error_message にエラーの詳細が格納されている
        std::cout << "PDF 失敗: " << r.error_message << "\n";

    // ---- Option C: プリンターに直接送信 ----
    // デフォルトプリンターを確認
    // default_printer() はシステムのデフォルトプリンター名を文字列で返す
    std::string def_printer = printer.default_printer();
    std::cout << "デフォルトプリンター: " << def_printer << "\n";

    // プリンター一覧
    // list_printers() は利用可能なプリンターの名前一覧を vector<string> で返す
    // auto& p でループ変数を参照として受け取り (コピーを避ける)
    for (auto& p : printer.list_printers())
        std::cout << "  - " << p << "\n";

    // 印刷（PageSetup を直接渡す方法）
    // print() は Excel COM 経由でプリンターに送信する (Windows のみ)
    auto pr = printer.print(
        "report.xlsx",
        {.printer_name = def_printer},   // 空文字でもデフォルトプリンターに送られる
        "売上レポート",
        &ps   // ページ設定をその場で指定（ファイルを書き換えない）
        // &ps はポインタ渡し。apply_page_setup() と違い、ファイル自体は変更されない
    );
    // pr.success が true なら印刷ジョブの送信に成功
    if (pr.success)
        std::cout << "印刷を送信しました\n";
    else
        std::cout << "印刷失敗: " << pr.error_message << "\n";

    // 正常終了
    return 0;
}
