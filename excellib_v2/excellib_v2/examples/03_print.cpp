/**
 * examples/03_print.cpp
 *
 * 単一ファイルの印刷・PDF 出力サンプル。
 * PageSetup でページ設定を組み立て、3 つの方法 (設定の書き込み・PDF・プリンター) を示す。
 *
 * 実行:
 *   ./03_print
 *   → report.xlsx と report.pdf が生成される (PDF は Windows + Excel が必要)
 */
// excellib.hpp: WorkbookFactory, ExcelPrinter, PageSetup など全クラスを一括インクルード
#include "excellib/excellib.hpp"
// std::cout: 標準出力
#include <iostream>

// excellib 名前空間を使用する宣言
using namespace excellib;

int main() {
    // ============================================================
    //  テスト用 XLSX ファイルを作成する
    // ============================================================
    {
        // スコープ内で作成して、ブロックを抜けたら wb が自動解放される
        auto wb = WorkbookFactory::create();
        auto& sh = wb->add_sheet("売上レポート");
        // set_table() で売上データをまとめて書き込む
        sh.set_table("A1", {
            {std::string{"製品"},  std::string{"Q1"},      std::string{"Q2"},      std::string{"合計"}},
            {std::string{"A"},     int64_t{1200000},    int64_t{1350000},    int64_t{2550000}},
            {std::string{"B"},     int64_t{ 980000},    int64_t{1100000},    int64_t{2080000}},
            {std::string{"C"},     int64_t{ 760000},    int64_t{ 820000},    int64_t{1580000}},
            {std::string{"合計"},  int64_t{2940000},    int64_t{3270000},    int64_t{6210000}},
        });
        // report.xlsx として保存する
        wb->save("report.xlsx");
    }

    // ============================================================
    //  ExcelPrinter の初期化
    // ============================================================
    // ExcelPrinter: ページ設定の適用・PDF 出力・プリンター送信を担うクラス
    // コンストラクタで Windows の COM (Component Object Model) を初期化する
    // 非 Windows では COM 初期化はスキップされる
    ExcelPrinter printer;

    // ============================================================
    //  PageSetup でページ設定を組み立てる
    // ============================================================
    PageSetup ps;
    // paper_size: 用紙サイズ (PaperSize::A4 = 日本でよく使われる A4 サイズ)
    ps.paper_size   = PaperSize::A4;
    // orientation: 印刷方向 (Landscape = 横向き、Portrait = 縦向き)
    ps.orientation  = Orientation::Landscape;
    // fit_to: 印刷サイズの自動調整
    //   FitTo::Width = 横方向を 1 ページ幅に収める (縦は自動)
    //   FitTo::WidthAndHeight = 縦横両方を指定ページ数に収める
    //   FitTo::None = 自動縮小なし (scale_percent で倍率指定)
    ps.fit_to       = FitTo::Width;
    // print_area: 印刷範囲。"A1:D5" のセル範囲だけ印刷する
    ps.print_area   = PrintArea::from_range("A1:D5");
    // margins: 余白 (インチ単位)。左右を 0.5 インチに設定する
    ps.margins.left = ps.margins.right = 0.5;

    // ---- ヘッダー・フッターの設定 ----
    // Excel のヘッダー/フッターは特殊コードで書式を指定する:
    //   &C  = 中央寄せ
    //   &L  = 左寄せ
    //   &R  = 右寄せ
    //   &14 = フォントサイズ 14pt
    //   &P  = 現在のページ番号
    //   &N  = 総ページ数
    //   &D  = 今日の日付
    //   &"Arial,Bold" = フォント (Arial, 太字) を指定
    ps.header_footer.odd_header = "&C&\"Arial,Bold\"&14売上レポート";  // 中央に太字 14pt で "売上レポート"
    ps.header_footer.odd_footer = "&Lページ &P / &N&R&D";             // 左にページ番号、右に日付

    // ---- タイトル行の繰り返し ----
    // RepeatTitles: 複数ページにまたがる場合に特定の行・列を全ページに繰り返す
    RepeatTitles rt;
    rt.row_start = 0; rt.row_end = 0;   // 0 行目 (= 1 行目) をヘッダーとして全ページに印刷
    ps.repeat_titles = rt;

    // ============================================================
    //  方法 A: ページ設定をファイルに直接書き込む
    // ============================================================
    // apply_page_setup(): XLSX の内部 XML を直接書き換えてページ設定を保存する
    // 引数: (ファイルパス, シート名, PageSetup)
    // Excel で開いても設定が維持される。印刷プレビューで確認してから印刷する用途に向く
    // ★ Windows 以外でも動作する (COM 不要)
    printer.apply_page_setup("report.xlsx", "売上レポート", ps);
    std::cout << "ページ設定を report.xlsx に書き込みました\n";

    // ============================================================
    //  方法 B: PDF として出力する (Windows + Excel が必要)
    // ============================================================
    // to_pdf(): Excel COM を使ってファイルを PDF に変換する
    // 引数:
    //   1. 入力ファイルパス
    //   2. PdfOptions (output_path が必須)
    //   3. シート名 (空文字ならアクティブシート)
    //   4. PageSetup* (nullptr なら apply_page_setup で書き込んだ設定を使用)
    //   5. PrintProgressCallback (進捗メッセージを受け取るラムダ、省略可)
    auto r = printer.to_pdf(
        "report.xlsx",
        {.output_path = "report.pdf"},  // C++20 指示初期化子でフィールド名を指定
        "売上レポート",
        nullptr,    // nullptr = ファイルに書き込まれた設定をそのまま使う
        // 進捗コールバック: VBScript の実行ログを受け取る
        [](const std::string& msg) { std::cout << "[PDF] " << msg << "\n"; }
    );
    if (r.success)
        // output_path: PDF の絶対パスが入っている
        std::cout << "PDF を生成: " << r.output_path << "\n";
    else
        // error_message: Excel COM のエラー説明が入っている
        std::cout << "PDF 失敗: " << r.error_message << "\n";

    // ============================================================
    //  方法 C: プリンターに直接送信する (Windows + Excel が必要)
    // ============================================================
    // default_printer(): デフォルトプリンター名を返す (非 Windows は空文字列)
    std::string def_printer = printer.default_printer();
    std::cout << "デフォルトプリンター: " << def_printer << "\n";

    // list_printers(): 使用可能な全プリンターの名前リストを返す
    for (auto& p : printer.list_printers())
        std::cout << "  - " << p << "\n";

    // print(): Excel COM を使ってプリンターに送信する
    // 引数:
    //   1. 入力ファイルパス
    //   2. PrinterOptions (printer_name: 空ならデフォルトプリンター)
    //   3. シート名
    //   4. &ps: PageSetup を直接指定する (ファイルを書き換えない)
    auto pr = printer.print(
        "report.xlsx",
        {.printer_name = def_printer},  // 空文字でもデフォルトプリンターに送られる
        "売上レポート",
        &ps   // ページ設定をその場で指定 (ファイルを書き換えずに印刷する)
    );
    if (pr.success)
        std::cout << "印刷を送信しました\n";
    else
        std::cout << "印刷失敗: " << pr.error_message << "\n";

    return 0;
}
