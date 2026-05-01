// ExcelPrinter の実装ファイル
// Windows の Excel COM (Component Object Model) オートメーションを使って
// Excel ファイルを PDF に変換したり、プリンターに送信したりする機能を提供する
// COM は Windows 固有の仕組みのため、#ifdef _WIN32 で非 Windows ビルドを保護している
#include "excellib/excel_printer.hpp"
// std::ostringstream: 文字列を組み立てるストリーム (VBScript 生成に使う)
#include <sstream>
// std::ofstream / std::ifstream: ファイルの読み書き
#include <fstream>
// std::memcmp / std::memset: バイト列操作
#include <cstring>
// std::system(): 外部コマンド実行 (cscript でVBScriptを起動)
#include <cstdlib>

// Windows 専用ヘッダー群 (非 Windows では無視される)
#ifdef _WIN32
#define WIN32_LEAN_AND_MEAN             // windows.h から不要な部分を除外する
#include <windows.h>                    // Windows API の基本ヘッダー
#include <objbase.h>                    // COM の基盤 API (CoInitializeEx など)
#pragma comment(lib,"ole32.lib")       // COM ライブラリをリンクする
#pragma comment(lib,"oleaut32.lib")    // OLE オートメーションライブラリをリンクする
#pragma comment(lib,"winspool.lib")    // プリンタースプーラー API をリンクする
#endif

namespace excellib {

// ============================================================
//  COM 初期化ガード (RAII パターン)
// ============================================================
// COMGuard: COM の初期化と終了処理を RAII で管理する内部構造体
// RAII (Resource Acquisition Is Initialization):
//   コンストラクタで COM を初期化し、デストラクタで解放する
//   例外が発生しても確実にクリーンアップされる
struct ExcelPrinter::COMGuard {
#ifdef _WIN32
    bool ok{false};   // COM 初期化が成功したかどうかのフラグ
    // コンストラクタ: CoInitializeEx で COM をアパートメントスレッドモードで初期化する
    // COINIT_APARTMENTTHREADED: STA (Single-Threaded Apartment) モード
    // Excel など多くの COM オブジェクトは STA が必要
    COMGuard()  { ok = SUCCEEDED(CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED)); }
    // デストラクタ: 初期化が成功していれば CoUninitialize で COM を解放する
    ~COMGuard() { if (ok) CoUninitialize(); }
#endif
};

// コンストラクタ: COMGuard をヒープに生成して unique_ptr で管理する
// これにより ExcelPrinter が破棄されると COMGuard も自動的に破棄される
ExcelPrinter::ExcelPrinter() : com_(std::make_unique<COMGuard>()) {}
// デストラクタ: unique_ptr が自動的に COMGuard を解放するため default で十分
ExcelPrinter::~ExcelPrinter() = default;

// log(): プログレスコールバックが設定されていれば呼び出す
void ExcelPrinter::log(const std::string& msg) {
    if (progress_cb_) progress_cb_(msg);
}

// ============================================================
//  プリンター一覧の取得
// ============================================================
// list_printers(): インストール済みプリンターの名前一覧を返す
// Windows API の EnumPrintersA を使ってローカルおよびネットワークプリンターを列挙する
std::vector<std::string> ExcelPrinter::list_printers() const {
    std::vector<std::string> result;
#ifdef _WIN32
    DWORD need=0, cnt=0;
    // 1 回目の呼び出し: バッファサイズを確認する (buf=nullptr で必要サイズが need に入る)
    EnumPrintersA(PRINTER_ENUM_LOCAL|PRINTER_ENUM_CONNECTIONS,nullptr,4,nullptr,0,&need,&cnt);
    std::vector<uint8_t> buf(need);   // 必要なサイズのバッファを確保する
    // 2 回目の呼び出し: 実際にプリンター情報を取得する
    if (EnumPrintersA(PRINTER_ENUM_LOCAL|PRINTER_ENUM_CONNECTIONS,nullptr,4,
                      buf.data(),need,&need,&cnt)) {
        // バッファを PRINTER_INFO_4A 構造体の配列としてキャストする
        auto* info = reinterpret_cast<PRINTER_INFO_4A*>(buf.data());
        for (DWORD i=0; i<cnt; ++i)
            // プリンター名が有効なら (NULL でなければ) リストに追加する
            if (info[i].pPrinterName) result.push_back(info[i].pPrinterName);
    }
#endif
    return result;
}

// default_printer(): デフォルトプリンターの名前を返す
std::string ExcelPrinter::default_printer() const {
#ifdef _WIN32
    char buf[256]={}; DWORD sz=sizeof(buf);
    // GetDefaultPrinterA: デフォルトプリンター名を buf に格納する
    GetDefaultPrinterA(buf,&sz); return buf;
#else
    return "";   // 非 Windows は空文字列を返す
#endif
}

// ============================================================
//  PageSetup 構造体を VBScript コードに変換する
// ============================================================
// ps_to_vbs(): PageSetup の設定値を Excel VBA の PageSetup プロパティへの代入文として出力する
// ws: VBScript 内でシートを参照する変数名 (例: "ws")
// ps: ページ設定 (用紙サイズ, 向き, 余白, ヘッダー/フッター etc.)
// 戻り値: VBScript コードの断片 (各行が一つのプロパティ設定)
static std::string ps_to_vbs(const std::string& ws, const PageSetup& ps) {
    std::ostringstream o;
    // 用紙サイズを Excel の PaperSize 定数 (uint16_t) として設定
    // 例: A4 = 9, Letter = 1
    o << ws << ".PageSetup.PaperSize = " << uint16_t(ps.paper_size) << "\n";
    // 印刷方向: Landscape=2, Portrait=1 (Excel VBA の定数)
    o << ws << ".PageSetup.Orientation = "
      << (ps.orientation==Orientation::Landscape ? "2" : "1") << "\n";

    // FitTo (印刷サイズの自動調整) の設定
    if (ps.fit_to != FitTo::None) {
        // Zoom=False でズームを無効化し、FitToPages を有効にする
        o << ws << ".PageSetup.Zoom = False\n";
        // 幅・高さそれぞれの「ページ数に収める」設定
        bool fw = ps.fit_to==FitTo::Width  || ps.fit_to==FitTo::WidthAndHeight;
        bool fh = ps.fit_to==FitTo::Height || ps.fit_to==FitTo::WidthAndHeight;
        // fit_to==Width の場合のみ幅ページ数を設定、それ以外は False (無制限)
        o << ws << ".PageSetup.FitToPagesWide = "
          << (fw ? std::to_string(ps.fit_to_pages_wide) : "False") << "\n";
        o << ws << ".PageSetup.FitToPagesTall = "
          << (fh ? std::to_string(ps.fit_to_pages_tall) : "False") << "\n";
    } else {
        // FitTo なし: 倍率 (%) で指定する (100 = 等倍)
        o << ws << ".PageSetup.Zoom = " << ps.scale_percent << "\n";
    }

    // 余白: Excel は「ポイント」単位で受け取るが、InchesToPoints() で変換する
    // InchesToPoints() は 1 インチ = 72 ポイントとして計算する
    auto& m = ps.margins;
    o << ws << ".PageSetup.LeftMargin   = Application.InchesToPoints(" << m.left   << ")\n"
      << ws << ".PageSetup.RightMargin  = Application.InchesToPoints(" << m.right  << ")\n"
      << ws << ".PageSetup.TopMargin    = Application.InchesToPoints(" << m.top    << ")\n"
      << ws << ".PageSetup.BottomMargin = Application.InchesToPoints(" << m.bottom << ")\n"
      << ws << ".PageSetup.HeaderMargin = Application.InchesToPoints(" << m.header << ")\n"
      << ws << ".PageSetup.FooterMargin = Application.InchesToPoints(" << m.footer << ")\n";

    // ヘッダー・フッターの設定 (空でない場合のみ設定する)
    auto& hf = ps.header_footer;
    if (!hf.odd_header.empty())
        // CenterHeader: 奇数ページのヘッダー中央部
        o << ws << ".PageSetup.CenterHeader = \"" << hf.odd_header << "\"\n";
    if (!hf.odd_footer.empty())
        // CenterFooter: 奇数ページのフッター中央部
        o << ws << ".PageSetup.CenterFooter = \"" << hf.odd_footer << "\"\n";
    if (!hf.even_header.empty())
        // EvenHeader: 偶数ページのヘッダー
        o << ws << ".PageSetup.EvenHeader = \"" << hf.even_header << "\"\n";
    if (!hf.even_footer.empty())
        // EvenFooter: 偶数ページのフッター
        o << ws << ".PageSetup.EvenFooter = \"" << hf.even_footer << "\"\n";

    // その他のページ設定オプション
    o << ws << ".PageSetup.PrintGridlines = "
      << (ps.print_gridlines ? "True" : "False") << "\n"     // グリッド線の印刷
      << ws << ".PageSetup.BlackAndWhite = "
      << (ps.black_and_white ? "True" : "False") << "\n"     // 白黒印刷
      << ws << ".PageSetup.Draft = "
      << (ps.draft_quality ? "True" : "False") << "\n"       // 下書き品質
      << ws << ".PageSetup.PrintHeadings = "
      << (ps.print_row_col_headings ? "True" : "False") << "\n"; // 行・列見出しの印刷

    // 印刷範囲の設定 (例: "A1:H30")
    if (ps.print_area)
        o << ws << ".PageSetup.PrintArea = \"" << ps.print_area->to_range() << "\"\n";

    // タイトル行・列の繰り返し設定 (各ページの先頭に印刷する行・列)
    if (ps.repeat_titles) {
        auto& rt = *ps.repeat_titles;
        // 繰り返し行の設定: 例 "$1:$3" (1〜3 行目を全ページに印刷)
        if (rt.row_start && rt.row_end)
            o << ws << ".PageSetup.PrintTitleRows = \"$"
              << (*rt.row_start+1) << ":$" << (*rt.row_end+1) << "\"\n";
        // 繰り返し列の設定: 例 "$A:$C" (A〜C 列目を全ページに印刷)
        if (rt.col_start && rt.col_end) {
            // 列番号 (0-indexed) を列文字 (A, B, ... Z, AA, ...) に変換するラムダ
            auto col_str = [](uint32_t c) {
                std::string r; ++c;   // 1-indexed に変換
                while (c>0) { --c; r+=char('A'+c%26); c/=26; }
                std::reverse(r.begin(),r.end()); return r;
            };
            o << ws << ".PageSetup.PrintTitleColumns = \"$"
              << col_str(*rt.col_start) << ":$" << col_str(*rt.col_end) << "\"\n";
        }
    }

    // 先頭ページ番号の設定 (0 = デフォルト、1 以上なら指定値から開始)
    if (ps.first_page_number > 0)
        o << ws << ".PageSetup.FirstPageNumber = " << ps.first_page_number << "\n";

    return o.str();
}

// ============================================================
//  VBScript 実行ヘルパー
// ============================================================
// run_vbs(): VBScript コードを一時ファイルに書き出し cscript.exe で実行する
// script: 実行する VBScript コードの文字列
// err_out: エラーがあれば内容が書き出される (出力引数)
// 戻り値: 成功なら true、失敗なら false
static bool run_vbs(const std::string& script, std::string& err_out) {
#ifdef _WIN32
    // 一時ファイルのパスを格納する配列
    char tmp[MAX_PATH], vbs[MAX_PATH+4], log[MAX_PATH+4];
    // GetTempPathA: 一時フォルダのパスを取得する (例: C:\Users\...\AppData\Local\Temp\)
    GetTempPathA(MAX_PATH, tmp);
    // GetTempFileNameA: 一時ファイルの名前を生成する ("exc" がプレフィックス)
    GetTempFileNameA(tmp, "exc", 0, vbs);
    // log ファイルのパスを作成 (VBS ファイル名 + ".log")
    strcpy(log, vbs); strcat(log, ".log"); strcat(vbs, ".vbs");

    // VBScript にエラーハンドリングを追加するラッパーを作成する
    // "On Error Resume Next": VBScript でエラーが発生しても処理を継続する
    // エラーが発生すると Err.Number に非 0 が入るので、それをログファイルに書き出す
    std::string full =
        "On Error Resume Next\n" + script + "\n"
        "If Err.Number <> 0 Then\n"
        "  Dim fso,f\n"
        // FileSystemObject を使ってログファイルを作成
        "  Set fso=CreateObject(\"Scripting.FileSystemObject\")\n"
        "  Set f=fso.CreateTextFile(\"" + std::string(log) + "\",True)\n"
        "  f.WriteLine Err.Number & \": \" & Err.Description\n"
        "  f.Close\n"
        "End If\n";

    // VBScript ファイルに書き出す
    { std::ofstream f(vbs); f << full; }
    // cscript //nologo: ロゴ表示なしで VBScript を実行する
    int ret = std::system(("cscript //nologo \"" + std::string(vbs) + "\"").c_str());
    // ログファイルを読んでエラー内容を取得する
    { std::ifstream f(log); if(f) std::getline(f, err_out); }
    // 一時ファイルを削除する (クリーンアップ)
    DeleteFileA(vbs); DeleteFileA(log);
    // 終了コード 0 かつエラーログが空なら成功
    return ret==0 && err_out.empty();
#else
    // 非 Windows では実行不可
    err_out = "Windows only";
    return false;
#endif
}

// abs_path(): 相対パスを絶対パスに変換するヘルパー
// VBScript に渡すパスは絶対パスでないと Excel が見つけられないため使用する
static std::string abs_path(const std::string& p) {
#ifdef _WIN32
    char buf[MAX_PATH];
    // GetFullPathNameA: 相対パスを絶対パスに変換する
    return GetFullPathNameA(p.c_str(),MAX_PATH,buf,nullptr) ? buf : p;
#else
    return p;
#endif
}

// ============================================================
//  Excel COM: PDF 出力
// ============================================================
// excel_to_pdf(): Excel COM を使って XLSX/XLS ファイルを PDF に変換する
// path: 入力ファイルパス, opts: PDF オプション (出力パスなど)
// sheet: 対象シート名 (空なら ActiveSheet)
// ps: 適用するページ設定 (nullptr なら設定変更なし)
PrintResult ExcelPrinter::excel_to_pdf(const std::string& path,
                                        const PdfOptions& opts,
                                        const std::string& sheet,
                                        const PageSetup* ps) {
    // 入力ファイルと出力ファイルを絶対パスに変換
    std::string ai = abs_path(path);
    std::string ao = abs_path(opts.output_path);

    std::ostringstream vbs;
    // Excel.Application を COM で生成して非表示で起動する
    vbs << "Dim xl,wb,ws\n"
        << "Set xl=CreateObject(\"Excel.Application\")\n"
        // Visible=False: Excel ウィンドウを非表示にする
        << "xl.Visible=False\n"
        // DisplayAlerts=False: 警告ダイアログを非表示にする
        << "xl.DisplayAlerts=False\n"
        // Workbooks.Open: ファイルを開く
        << "Set wb=xl.Workbooks.Open(\"" << ai << "\")\n";

    // 対象シートを選択する
    if (!sheet.empty())
        // 指定シート名でシートを取得してアクティブにする
        vbs << "Set ws=wb.Sheets(\"" << sheet << "\")\nws.Activate\n";
    else
        // シート名が空なら ActiveSheet (現在アクティブなシート) を使う
        vbs << "Set ws=wb.ActiveSheet\n";

    // ページ設定が指定されていれば VBScript に変換して追記する
    if (ps) vbs << ps_to_vbs("ws", *ps);

    // ExportAsFixedFormat: PDF (Type:=0) として書き出す
    // Type:=0 = xlTypePDF (PDF 形式)
    // Quality:=0 = xlQualityStandard (標準品質)
    // IgnorePrintAreas:=False = 印刷範囲を無視しない
    vbs << "ws.ExportAsFixedFormat Type:=0,Filename:=\"" << ao << "\","
        << "Quality:=0,IncludeDocProperties:=True,"
        << "IgnorePrintAreas:=False,OpenAfterPublish:="
        << (opts.open_after ? "True" : "False") << "\n"
        // ファイルを閉じて Excel を終了する
        << "wb.Close False\nxl.Quit\n"
        // COM オブジェクトを解放する (メモリリーク防止)
        << "Set ws=Nothing:Set wb=Nothing:Set xl=Nothing\n";

    std::string err;
    bool ok = run_vbs(vbs.str(), err);   // VBScript を実行
    return {ok, "Excel", err, ao};       // PrintResult に結果を詰めて返す
}

// ============================================================
//  Excel COM: プリンター送信
// ============================================================
// excel_print(): Excel COM を使って XLSX/XLS ファイルをプリンターに送信する
PrintResult ExcelPrinter::excel_print(const std::string& path,
                                       const PrinterOptions& po,
                                       const std::string& sheet,
                                       const PageSetup* ps) {
    std::string ai = abs_path(path);   // 絶対パスに変換

    std::ostringstream vbs;
    // Excel.Application を COM で生成して非表示で起動する
    vbs << "Dim xl,wb,ws\n"
        << "Set xl=CreateObject(\"Excel.Application\")\n"
        << "xl.Visible=False\n"
        << "xl.DisplayAlerts=False\n";

    // プリンター名が指定されていれば ActivePrinter を設定する
    if (!po.printer_name.empty())
        vbs << "xl.ActivePrinter=\"" << po.printer_name << "\"\n";

    vbs << "Set wb=xl.Workbooks.Open(\"" << ai << "\")\n";

    // 対象シートを選択する
    if (!sheet.empty())
        vbs << "Set ws=wb.Sheets(\"" << sheet << "\")\nws.Activate\n";
    else
        vbs << "Set ws=wb.ActiveSheet\n";

    // ページ設定を適用する
    if (ps) vbs << ps_to_vbs("ws", *ps);

    // 印刷部数を取得する (PageSetup.copies が 1 より大きい場合はその値を使う)
    int copies = (ps && ps->copies > 1) ? ps->copies : 1;
    // PrintOut: プリンターへ送信する
    // Copies: 印刷部数, Collate: 部単位で印刷するか, PrintToFile:=False = ファイルに出力しない
    vbs << "ws.PrintOut Copies:=" << copies
        << ",Collate:=" << (po.collate ? "True" : "False")
        << ",PrintToFile:=False\n"
        << "wb.Close False\nxl.Quit\n"
        << "Set ws=Nothing:Set wb=Nothing:Set xl=Nothing\n";

    std::string err;
    bool ok = run_vbs(vbs.str(), err);
    return {ok, "Excel", err, ""};   // 印刷なので output_path は空文字列
}

// ============================================================
//  公開 API
// ============================================================
// to_pdf(): XLSX/XLS ファイルを PDF として書き出す (公開インターフェース)
PrintResult ExcelPrinter::to_pdf(const std::string& path,
                                  const PdfOptions& opts,
                                  const std::string& sheet,
                                  const PageSetup* setup,
                                  PrintProgressCallback cb) {
    // コールバックが指定されていればセットする
    if (cb) set_progress_callback(cb);
    // output_path が空なら必須引数エラー
    if (opts.output_path.empty()) throw PrintError("output_path is required");
    log("to_pdf: " + path);   // プログレスコールバックに開始メッセージを送る
    return excel_to_pdf(path, opts, sheet, setup);   // 内部実装に委譲
}

// print(): XLSX/XLS ファイルをプリンターに送信する (公開インターフェース)
PrintResult ExcelPrinter::print(const std::string& path,
                                 const PrinterOptions& po,
                                 const std::string& sheet,
                                 const PageSetup* setup,
                                 PrintProgressCallback cb) {
    if (cb) set_progress_callback(cb);
    log("print: " + path);
    return excel_print(path, po, sheet, setup);   // 内部実装に委譲
}

} // namespace excellib
