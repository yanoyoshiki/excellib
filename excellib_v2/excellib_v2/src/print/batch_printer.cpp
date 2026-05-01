// BatchPrinter の実装ファイル
// 複数の Excel ファイルを一括で PDF 変換またはプリンター送信する機能を提供する
// 全ジョブを 1 つの VBScript に まとめることで、Excel.Application の起動コストを削減する
#include "excellib/batch_printer.hpp"
// excel_printer.hpp: ExcelPrinter クラスの定義 (PrintResult, PrinterOptions, PageSetup)
#include "excellib/excel_printer.hpp"
// std::ostringstream: 文字列を組み立てるストリーム (VBScript 生成に使う)
#include <sstream>
// std::ofstream / std::ifstream: ファイルの読み書き
#include <fstream>
// std::chrono: 時間計測用
#include <chrono>
// std::system(): 外部コマンド実行 (cscript でVBScriptを起動)
#include <cstdlib>
// std::memcmp / std::memset: バイト列操作
#include <cstring>
// std::all_of などのアルゴリズム
#include <algorithm>

// Windows 専用ヘッダー (非 Windows では無視される)
#ifdef _WIN32
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#endif

namespace excellib {

// Timer: ジョブの実行時間を計測するヘルパー構造体
// std::chrono::steady_clock は単調増加するクロック (OS の時刻変更の影響を受けない)
struct Timer {
    using Clk = std::chrono::steady_clock;
    Clk::time_point t0 = Clk::now();   // 計測開始時刻 (構築時に記録)
    // ms(): 開始からの経過時間をミリ秒 (double) で返す
    double ms() const {
        return std::chrono::duration<double,std::milli>(Clk::now()-t0).count();
    }
};

// ============================================================
//  ユーティリティ
// ============================================================
// abs_path(): 相対パスを絶対パスに変換するヘルパー
// VBScript に渡すパスは絶対パスでないと Excel が見つけられないため使用する
std::string BatchPrinter::abs_path(const std::string& p) {
#ifdef _WIN32
    char buf[MAX_PATH];
    return GetFullPathNameA(p.c_str(),MAX_PATH,buf,nullptr) ? buf : p;
#else
    return p;
#endif
}

// stem(): ファイルパスからベースネーム (拡張子なし) を取り出す
// 例: "/path/to/report.xlsx" → "report"
std::string BatchPrinter::stem(const std::string& path) {
    // 最後の "/" または "\" を探してファイル名部分を切り出す
    size_t s = path.find_last_of("/\\");
    std::string f = (s==std::string::npos) ? path : path.substr(s+1);
    // "." を探して拡張子を除く
    size_t d = f.rfind('.');
    return (d==std::string::npos) ? f : f.substr(0,d);
}

// job_pdf_name(): ジョブから PDF ファイル名を生成する
// 書式: "[ラベル or ファイル名]_[シート名]_[連番].pdf"
// ファイル名に使えない文字 (/ \ : * ? " < > |) は "_" に置換する
std::string BatchPrinter::job_pdf_name(const PrintJob& job, size_t idx) {
    // ラベルが設定されていればラベルを使い、なければファイル名のベースネームを使う
    std::string base = job.label.empty() ? stem(job.file_path) : job.label;
    // シート名がある場合は "_シート名" を追加する
    if (!job.sheet_name.empty()) base += "_" + job.sheet_name;
    // "_連番.pdf" を追加する (idx は 0-indexed なので +1 して 1-indexed にする)
    base += "_" + std::to_string(idx+1) + ".pdf";
    // Windows ファイル名に使えない特殊文字を "_" に置換する
    for (char& c : base)
        if (c=='/'||c=='\\'||c==':'||c=='*'||c=='?'||c=='"'||c=='<'||c=='>'||c=='|')
            c='_';
    return base;
}

// コンストラクタ・デストラクタ (trivial なため default で自動生成)
BatchPrinter::BatchPrinter()  = default;
BatchPrinter::~BatchPrinter() = default;

// add(): PrintJob 構造体でジョブを追加する。フルエント API のために *this を返す
BatchPrinter& BatchPrinter::add(const PrintJob& j) { jobs_.push_back(j); return *this; }
// add(): ファイルパス・シート名・ページ設定・ラベルを引数で直接指定してジョブを追加する
BatchPrinter& BatchPrinter::add(const std::string& f, const std::string& s,
                                 std::optional<PageSetup> ps, const std::string& l) {
    jobs_.push_back({f,s,std::move(ps),l}); return *this;
}
// add_all(): PrintJob のリストをまとめて追加する
BatchPrinter& BatchPrinter::add_all(const std::vector<PrintJob>& v) {
    for (auto& j:v) jobs_.push_back(j); return *this;
}

// log(): プログレスコールバックが設定されていれば呼び出す
// done: 完了ジョブ数, job: 現在のジョブ, msg: メッセージ
void BatchPrinter::log(size_t done, const PrintJob& job, const std::string& msg) {
    if (cb_) cb_(done, jobs_.size(), job, msg);
}

// ============================================================
//  PageSetup → VBScript (バッチ印刷用)
// ============================================================
// ps_to_vbs(): PageSetup の設定を Excel VBA の PageSetup プロパティ設定コードに変換する
// ws: VBScript 内でシートオブジェクトを参照する変数名
// ps: ページ設定
static std::string ps_to_vbs(const std::string& ws, const PageSetup& ps) {
    std::ostringstream o;
    // 用紙サイズ (例: A4=9, Letter=1)
    o << ws << ".PageSetup.PaperSize = " << uint16_t(ps.paper_size) << "\n";
    // 印刷方向 (Landscape=2, Portrait=1)
    o << ws << ".PageSetup.Orientation = "
      << (ps.orientation==Orientation::Landscape?"2":"1") << "\n";
    // FitTo 設定
    if (ps.fit_to != FitTo::None) {
        o << ws << ".PageSetup.Zoom = False\n";   // ズームを無効化
        bool fw=ps.fit_to==FitTo::Width||ps.fit_to==FitTo::WidthAndHeight;
        bool fh=ps.fit_to==FitTo::Height||ps.fit_to==FitTo::WidthAndHeight;
        // 幅・高さそれぞれのページ数 (False = 無制限)
        o<<ws<<".PageSetup.FitToPagesWide = "<<(fw?std::to_string(ps.fit_to_pages_wide):"False")<<"\n";
        o<<ws<<".PageSetup.FitToPagesTall = "<<(fh?std::to_string(ps.fit_to_pages_tall):"False")<<"\n";
    } else {
        // 倍率指定 (100 = 等倍)
        o << ws << ".PageSetup.Zoom = " << ps.scale_percent << "\n";
    }
    // 余白 (インチ → ポイント変換: Application.InchesToPoints())
    auto& m=ps.margins;
    o<<ws<<".PageSetup.LeftMargin   = Application.InchesToPoints("<<m.left  <<")\n"
     <<ws<<".PageSetup.RightMargin  = Application.InchesToPoints("<<m.right <<")\n"
     <<ws<<".PageSetup.TopMargin    = Application.InchesToPoints("<<m.top   <<")\n"
     <<ws<<".PageSetup.BottomMargin = Application.InchesToPoints("<<m.bottom<<")\n"
     <<ws<<".PageSetup.HeaderMargin = Application.InchesToPoints("<<m.header<<")\n"
     <<ws<<".PageSetup.FooterMargin = Application.InchesToPoints("<<m.footer<<")\n";
    // ヘッダー・フッター (空でない場合のみ設定)
    auto& hf=ps.header_footer;
    if (!hf.odd_header.empty()) o<<ws<<".PageSetup.CenterHeader=\""<<hf.odd_header<<"\"\n";
    if (!hf.odd_footer.empty())  o<<ws<<".PageSetup.CenterFooter=\""<<hf.odd_footer<<"\"\n";
    // グリッド線・白黒印刷の設定
    o<<ws<<".PageSetup.PrintGridlines="<<(ps.print_gridlines?"True":"False")<<"\n"
     <<ws<<".PageSetup.BlackAndWhite="<<(ps.black_and_white?"True":"False")<<"\n";
    // 印刷範囲 (指定されている場合のみ)
    if (ps.print_area)
        o<<ws<<".PageSetup.PrintArea=\""<<ps.print_area->to_range()<<"\"\n";
    // タイトル行の繰り返し (指定されている場合のみ)
    if (ps.repeat_titles && ps.repeat_titles->row_start && ps.repeat_titles->row_end)
        o<<ws<<".PageSetup.PrintTitleRows=\"$"
         <<(*ps.repeat_titles->row_start+1)<<":$"<<(*ps.repeat_titles->row_end+1)<<"\"\n";
    return o.str();
}

// ============================================================
//  VBScript 実行ヘルパー
// ============================================================
// run_vbs(): VBScript コードを一時ファイルに書き出して cscript.exe で実行する
// script: 実行するVBScriptコード文字列, err: エラー内容 (出力引数)
// 戻り値: 成功なら true
static bool run_vbs(const std::string& script, std::string& err) {
#ifdef _WIN32
    char tmp[MAX_PATH], vbs[MAX_PATH+4], log[MAX_PATH+4];
    GetTempPathA(MAX_PATH,tmp);                            // 一時フォルダのパスを取得
    GetTempFileNameA(tmp,"bat",0,vbs);                     // 一時ファイル名を生成 (プレフィックス "bat")
    strcpy(log,vbs); strcat(log,".log"); strcat(vbs,".vbs"); // .vbs と .log の拡張子を付ける
    // エラーハンドリングを追加したVBScriptを組み立てる
    std::string full =
        "On Error Resume Next\n"+script+"\n"          // エラーが起きても継続
        "If Err.Number<>0 Then\n"
        "  Dim fso,f\n"
        "  Set fso=CreateObject(\"Scripting.FileSystemObject\")\n"
        "  Set f=fso.CreateTextFile(\""+std::string(log)+"\",True)\n"
        "  f.WriteLine Err.Number & \": \" & Err.Description\n"  // エラー番号と説明を記録
        "  f.Close\n"
        "End If\n";
    {std::ofstream f(vbs); f<<full;}                   // VBScript をファイルに書き出す
    int ret=std::system(("cscript //nologo \""+std::string(vbs)+"\"").c_str()); // 実行
    {std::ifstream f(log); if(f) std::getline(f,err);}  // エラーログを読む
    DeleteFileA(vbs); DeleteFileA(log);                // 一時ファイルを削除
    return ret==0&&err.empty();                        // 終了コード 0 かつエラーなしで成功
#else
    err="Windows only"; return false;   // 非 Windows は常に失敗
#endif
}

// ============================================================
//  Excel 1 インスタンスで全ジョブをプリンター送信
// ============================================================
// print_all(): 全ジョブを 1 つの VBScript にまとめてプリンターに送信する
// Excel.Application を 1 度だけ起動するため、個別に起動するより高速
BatchResult BatchPrinter::print_all(const PrinterOptions& opts, BatchProgressCallback cb) {
    if (cb) cb_=cb;              // コールバックを設定
    BatchResult r; r.total=jobs_.size();
    if (jobs_.empty()) return r; // ジョブがなければ即座に返す
    Timer t;                     // 全体の実行時間を計測開始

    log(0,jobs_[0],"Starting batch print: "+std::to_string(jobs_.size())+" jobs");

    std::ostringstream vbs;
    // Excel.Application を COM で起動し、非表示・アラートなし・画面更新なしに設定
    vbs << "Dim xl,wb,ws\n"
        << "Set xl=CreateObject(\"Excel.Application\")\n"
        << "xl.Visible=False\n"
        << "xl.DisplayAlerts=False\n"
        << "xl.ScreenUpdating=False\n";   // 画面更新をオフにして高速化する

    // プリンター名が指定されていれば設定する
    if (!opts.printer_name.empty())
        vbs << "xl.ActivePrinter=\"" << opts.printer_name << "\"\n";

    // 各ジョブを VBScript コードに変換して追記する
    for (size_t i=0; i<jobs_.size(); ++i) {
        auto& job=jobs_[i];
        std::string ai=abs_path(job.file_path);   // 絶対パスに変換
        // 変数名にインデックスを付ける (ws0, ws1, ...) ことで変数が衝突しないようにする
        std::string wsv="ws"+std::to_string(i);

        // ジョブのコメントを挿入 (VBScript の ' は行コメント)
        vbs << "\n' ---- Job " << i+1 << "/" << jobs_.size()
            << ": " << job.file_path << " ----\n"
            << "Set wb=xl.Workbooks.Open(\"" << ai << "\")\n";

        // 対象シートを選択する
        if (!job.sheet_name.empty())
            vbs << "Set " << wsv << "=wb.Sheets(\"" << job.sheet_name << "\")\n"
                << wsv << ".Activate\n";
        else
            vbs << "Set " << wsv << "=wb.ActiveSheet\n";

        // ページ設定がある場合は適用する
        if (job.setup) vbs << ps_to_vbs(wsv, *job.setup);

        // 印刷部数を取得する
        int copies = (job.setup && job.setup->copies>1) ? job.setup->copies : 1;
        // PrintOut: プリンターへ送信
        vbs << wsv << ".PrintOut Copies:=" << copies
            << ",Collate:=" << (opts.collate?"True":"False")
            << ",PrintToFile:=False\n"
            // ファイルを閉じてシート変数を解放する
            << "wb.Close False\n"
            << "Set " << wsv << "=Nothing:Set wb=Nothing\n";
    }

    // Excel を終了してオブジェクトを解放する
    vbs << "\nxl.Quit\nSet xl=Nothing\n";

    std::string err; bool ok=run_vbs(vbs.str(),err);   // VBScript を実行
    r.total_elapsed_ms=t.ms();   // 全体の実行時間を記録

    // 全ジョブの結果を BatchResult に登録する
    // VBScript は全ジョブをまとめて実行するため、個別の成否が分からない
    // そのため全ジョブに同じ成否 (ok/err) を設定する
    for (size_t i=0; i<jobs_.size(); ++i) {
        r.jobs.push_back({i,jobs_[i].file_path,jobs_[i].sheet_name,ok,ok?"":err});
        ok ? ++r.succeeded : ++r.failed;
    }
    return r;
}

// ============================================================
//  Excel 1 インスタンスで全ジョブを 1 つの PDF に結合
// ============================================================
// to_pdf_merged(): 全ジョブを 1 つの PDF ファイルにまとめる
// 方法: 空の作業ワークブックに各ジョブのシートをコピーして、
//       まとめて 1 つの PDF として出力する
BatchResult BatchPrinter::to_pdf_merged(const std::string& out_path,
                                         BatchProgressCallback cb) {
    if (cb) cb_=cb;
    BatchResult r; r.total=jobs_.size(); r.output_path=out_path;
    if (jobs_.empty()) return r;
    if (out_path.empty()) throw PrintError("output_path required");
    Timer t;

    log(0,jobs_[0],"Starting batch PDF (merged): "+std::to_string(jobs_.size())+" jobs");

    std::string ao=abs_path(out_path);   // 出力パスを絶対パスに変換
    std::ostringstream vbs;
    // Excel.Application を起動する
    vbs << "Dim xl,tmp_wb,wb,ws,dummy\n"
        << "Set xl=CreateObject(\"Excel.Application\")\n"
        << "xl.Visible=False\n"
        << "xl.DisplayAlerts=False\n"
        << "xl.ScreenUpdating=False\n"
        // 結合先となる空のワークブックを新規作成する
        << "Set tmp_wb=xl.Workbooks.Add\n"
        // デフォルトで作成されるシートの名前を記録しておく (後で削除するため)
        << "dummy=tmp_wb.Sheets(1).Name\n";

    // 各ジョブのシートを tmp_wb にコピーする
    for (size_t i=0; i<jobs_.size(); ++i) {
        auto& job=jobs_[i];
        std::string ai=abs_path(job.file_path);   // 入力ファイルの絶対パス
        std::string wsv="ws"+std::to_string(i);   // ユニークなシート変数名

        vbs << "\n' ---- Job " << i+1 << " ----\n"
            << "Set wb=xl.Workbooks.Open(\"" << ai << "\")\n";

        // 対象シートを選択する
        if (!job.sheet_name.empty())
            vbs << "Set " << wsv << "=wb.Sheets(\"" << job.sheet_name << "\")\n";
        else
            vbs << "Set " << wsv << "=wb.ActiveSheet\n";

        // ページ設定を適用する
        if (job.setup) vbs << ps_to_vbs(wsv, *job.setup);

        // tmp_wb の末尾にシートをコピーする
        // After:=tmp_wb.Sheets(tmp_wb.Sheets.Count) = tmp_wb の最後のシートの後
        vbs << wsv << ".Copy After:=tmp_wb.Sheets(tmp_wb.Sheets.Count)\n"
            // 元のファイルを閉じてシート変数を解放する
            << "wb.Close False\n"
            << "Set " << wsv << "=Nothing:Set wb=Nothing\n";
    }

    // Workbooks.Add 時に自動作成されたデフォルト空シートを削除する
    // (このシートがあると PDF に余分なページが含まれてしまう)
    vbs << "\nApplication.DisplayAlerts=False\n"
        << "tmp_wb.Sheets(dummy).Delete\n"
        // tmp_wb 全体を 1 つの PDF として書き出す
        << "tmp_wb.ExportAsFixedFormat Type:=0,Filename:=\"" << ao << "\","
        << "Quality:=0,IncludeDocProperties:=True,IgnorePrintAreas:=False\n"
        << "tmp_wb.Close False\n"
        << "xl.Quit\nSet xl=Nothing\n";

    std::string err; bool ok=run_vbs(vbs.str(),err);
    r.total_elapsed_ms=t.ms();
    // 全ジョブに同じ成否を設定する (VBScript が一括実行のため個別判断不可)
    for (size_t i=0; i<jobs_.size(); ++i) {
        r.jobs.push_back({i,jobs_[i].file_path,jobs_[i].sheet_name,ok,ok?"":err});
        ok ? ++r.succeeded : ++r.failed;
    }
    return r;
}

// ============================================================
//  Excel 1 インスタンスでジョブごとに個別 PDF
// ============================================================
// to_pdf_individual(): 各ジョブを個別の PDF ファイルとして書き出す
// 出力先ディレクトリに [ラベル]_[シート名]_[連番].pdf という名前で保存する
BatchResult BatchPrinter::to_pdf_individual(const std::string& out_dir,
                                             BatchProgressCallback cb) {
    if (cb) cb_=cb;
    BatchResult r; r.total=jobs_.size();
    if (jobs_.empty()) return r;
    Timer t;

    log(0,jobs_[0],"Starting batch PDF (individual): "+std::to_string(jobs_.size())+" jobs");

    std::ostringstream vbs;
    // Excel.Application を起動する
    vbs << "Dim xl,wb,ws\n"
        << "Set xl=CreateObject(\"Excel.Application\")\n"
        << "xl.Visible=False\n"
        << "xl.DisplayAlerts=False\n"
        << "xl.ScreenUpdating=False\n";

    // 各ジョブを処理する
    for (size_t i=0; i<jobs_.size(); ++i) {
        auto& job=jobs_[i];
        std::string ai=abs_path(job.file_path);          // 入力ファイルの絶対パス
        // 出力 PDF のパス: 出力ディレクトリ + "/" + PDF ファイル名
        std::string pdf=abs_path(out_dir+"/"+job_pdf_name(job,i));
        std::string wsv="ws"+std::to_string(i);          // ユニークなシート変数名

        vbs << "\n' ---- Job " << i+1 << " ----\n"
            << "Set wb=xl.Workbooks.Open(\"" << ai << "\")\n";

        // 対象シートを選択する
        if (!job.sheet_name.empty())
            vbs << "Set " << wsv << "=wb.Sheets(\"" << job.sheet_name << "\")\n"
                << wsv << ".Activate\n";
        else
            vbs << "Set " << wsv << "=wb.ActiveSheet\n";

        // ページ設定を適用する
        if (job.setup) vbs << ps_to_vbs(wsv, *job.setup);

        // 個別 PDF として書き出す
        vbs << wsv << ".ExportAsFixedFormat Type:=0,Filename:=\"" << pdf << "\","
            << "Quality:=0,IgnorePrintAreas:=False\n"
            // ファイルを閉じてシート変数を解放する
            << "wb.Close False\n"
            << "Set " << wsv << "=Nothing:Set wb=Nothing\n";
    }

    // Excel を終了してオブジェクトを解放する
    vbs << "\nxl.Quit\nSet xl=Nothing\n";

    std::string err; bool ok=run_vbs(vbs.str(),err);
    r.total_elapsed_ms=t.ms();
    // 全ジョブに同じ成否を設定する
    for (size_t i=0; i<jobs_.size(); ++i) {
        r.jobs.push_back({i,jobs_[i].file_path,jobs_[i].sheet_name,ok,ok?"":err});
        ok ? ++r.succeeded : ++r.failed;
    }
    return r;
}

} // namespace excellib
