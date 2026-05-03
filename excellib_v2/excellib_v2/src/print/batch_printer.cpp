#include "excellib/batch_printer.hpp"
#include "excellib/excel_printer.hpp"
#include <sstream>
#include <fstream>
#include <chrono>
#include <cstdlib>
#include <cstring>
#include <algorithm>

#ifdef _WIN32
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#endif

namespace excellib {

struct Timer {
    using Clk = std::chrono::steady_clock;
    Clk::time_point t0 = Clk::now();
    double ms() const {
        return std::chrono::duration<double,std::milli>(Clk::now()-t0).count();
    }
};

// ============================================================
//  ユーティリティ
// ============================================================
std::string BatchPrinter::abs_path(const std::string& p) {
#ifdef _WIN32
    char buf[MAX_PATH];
    return GetFullPathNameA(p.c_str(),MAX_PATH,buf,nullptr) ? buf : p;
#else
    return p;
#endif
}

std::string BatchPrinter::stem(const std::string& path) {
    size_t s = path.find_last_of("/\\");
    std::string f = (s==std::string::npos) ? path : path.substr(s+1);
    size_t d = f.rfind('.');
    return (d==std::string::npos) ? f : f.substr(0,d);
}

std::string BatchPrinter::job_pdf_name(const PrintJob& job, size_t idx) {
    std::string base = job.label.empty() ? stem(job.file_path) : job.label;
    if (!job.sheet_name.empty()) base += "_" + job.sheet_name;
    base += "_" + std::to_string(idx+1) + ".pdf";
    for (char& c : base)
        if (c=='/'||c=='\\'||c==':'||c=='*'||c=='?'||c=='"'||c=='<'||c=='>'||c=='|')
            c='_';
    return base;
}

BatchPrinter::BatchPrinter()  = default;
BatchPrinter::~BatchPrinter() = default;

BatchPrinter& BatchPrinter::add(const PrintJob& j) { jobs_.push_back(j); return *this; }
BatchPrinter& BatchPrinter::add(const std::string& f, const std::string& s,
                                 std::optional<PageSetup> ps, const std::string& l) {
    jobs_.push_back({f,s,std::move(ps),l}); return *this;
}
BatchPrinter& BatchPrinter::add_all(const std::vector<PrintJob>& v) {
    for (auto& j:v) { jobs_.push_back(j); }
    return *this;
}

void BatchPrinter::log(size_t done, const PrintJob& job, const std::string& msg) {
    if (cb_) cb_(done, jobs_.size(), job, msg);
}

// ============================================================
//  PageSetup → VBScript（批处理用）
// ============================================================
static std::string ps_to_vbs(const std::string& ws, const PageSetup& ps) {
    std::ostringstream o;
    o << ws << ".PageSetup.PaperSize = " << uint16_t(ps.paper_size) << "\n";
    o << ws << ".PageSetup.Orientation = "
      << (ps.orientation==Orientation::Landscape?"2":"1") << "\n";
    if (ps.fit_to != FitTo::None) {
        o << ws << ".PageSetup.Zoom = False\n";
        bool fw=ps.fit_to==FitTo::Width||ps.fit_to==FitTo::WidthAndHeight;
        bool fh=ps.fit_to==FitTo::Height||ps.fit_to==FitTo::WidthAndHeight;
        o<<ws<<".PageSetup.FitToPagesWide = "<<(fw?std::to_string(ps.fit_to_pages_wide):"False")<<"\n";
        o<<ws<<".PageSetup.FitToPagesTall = "<<(fh?std::to_string(ps.fit_to_pages_tall):"False")<<"\n";
    } else {
        o << ws << ".PageSetup.Zoom = " << ps.scale_percent << "\n";
    }
    auto& m=ps.margins;
    o<<ws<<".PageSetup.LeftMargin   = Application.InchesToPoints("<<m.left  <<")\n"
     <<ws<<".PageSetup.RightMargin  = Application.InchesToPoints("<<m.right <<")\n"
     <<ws<<".PageSetup.TopMargin    = Application.InchesToPoints("<<m.top   <<")\n"
     <<ws<<".PageSetup.BottomMargin = Application.InchesToPoints("<<m.bottom<<")\n"
     <<ws<<".PageSetup.HeaderMargin = Application.InchesToPoints("<<m.header<<")\n"
     <<ws<<".PageSetup.FooterMargin = Application.InchesToPoints("<<m.footer<<")\n";
    auto& hf=ps.header_footer;
    if (!hf.odd_header.empty()) o<<ws<<".PageSetup.CenterHeader=\""<<hf.odd_header<<"\"\n";
    if (!hf.odd_footer.empty())  o<<ws<<".PageSetup.CenterFooter=\""<<hf.odd_footer<<"\"\n";
    o<<ws<<".PageSetup.PrintGridlines="<<(ps.print_gridlines?"True":"False")<<"\n"
     <<ws<<".PageSetup.BlackAndWhite="<<(ps.black_and_white?"True":"False")<<"\n";
    if (ps.print_area)
        o<<ws<<".PageSetup.PrintArea=\""<<ps.print_area->to_range()<<"\"\n";
    if (ps.repeat_titles && ps.repeat_titles->row_start && ps.repeat_titles->row_end)
        o<<ws<<".PageSetup.PrintTitleRows=\"$"
         <<(*ps.repeat_titles->row_start+1)<<":$"<<(*ps.repeat_titles->row_end+1)<<"\"\n";
    return o.str();
}

// ============================================================
//  VBScript 実行
// ============================================================
static bool run_vbs(const std::string& script, std::string& err) {
#ifdef _WIN32
    char tmp[MAX_PATH], vbs[MAX_PATH+4], log[MAX_PATH+4];
    GetTempPathA(MAX_PATH,tmp);
    GetTempFileNameA(tmp,"bat",0,vbs);
    strcpy(log,vbs); strcat(log,".log"); strcat(vbs,".vbs");
    std::string full =
        "On Error Resume Next\n"+script+"\n"
        "If Err.Number<>0 Then\n"
        "  Dim fso,f\n"
        "  Set fso=CreateObject(\"Scripting.FileSystemObject\")\n"
        "  Set f=fso.CreateTextFile(\""+std::string(log)+"\",True)\n"
        "  f.WriteLine Err.Number & \": \" & Err.Description\n"
        "  f.Close\n"
        "End If\n";
    {std::ofstream f(vbs); f<<full;}
    int ret=std::system(("cscript //nologo \""+std::string(vbs)+"\"").c_str());
    {std::ifstream f(log); if(f) std::getline(f,err);}
    DeleteFileA(vbs); DeleteFileA(log);
    return ret==0&&err.empty();
#else
    err="Windows only"; return false;
#endif
}

// ============================================================
//  Excel 1インスタンスで全ジョブをプリンター送信
// ============================================================
BatchResult BatchPrinter::print_all(const PrinterOptions& opts, BatchProgressCallback cb) {
    if (cb) cb_=cb;
    BatchResult r; r.total=jobs_.size();
    if (jobs_.empty()) return r;
    Timer t;

    log(0,jobs_[0],"Starting batch print: "+std::to_string(jobs_.size())+" jobs");

    std::ostringstream vbs;
    vbs << "Dim xl,wb,ws\n"
        << "Set xl=CreateObject(\"Excel.Application\")\n"
        << "xl.Visible=False\n"
        << "xl.DisplayAlerts=False\n"
        << "xl.ScreenUpdating=False\n";   // 画面更新オフで高速化

    if (!opts.printer_name.empty())
        vbs << "xl.ActivePrinter=\"" << opts.printer_name << "\"\n";

    for (size_t i=0; i<jobs_.size(); ++i) {
        auto& job=jobs_[i];
        std::string ai=abs_path(job.file_path);
        std::string wsv="ws"+std::to_string(i);

        vbs << "\n' ---- Job " << i+1 << "/" << jobs_.size()
            << ": " << job.file_path << " ----\n"
            << "Set wb=xl.Workbooks.Open(\"" << ai << "\")\n";

        if (!job.sheet_name.empty())
            vbs << "Set " << wsv << "=wb.Sheets(\"" << job.sheet_name << "\")\n"
                << wsv << ".Activate\n";
        else
            vbs << "Set " << wsv << "=wb.ActiveSheet\n";

        if (job.setup) vbs << ps_to_vbs(wsv, *job.setup);

        int copies = (job.setup && job.setup->copies>1) ? job.setup->copies : 1;
        vbs << wsv << ".PrintOut Copies:=" << copies
            << ",Collate:=" << (opts.collate?"True":"False")
            << ",PrintToFile:=False\n"
            << "wb.Close False\n"
            << "Set " << wsv << "=Nothing:Set wb=Nothing\n";
    }

    vbs << "\nxl.Quit\nSet xl=Nothing\n";

    std::string err; bool ok=run_vbs(vbs.str(),err);
    r.total_elapsed_ms=t.ms();

    for (size_t i=0; i<jobs_.size(); ++i) {
        r.jobs.push_back({i,jobs_[i].file_path,jobs_[i].sheet_name,ok,ok?"":err});
        ok ? ++r.succeeded : ++r.failed;
    }
    return r;
}

// ============================================================
//  Excel 1インスタンスで全ジョブを1 PDF に結合
// ============================================================
BatchResult BatchPrinter::to_pdf_merged(const std::string& out_path,
                                         BatchProgressCallback cb) {
    if (cb) cb_=cb;
    BatchResult r; r.total=jobs_.size(); r.output_path=out_path;
    if (jobs_.empty()) return r;
    if (out_path.empty()) throw PrintError("output_path required");
    Timer t;

    log(0,jobs_[0],"Starting batch PDF (merged): "+std::to_string(jobs_.size())+" jobs");

    std::string ao=abs_path(out_path);
    std::ostringstream vbs;
    vbs << "Dim xl,tmp_wb,wb,ws,dummy\n"
        << "Set xl=CreateObject(\"Excel.Application\")\n"
        << "xl.Visible=False\n"
        << "xl.DisplayAlerts=False\n"
        << "xl.ScreenUpdating=False\n"
        // 結合先の空ワークブックを作成
        << "Set tmp_wb=xl.Workbooks.Add\n"
        << "dummy=tmp_wb.Sheets(1).Name\n";

    for (size_t i=0; i<jobs_.size(); ++i) {
        auto& job=jobs_[i];
        std::string ai=abs_path(job.file_path);
        std::string wsv="ws"+std::to_string(i);

        vbs << "\n' ---- Job " << i+1 << " ----\n"
            << "Set wb=xl.Workbooks.Open(\"" << ai << "\")\n";

        if (!job.sheet_name.empty())
            vbs << "Set " << wsv << "=wb.Sheets(\"" << job.sheet_name << "\")\n";
        else
            vbs << "Set " << wsv << "=wb.ActiveSheet\n";

        if (job.setup) vbs << ps_to_vbs(wsv, *job.setup);

        // tmp_wb の末尾にシートをコピー
        vbs << wsv << ".Copy After:=tmp_wb.Sheets(tmp_wb.Sheets.Count)\n"
            << "wb.Close False\n"
            << "Set " << wsv << "=Nothing:Set wb=Nothing\n";
    }

    // デフォルトの空シートを削除して一括 PDF 出力
    vbs << "\nApplication.DisplayAlerts=False\n"
        << "tmp_wb.Sheets(dummy).Delete\n"
        << "tmp_wb.ExportAsFixedFormat Type:=0,Filename:=\"" << ao << "\","
        << "Quality:=0,IncludeDocProperties:=True,IgnorePrintAreas:=False\n"
        << "tmp_wb.Close False\n"
        << "xl.Quit\nSet xl=Nothing\n";

    std::string err; bool ok=run_vbs(vbs.str(),err);
    r.total_elapsed_ms=t.ms();
    for (size_t i=0; i<jobs_.size(); ++i) {
        r.jobs.push_back({i,jobs_[i].file_path,jobs_[i].sheet_name,ok,ok?"":err});
        ok ? ++r.succeeded : ++r.failed;
    }
    return r;
}

// ============================================================
//  Excel 1インスタンスでジョブごとに個別 PDF
// ============================================================
BatchResult BatchPrinter::to_pdf_individual(const std::string& out_dir,
                                             BatchProgressCallback cb) {
    if (cb) cb_=cb;
    BatchResult r; r.total=jobs_.size();
    if (jobs_.empty()) return r;
    Timer t;

    log(0,jobs_[0],"Starting batch PDF (individual): "+std::to_string(jobs_.size())+" jobs");

    std::ostringstream vbs;
    vbs << "Dim xl,wb,ws\n"
        << "Set xl=CreateObject(\"Excel.Application\")\n"
        << "xl.Visible=False\n"
        << "xl.DisplayAlerts=False\n"
        << "xl.ScreenUpdating=False\n";

    for (size_t i=0; i<jobs_.size(); ++i) {
        auto& job=jobs_[i];
        std::string ai=abs_path(job.file_path);
        std::string pdf=abs_path(out_dir+"/"+job_pdf_name(job,i));
        std::string wsv="ws"+std::to_string(i);

        vbs << "\n' ---- Job " << i+1 << " ----\n"
            << "Set wb=xl.Workbooks.Open(\"" << ai << "\")\n";

        if (!job.sheet_name.empty())
            vbs << "Set " << wsv << "=wb.Sheets(\"" << job.sheet_name << "\")\n"
                << wsv << ".Activate\n";
        else
            vbs << "Set " << wsv << "=wb.ActiveSheet\n";

        if (job.setup) vbs << ps_to_vbs(wsv, *job.setup);

        vbs << wsv << ".ExportAsFixedFormat Type:=0,Filename:=\"" << pdf << "\","
            << "Quality:=0,IgnorePrintAreas:=False\n"
            << "wb.Close False\n"
            << "Set " << wsv << "=Nothing:Set wb=Nothing\n";
    }

    vbs << "\nxl.Quit\nSet xl=Nothing\n";

    std::string err; bool ok=run_vbs(vbs.str(),err);
    r.total_elapsed_ms=t.ms();
    for (size_t i=0; i<jobs_.size(); ++i) {
        r.jobs.push_back({i,jobs_[i].file_path,jobs_[i].sheet_name,ok,ok?"":err});
        ok ? ++r.succeeded : ++r.failed;
    }
    return r;
}

} // namespace excellib
