#include "excellib/excel_printer.hpp"
#include <sstream>
#include <fstream>
#include <cstring>
#include <cstdlib>

#ifdef _WIN32
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <objbase.h>
#pragma comment(lib,"ole32.lib")
#pragma comment(lib,"oleaut32.lib")
#pragma comment(lib,"winspool.lib")
#endif

namespace excellib {

// ============================================================
//  COM 初期化ガード
// ============================================================
struct ExcelPrinter::COMGuard {
#ifdef _WIN32
    bool ok{false};
    COMGuard()  { ok = SUCCEEDED(CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED)); }
    ~COMGuard() { if (ok) CoUninitialize(); }
#endif
};

ExcelPrinter::ExcelPrinter() : com_(std::make_unique<COMGuard>()) {}
ExcelPrinter::~ExcelPrinter() = default;

void ExcelPrinter::log(const std::string& msg) {
    if (progress_cb_) progress_cb_(msg);
}

// ============================================================
//  プリンター一覧
// ============================================================
std::vector<std::string> ExcelPrinter::list_printers() const {
    std::vector<std::string> result;
#ifdef _WIN32
    DWORD need=0, cnt=0;
    EnumPrintersA(PRINTER_ENUM_LOCAL|PRINTER_ENUM_CONNECTIONS,nullptr,4,nullptr,0,&need,&cnt);
    std::vector<uint8_t> buf(need);
    if (EnumPrintersA(PRINTER_ENUM_LOCAL|PRINTER_ENUM_CONNECTIONS,nullptr,4,
                      buf.data(),need,&need,&cnt)) {
        auto* info = reinterpret_cast<PRINTER_INFO_4A*>(buf.data());
        for (DWORD i=0; i<cnt; ++i)
            if (info[i].pPrinterName) result.push_back(info[i].pPrinterName);
    }
#endif
    return result;
}

std::string ExcelPrinter::default_printer() const {
#ifdef _WIN32
    char buf[256]={}; DWORD sz=sizeof(buf);
    GetDefaultPrinterA(buf,&sz); return buf;
#else
    return "";
#endif
}

// ============================================================
//  PageSetup → VBScript 変換
// ============================================================
static std::string ps_to_vbs(const std::string& ws, const PageSetup& ps) {
    std::ostringstream o;
    o << ws << ".PageSetup.PaperSize = " << uint16_t(ps.paper_size) << "\n";
    o << ws << ".PageSetup.Orientation = "
      << (ps.orientation==Orientation::Landscape ? "2" : "1") << "\n";

    if (ps.fit_to != FitTo::None) {
        o << ws << ".PageSetup.Zoom = False\n";
        bool fw = ps.fit_to==FitTo::Width  || ps.fit_to==FitTo::WidthAndHeight;
        bool fh = ps.fit_to==FitTo::Height || ps.fit_to==FitTo::WidthAndHeight;
        o << ws << ".PageSetup.FitToPagesWide = "
          << (fw ? std::to_string(ps.fit_to_pages_wide) : "False") << "\n";
        o << ws << ".PageSetup.FitToPagesTall = "
          << (fh ? std::to_string(ps.fit_to_pages_tall) : "False") << "\n";
    } else {
        o << ws << ".PageSetup.Zoom = " << ps.scale_percent << "\n";
    }

    auto& m = ps.margins;
    o << ws << ".PageSetup.LeftMargin   = Application.InchesToPoints(" << m.left   << ")\n"
      << ws << ".PageSetup.RightMargin  = Application.InchesToPoints(" << m.right  << ")\n"
      << ws << ".PageSetup.TopMargin    = Application.InchesToPoints(" << m.top    << ")\n"
      << ws << ".PageSetup.BottomMargin = Application.InchesToPoints(" << m.bottom << ")\n"
      << ws << ".PageSetup.HeaderMargin = Application.InchesToPoints(" << m.header << ")\n"
      << ws << ".PageSetup.FooterMargin = Application.InchesToPoints(" << m.footer << ")\n";

    auto& hf = ps.header_footer;
    if (!hf.odd_header.empty())
        o << ws << ".PageSetup.CenterHeader = \"" << hf.odd_header << "\"\n";
    if (!hf.odd_footer.empty())
        o << ws << ".PageSetup.CenterFooter = \"" << hf.odd_footer << "\"\n";
    if (!hf.even_header.empty())
        o << ws << ".PageSetup.EvenHeader = \"" << hf.even_header << "\"\n";
    if (!hf.even_footer.empty())
        o << ws << ".PageSetup.EvenFooter = \"" << hf.even_footer << "\"\n";

    o << ws << ".PageSetup.PrintGridlines = "
      << (ps.print_gridlines ? "True" : "False") << "\n"
      << ws << ".PageSetup.BlackAndWhite = "
      << (ps.black_and_white ? "True" : "False") << "\n"
      << ws << ".PageSetup.Draft = "
      << (ps.draft_quality ? "True" : "False") << "\n"
      << ws << ".PageSetup.PrintHeadings = "
      << (ps.print_row_col_headings ? "True" : "False") << "\n";

    if (ps.print_area)
        o << ws << ".PageSetup.PrintArea = \"" << ps.print_area->to_range() << "\"\n";

    if (ps.repeat_titles) {
        auto& rt = *ps.repeat_titles;
        if (rt.row_start && rt.row_end)
            o << ws << ".PageSetup.PrintTitleRows = \"$"
              << (*rt.row_start+1) << ":$" << (*rt.row_end+1) << "\"\n";
        if (rt.col_start && rt.col_end) {
            auto col_str = [](uint32_t c) {
                std::string r; ++c;
                while (c>0) { --c; r+=char('A'+c%26); c/=26; }
                std::reverse(r.begin(),r.end()); return r;
            };
            o << ws << ".PageSetup.PrintTitleColumns = \"$"
              << col_str(*rt.col_start) << ":$" << col_str(*rt.col_end) << "\"\n";
        }
    }

    if (ps.first_page_number > 0)
        o << ws << ".PageSetup.FirstPageNumber = " << ps.first_page_number << "\n";

    return o.str();
}

// ============================================================
//  VBScript 実行ヘルパー
// ============================================================
static bool run_vbs(const std::string& script, std::string& err_out) {
#ifdef _WIN32
    char tmp[MAX_PATH], vbs[MAX_PATH+4], log[MAX_PATH+4];
    GetTempPathA(MAX_PATH, tmp);
    GetTempFileNameA(tmp, "exc", 0, vbs);
    strcpy(log, vbs); strcat(log, ".log"); strcat(vbs, ".vbs");

    std::string full =
        "On Error Resume Next\n" + script + "\n"
        "If Err.Number <> 0 Then\n"
        "  Dim fso,f\n"
        "  Set fso=CreateObject(\"Scripting.FileSystemObject\")\n"
        "  Set f=fso.CreateTextFile(\"" + std::string(log) + "\",True)\n"
        "  f.WriteLine Err.Number & \": \" & Err.Description\n"
        "  f.Close\n"
        "End If\n";

    { std::ofstream f(vbs); f << full; }
    int ret = std::system(("cscript //nologo \"" + std::string(vbs) + "\"").c_str());
    { std::ifstream f(log); if(f) std::getline(f, err_out); }
    DeleteFileA(vbs); DeleteFileA(log);
    return ret==0 && err_out.empty();
#else
    err_out = "Windows only";
    return false;
#endif
}

static std::string abs_path(const std::string& p) {
#ifdef _WIN32
    char buf[MAX_PATH];
    return GetFullPathNameA(p.c_str(),MAX_PATH,buf,nullptr) ? buf : p;
#else
    return p;
#endif
}

// ============================================================
//  Excel COM: PDF 出力
// ============================================================
PrintResult ExcelPrinter::excel_to_pdf(const std::string& path,
                                        const PdfOptions& opts,
                                        const std::string& sheet,
                                        const PageSetup* ps) {
    std::string ai = abs_path(path);
    std::string ao = abs_path(opts.output_path);

    std::ostringstream vbs;
    vbs << "Dim xl,wb,ws\n"
        << "Set xl=CreateObject(\"Excel.Application\")\n"
        << "xl.Visible=False\nxl.DisplayAlerts=False\n"
        << "Set wb=xl.Workbooks.Open(\"" << ai << "\")\n";

    if (!sheet.empty())
        vbs << "Set ws=wb.Sheets(\"" << sheet << "\")\nws.Activate\n";
    else
        vbs << "Set ws=wb.ActiveSheet\n";

    if (ps) vbs << ps_to_vbs("ws", *ps);

    vbs << "ws.ExportAsFixedFormat Type:=0,Filename:=\"" << ao << "\","
        << "Quality:=0,IncludeDocProperties:=True,"
        << "IgnorePrintAreas:=False,OpenAfterPublish:="
        << (opts.open_after ? "True" : "False") << "\n"
        << "wb.Close False\nxl.Quit\n"
        << "Set ws=Nothing:Set wb=Nothing:Set xl=Nothing\n";

    std::string err;
    bool ok = run_vbs(vbs.str(), err);
    return {ok, "Excel", err, ao};
}

// ============================================================
//  Excel COM: プリンター送信
// ============================================================
PrintResult ExcelPrinter::excel_print(const std::string& path,
                                       const PrinterOptions& po,
                                       const std::string& sheet,
                                       const PageSetup* ps) {
    std::string ai = abs_path(path);

    std::ostringstream vbs;
    vbs << "Dim xl,wb,ws\n"
        << "Set xl=CreateObject(\"Excel.Application\")\n"
        << "xl.Visible=False\nxl.DisplayAlerts=False\n";

    if (!po.printer_name.empty())
        vbs << "xl.ActivePrinter=\"" << po.printer_name << "\"\n";

    vbs << "Set wb=xl.Workbooks.Open(\"" << ai << "\")\n";

    if (!sheet.empty())
        vbs << "Set ws=wb.Sheets(\"" << sheet << "\")\nws.Activate\n";
    else
        vbs << "Set ws=wb.ActiveSheet\n";

    if (ps) vbs << ps_to_vbs("ws", *ps);

    int copies = (ps && ps->copies > 1) ? ps->copies : 1;
    vbs << "ws.PrintOut Copies:=" << copies
        << ",Collate:=" << (po.collate ? "True" : "False")
        << ",PrintToFile:=False\n"
        << "wb.Close False\nxl.Quit\n"
        << "Set ws=Nothing:Set wb=Nothing:Set xl=Nothing\n";

    std::string err;
    bool ok = run_vbs(vbs.str(), err);
    return {ok, "Excel", err, ""};
}

// ============================================================
//  公開 API
// ============================================================
PrintResult ExcelPrinter::to_pdf(const std::string& path,
                                  const PdfOptions& opts,
                                  const std::string& sheet,
                                  const PageSetup* setup,
                                  PrintProgressCallback cb) {
    if (cb) set_progress_callback(cb);
    if (opts.output_path.empty()) throw PrintError("output_path is required");
    log("to_pdf: " + path);
    return excel_to_pdf(path, opts, sheet, setup);
}

PrintResult ExcelPrinter::print(const std::string& path,
                                 const PrinterOptions& po,
                                 const std::string& sheet,
                                 const PageSetup* setup,
                                 PrintProgressCallback cb) {
    if (cb) set_progress_callback(cb);
    log("print: " + path);
    return excel_print(path, po, sheet, setup);
}

} // namespace excellib
