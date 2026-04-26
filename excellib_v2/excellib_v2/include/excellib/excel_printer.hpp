#pragma once
/**
 * excellib/excel_printer.hpp
 *
 * 印刷 / PDF エクスポート API。
 * Windows の Excel COM オートメーションのみを使用。
 * Excel は常にインストール済みを前提とする。
 */
#include "excellib/print_settings.hpp"
#include <string>
#include <vector>
#include <memory>
#include <functional>

namespace excellib {

using PrintProgressCallback = std::function<void(const std::string& message)>;

struct PrintResult {
    bool        success{false};
    std::string engine_used;
    std::string error_message;
    std::string output_path;
};

/**
 * ExcelPrinter
 *
 *   ExcelPrinter p;
 *
 *   PageSetup ps;
 *   ps.orientation = Orientation::Landscape;
 *   ps.fit_to      = FitTo::Width;
 *
 *   p.apply_page_setup("report.xlsx", "Sheet1", ps);
 *   p.to_pdf("report.xlsx", {.output_path="out.pdf"});
 *   p.print("report.xlsx", {.printer_name="Canon MF3010"});
 */
class ExcelPrinter {
public:
    ExcelPrinter();
    ~ExcelPrinter();

    // ---- XLSX にページ設定を焼き込む（Excel 不要、ZIP/XML 操作のみ）----
    void apply_page_setup(const std::string& xlsx_path,
                          const std::string& sheet_name,
                          const PageSetup&   setup);

    // ---- PDF 出力 ----
    PrintResult to_pdf(const std::string&     xlsx_path,
                       const PdfOptions&       opts,
                       const std::string&      sheet_name = "",
                       const PageSetup*        setup      = nullptr,
                       PrintProgressCallback   cb         = nullptr);

    // ---- プリンター送信 ----
    PrintResult print(const std::string&     xlsx_path,
                      const PrinterOptions&   printer_opts = {},
                      const std::string&      sheet_name   = "",
                      const PageSetup*        setup        = nullptr,
                      PrintProgressCallback   cb           = nullptr);

    // ---- プリンター一覧 ----
    std::vector<std::string> list_printers()   const;
    std::string              default_printer() const;

    void set_progress_callback(PrintProgressCallback cb) { progress_cb_ = cb; }

    // 内部公開（batch_printer.cpp / xlsx_page_setup.cpp から使用）
    static void write_page_setup_to_xlsx(const std::string& xlsx_path,
                                          const std::string& sheet_name,
                                          const PageSetup&   setup);
private:
    PrintProgressCallback progress_cb_;
    void log(const std::string& msg);

    PrintResult excel_to_pdf (const std::string&, const PdfOptions&,
                               const std::string&, const PageSetup*);
    PrintResult excel_print  (const std::string&, const PrinterOptions&,
                               const std::string&, const PageSetup*);

    struct COMGuard;
    std::unique_ptr<COMGuard> com_;
};

} // namespace excellib
