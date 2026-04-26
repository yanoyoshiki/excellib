#pragma once
/**
 * excellib/batch_printer.hpp
 *
 * 複数ファイル・複数シートを Excel 1インスタンスで一括印刷/PDF化。
 * Excel は常にインストール済みを前提とする。
 */
#include "excellib/excel_printer.hpp"
#include <string>
#include <vector>
#include <functional>
#include <optional>

namespace excellib {

// ---- 1件の印刷ジョブ ----
struct PrintJob {
    std::string              file_path;
    std::string              sheet_name;   ///< 空 = 先頭シート
    std::optional<PageSetup> setup;        ///< nullopt = ファイルの設定を使用
    std::string              label;        ///< 進捗表示・PDF ファイル名用

    static PrintJob make(const std::string& file,
                         const std::string& sheet = "",
                         std::optional<PageSetup> ps = std::nullopt,
                         const std::string& label = "") {
        return {file, sheet, std::move(ps), label};
    }
};

// ---- 結果 ----
struct JobResult {
    size_t      job_index{0};
    std::string file_path;
    std::string sheet_name;
    bool        success{false};
    std::string error_message;
    double      elapsed_ms{0.0};
};

struct BatchResult {
    std::vector<JobResult> jobs;
    size_t      total{0};
    size_t      succeeded{0};
    size_t      failed{0};
    double      total_elapsed_ms{0.0};
    std::string output_path;

    bool all_success() const { return failed == 0; }
    std::vector<const JobResult*> failures() const {
        std::vector<const JobResult*> v;
        for (auto& j : jobs) if (!j.success) v.push_back(&j);
        return v;
    }
};

using BatchProgressCallback = std::function<void(
    size_t done, size_t total,
    const PrintJob& job,
    const std::string& message
)>;

// ---- BatchPrinter ----
class BatchPrinter {
public:
    BatchPrinter();
    ~BatchPrinter();

    BatchPrinter& add(const PrintJob& job);
    BatchPrinter& add(const std::string& file,
                      const std::string& sheet = "",
                      std::optional<PageSetup> setup = std::nullopt,
                      const std::string& label = "");
    BatchPrinter& add_all(const std::vector<PrintJob>& jobs);

    size_t job_count() const { return jobs_.size(); }
    void   clear()           { jobs_.clear(); }

    /**
     * プリンターに一括送信。
     * Excel を1回起動して全ジョブを処理する。
     */
    BatchResult print_all(
        const PrinterOptions&   printer_opts = {},
        BatchProgressCallback   cb = nullptr
    );

    /**
     * 全ジョブを1つの PDF に結合して出力。
     * Excel で全シートを一時 Workbook にコピーし ExportAsFixedFormat で一括出力。
     */
    BatchResult to_pdf_merged(
        const std::string&    output_path,
        BatchProgressCallback cb = nullptr
    );

    /**
     * ジョブごとに個別 PDF を出力。
     * output_dir/{label}_{index}.pdf として保存。
     */
    BatchResult to_pdf_individual(
        const std::string&    output_dir,
        BatchProgressCallback cb = nullptr
    );

    void set_progress_callback(BatchProgressCallback cb) { cb_ = cb; }

private:
    std::vector<PrintJob> jobs_;
    BatchProgressCallback cb_;

    void log(size_t done, const PrintJob& job, const std::string& msg);

    static std::string job_pdf_name(const PrintJob& job, size_t index);
    static std::string abs_path(const std::string& p);
    static std::string stem(const std::string& path);
};

} // namespace excellib
