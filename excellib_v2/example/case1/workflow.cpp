#include "workflow.hpp"
#include "config.hpp"
#include <filesystem>
#include <stdexcept>

namespace case1 {

namespace fs = std::filesystem;

// ============================================================
//  パス組み立て
// ============================================================
std::string Workflow::construct_path_base(std::string_view input) {
    std::string tmpl(config::PATH_TEMPLATE);
    const std::string token = "{INPUT}";
    auto pos = tmpl.find(token);
    if (pos == std::string::npos) return tmpl;
    return tmpl.substr(0, pos) + std::string(input) + tmpl.substr(pos + token.size());
}

std::optional<std::string> Workflow::resolve_file(std::string_view path_base) {
    std::string base(path_base);
    for (const char* ext : {".xlsx", ".xls"}) {
        std::string candidate = base + ext;
        std::error_code ec;
        if (fs::exists(candidate, ec) && !ec) return candidate;
    }
    return std::nullopt;
}

// ============================================================
//  マスター読込
// ============================================================
std::vector<MasterRecord> Workflow::load_master(const std::string& path) {
    using namespace excellib;

    auto wb = WorkbookFactory::open(path);

    Sheet* sheet = nullptr;
    if (config::MASTER_SHEET_NAME.empty()) {
        sheet = &wb->sheet(size_t{0});
    } else {
        sheet = &wb->sheet(std::string(config::MASTER_SHEET_NAME));
    }

    // COL_KEY 列を空白まで読み下ろして行数を確定
    CellAddress start{config::MASTER_START_ROW, config::COL_KEY};
    auto key_cells = sheet->read_until_blank(start, Direction::Down);

    std::vector<MasterRecord> out;
    out.reserve(key_cells.size());

    for (auto& kc : key_cells) {
        MasterRecord rec;
        rec.row_index = kc.address.row;

        auto get_str = [&](uint32_t col) -> std::string {
            auto c = sheet->cell(kc.address.row, col);
            return is_blank(c.value) ? std::string{} : excellib::to_string(c.value);
        };

        rec.file_path  = get_str(config::COL_FILE_PATH);
        rec.sheet_name = get_str(config::COL_SHEET_NAME);
        rec.print_cols = get_str(config::COL_PRINT_COLS);
        rec.label      = get_str(config::COL_LABEL);

        if (rec.file_path.empty()) continue;  // KEY 以外の列が空でもスキップしない（KEY 空のみで打切）
        out.push_back(std::move(rec));
    }
    return out;
}

// ============================================================
//  PrintJob 変換
// ============================================================
std::vector<excellib::PrintJob> Workflow::records_to_jobs(
    const std::vector<MasterRecord>& records,
    bool default_landscape)
{
    using namespace excellib;
    std::vector<PrintJob> jobs;
    jobs.reserve(records.size());

    for (const auto& r : records) {
        if (!r.included) continue;

        PrintJob job;
        job.file_path  = r.file_path;
        job.sheet_name = r.sheet_name;
        job.label      = r.label;

        PageSetup ps;
        ps.orientation = default_landscape ? Orientation::Landscape : Orientation::Portrait;
        if (config::DEFAULT_FIT_TO_WIDTH) {
            ps.fit_to            = FitTo::Width;
            ps.fit_to_pages_wide = config::DEFAULT_FIT_PAGES_WIDE;
        }
        if (!r.print_cols.empty()) {
            ps.print_area = PrintArea::columns_only(r.print_cols);
        }
        job.setup = ps;

        jobs.push_back(std::move(job));
    }
    return jobs;
}

// ============================================================
//  バッチ実行
// ============================================================
excellib::BatchResult Workflow::run_batch(
    const std::vector<excellib::PrintJob>& jobs,
    const RunContext& ctx)
{
    using namespace excellib;

    batch_.clear();
    batch_.reset_cancel();
    batch_.add_all(jobs);

    auto cb = [&](size_t done, size_t total, const PrintJob& job, const std::string& msg) {
        if (ctx.on_progress) {
            std::string display = job.label.empty() ? job.file_path : job.label;
            ctx.on_progress(done, total, display + " — " + msg);
        }
    };

    BatchResult r;
    switch (ctx.mode) {
        case OutputMode::Printer: {
            PrinterOptions po;
            po.printer_name = ctx.printer_name;
            r = batch_.print_all(po, cb);
            break;
        }
        case OutputMode::PdfIndividual:
            r = batch_.to_pdf_individual(ctx.output_dir, cb);
            break;
        case OutputMode::PdfMerged:
            r = batch_.to_pdf_merged(ctx.output_pdf_path, cb);
            break;
    }

    if (ctx.on_job_done) {
        for (auto& jr : r.jobs)
            ctx.on_job_done(jr.job_index, jr.success, jr.error_message);
    }
    return r;
}

}  // namespace case1
