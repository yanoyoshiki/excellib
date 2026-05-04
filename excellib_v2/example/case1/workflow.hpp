#pragma once
/**
 * case1 — workflow.hpp
 *
 * GUI 非依存のビジネスロジック層。
 * テストもこのレイヤーを呼び出すだけで完結する。
 */
#include "excellib/excellib.hpp"
#include "excellib/batch_printer.hpp"
#include <string>
#include <string_view>
#include <vector>
#include <optional>
#include <atomic>
#include <functional>

namespace case1 {

// ============================================================
//  マスターから抽出した 1 レコード
// ============================================================
struct MasterRecord {
    uint32_t    row_index{0};        ///< マスター上の行番号 (0-based)
    std::string file_path;
    std::string sheet_name;
    std::string print_cols;          ///< "A:D" 等 / 空なら全列
    std::string label;
    bool        included{true};      ///< GUI チェックボックス用 (デフォルト ON)
};

// ============================================================
//  出力モード
// ============================================================
enum class OutputMode {
    Printer,         ///< プリンターに送信
    PdfIndividual,   ///< 個別 PDF
    PdfMerged,       ///< 結合 PDF
};

// ============================================================
//  実行コンテキスト (GUI から渡す)
// ============================================================
struct RunContext {
    OutputMode  mode{OutputMode::PdfIndividual};
    std::string output_dir;          ///< PdfIndividual: 出力ディレクトリ
    std::string output_pdf_path;     ///< PdfMerged: 出力 PDF パス
    std::string printer_name;        ///< Printer: 空文字列 = デフォルト

    /// 進捗通知コールバック (ワーカースレッドから呼ばれる)
    std::function<void(size_t done, size_t total, const std::string& msg)> on_progress;

    /// ジョブ完了 (成功/失敗) ごとの通知
    std::function<void(size_t index, bool success, const std::string& err)> on_job_done;
};

// ============================================================
//  ワークフロー本体
// ============================================================
class Workflow {
public:
    /// 入力文字列を PATH_TEMPLATE に埋め込んでフルパスを返す（拡張子なし）
    static std::string construct_path_base(std::string_view input);

    /// 拡張子を順に試して、実在するファイルパスを返す。なければ nullopt
    static std::optional<std::string> resolve_file(std::string_view path_base);

    /// マスター Excel を開いてレコードを抽出
    /// @throws excellib::ExcelError
    static std::vector<MasterRecord> load_master(const std::string& path);

    /// レコードを PrintJob に変換
    static std::vector<excellib::PrintJob> records_to_jobs(
        const std::vector<MasterRecord>& records,
        bool default_landscape = false);

    /// バッチ実行（ワーカースレッドから呼ぶ前提）
    excellib::BatchResult run_batch(
        const std::vector<excellib::PrintJob>& jobs,
        const RunContext& ctx);

    /// 実行中のキャンセル要求
    void request_cancel() { batch_.request_cancel(); }
    bool is_cancel_requested() const { return batch_.is_cancel_requested(); }
    void reset_cancel() { batch_.reset_cancel(); }

private:
    excellib::BatchPrinter batch_;
};

}  // namespace case1
