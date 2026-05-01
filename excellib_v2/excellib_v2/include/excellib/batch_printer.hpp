#pragma once
/**
 * excellib/batch_printer.hpp
 *
 * 複数ファイル・複数シートを Excel 1インスタンスで一括印刷/PDF化。
 * Excel は常にインストール済みを前提とする。
 */

// excel_printer.hpp: ExcelPrinter / PrintResult / PrinterOptions / PdfOptions など
#include "excellib/excel_printer.hpp"
// 文字列型 (std::string)
#include <string>
// 動的配列型 (std::vector)
#include <vector>
// std::function: 関数オブジェクト・ラムダを保持できる型
#include <functional>
// std::optional: 値があるかもしれない型
#include <optional>

namespace excellib {

// ---- 1件の印刷ジョブ ----
// PrintJob: 1つの印刷・PDF 化タスクを表す構造体
struct PrintJob {
    // file_path: 印刷対象の XLSX ファイルパス
    std::string              file_path;
    // sheet_name: 印刷するシートの名前 (空文字列 = 先頭シート)
    std::string              sheet_name;   ///< 空 = 先頭シート
    // setup: ページ設定 (nullopt = ファイルに焼き込まれた設定を使用)
    std::optional<PageSetup> setup;        ///< nullopt = ファイルの設定を使用
    // label: 進捗表示や PDF ファイル名に使う識別ラベル
    std::string              label;        ///< 進捗表示・PDF ファイル名用

    // make(): ファクトリ関数 — PrintJob オブジェクトを簡潔に作るためのヘルパー
    // static なのでインスタンスなしで PrintJob::make(...) と呼べる
    // std::move(ps) でページ設定を移動(コピーを避ける)して渡す
    static PrintJob make(const std::string& file,
                         const std::string& sheet = "",
                         std::optional<PageSetup> ps = std::nullopt,
                         const std::string& label = "") {
        return {file, sheet, std::move(ps), label};
    }
};

// ---- 結果 ----
// JobResult: 1件の印刷ジョブの実行結果を格納する構造体
struct JobResult {
    // job_index: ジョブリスト内のインデックス (0始まり)
    size_t      job_index{0};
    // file_path: 処理したファイルのパス
    std::string file_path;
    // sheet_name: 処理したシートの名前
    std::string sheet_name;
    // success: ジョブが成功したか (true = 成功)
    bool        success{false};
    // error_message: 失敗した場合のエラーメッセージ
    std::string error_message;
    // elapsed_ms: このジョブの処理時間 (ミリ秒)
    double      elapsed_ms{0.0};
};

// BatchResult: 複数ジョブをまとめて処理した結果全体を格納する構造体
struct BatchResult {
    // jobs: 各ジョブの結果のリスト
    std::vector<JobResult> jobs;
    // total: 全ジョブ数
    size_t      total{0};
    // succeeded: 成功したジョブ数
    size_t      succeeded{0};
    // failed: 失敗したジョブ数
    size_t      failed{0};
    // total_elapsed_ms: 全体の処理時間 (ミリ秒)
    double      total_elapsed_ms{0.0};
    // output_path: PDF 結合出力の場合の出力先ファイルパス
    std::string output_path;

    // all_success(): すべてのジョブが成功したかどうか
    bool all_success() const { return failed == 0; }
    // failures(): 失敗したジョブの結果へのポインタ一覧を返す
    // const JobResult* はポインタ経由で const 参照する (コピーを避ける)
    std::vector<const JobResult*> failures() const {
        std::vector<const JobResult*> v;
        // 全ジョブを走査し、失敗したものだけポインタをベクタに追加
        for (auto& j : jobs) if (!j.success) v.push_back(&j);
        return v;
    }
};

// BatchProgressCallback: バッチ処理の進捗を通知するコールバック関数の型
// done: 完了済みジョブ数, total: 全ジョブ数, job: 現在のジョブ, message: 進捗メッセージ
using BatchProgressCallback = std::function<void(
    size_t done, size_t total,
    const PrintJob& job,
    const std::string& message
)>;

// ---- BatchPrinter ----
// BatchPrinter: 複数の印刷ジョブを Excel 1インスタンスで一括処理するクラス
// Excel を1回だけ起動して全ジョブを処理するため、個別処理より高速
class BatchPrinter {
public:
    // コンストラクタ・デストラクタはデフォルト実装を使用
    BatchPrinter();
    ~BatchPrinter();

    // ジョブを1件追加してメソッドチェーン用に自身への参照を返す
    // 戻り値が BatchPrinter& なので .add().add().add() と連鎖して書ける
    BatchPrinter& add(const PrintJob& job);
    // ファイルパス・シート名・ページ設定・ラベルを個別に指定してジョブを追加
    BatchPrinter& add(const std::string& file,
                      const std::string& sheet = "",
                      std::optional<PageSetup> setup = std::nullopt,
                      const std::string& label = "");
    // std::vector<PrintJob> をまとめて追加する
    BatchPrinter& add_all(const std::vector<PrintJob>& jobs);

    // 登録済みジョブ数を返す
    size_t job_count() const { return jobs_.size(); }
    // 全ジョブをクリアする
    void   clear()           { jobs_.clear(); }

    /**
     * プリンターに一括送信。
     * Excel を1回起動して全ジョブを処理する。
     */
    // print_all(): 全ジョブをプリンターに一括送信する
    // printer_opts: プリンター設定 (printer_name が空 = デフォルトプリンター)
    // cb: 進捗コールバック
    BatchResult print_all(
        const PrinterOptions&   printer_opts = {},
        BatchProgressCallback   cb = nullptr
    );

    /**
     * 全ジョブを1つの PDF に結合して出力。
     * Excel で全シートを一時 Workbook にコピーし ExportAsFixedFormat で一括出力。
     */
    // to_pdf_merged(): 全ジョブのシートを1つの PDF ファイルにまとめて出力する
    // output_path: 出力先 PDF ファイルパス
    // cb: 進捗コールバック
    BatchResult to_pdf_merged(
        const std::string&    output_path,
        BatchProgressCallback cb = nullptr
    );

    /**
     * ジョブごとに個別 PDF を出力。
     * output_dir/{label}_{index}.pdf として保存。
     */
    // to_pdf_individual(): ジョブごとに別々の PDF ファイルを出力する
    // output_dir: PDF ファイルを保存するディレクトリパス
    // cb: 進捗コールバック
    BatchResult to_pdf_individual(
        const std::string&    output_dir,
        BatchProgressCallback cb = nullptr
    );

    // 進捗コールバックをセットする
    void set_progress_callback(BatchProgressCallback cb) { cb_ = cb; }

private:
    // 登録されたジョブのリスト
    std::vector<PrintJob> jobs_;
    // 進捗コールバック関数
    BatchProgressCallback cb_;

    // 内部ログ出力ヘルパー: done/total/job/msg を cb_ に渡す
    void log(size_t done, const PrintJob& job, const std::string& msg);

    // ジョブの PDF ファイル名を生成する静的ヘルパー
    static std::string job_pdf_name(const PrintJob& job, size_t index);
    // 相対パスを絶対パスに変換する静的ヘルパー (Windows: GetFullPathNameA を使用)
    static std::string abs_path(const std::string& p);
    // ファイルパスから拡張子なしのファイル名 (stem) を取得する静的ヘルパー
    static std::string stem(const std::string& path);
};

} // namespace excellib
