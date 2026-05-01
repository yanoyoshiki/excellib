#pragma once
/**
 * excellib/excel_printer.hpp
 *
 * 印刷 / PDF エクスポート API。
 * Windows の Excel COM オートメーションのみを使用。
 * Excel は常にインストール済みを前提とする。
 */

// print_settings.hpp: PageSetup / PrinterOptions / PdfOptions など印刷設定の型
#include "excellib/print_settings.hpp"
// 文字列型 (std::string)
#include <string>
// 動的配列型 (std::vector)
#include <vector>
// std::unique_ptr などスマートポインタ
#include <memory>
// std::function: 関数オブジェクト・ラムダを保持できる型
#include <functional>

namespace excellib {

// PrintProgressCallback: 進捗メッセージを受け取るコールバック関数の型
// std::function<void(const std::string&)> は文字列を受け取って何も返さない関数
using PrintProgressCallback = std::function<void(const std::string& message)>;

// PrintResult: 単一の印刷・PDF 出力操作の結果を格納する構造体
struct PrintResult {
    // success: 処理が成功したか (true = 成功)
    bool        success{false};
    // engine_used: 使用したエンジン名 (例: "Excel")
    std::string engine_used;
    // error_message: エラーがあった場合のメッセージ
    std::string error_message;
    // output_path: PDF 出力の場合の出力先ファイルパス
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
// ExcelPrinter: Excel ファイルの印刷・PDF 出力・ページ設定書き込みを担当するクラス
class ExcelPrinter {
public:
    // コンストラクタ: Windows の COM (Component Object Model) を初期化する
    ExcelPrinter();
    // デストラクタ: COM のクリーンアップを行う
    ~ExcelPrinter();

    // ---- XLSX にページ設定を焼き込む（Excel 不要、ZIP/XML 操作のみ）----
    // XLSX ファイル内の XML を直接編集してページ設定を書き込む
    // Excel がインストールされていない環境でも動作する
    // xlsx_path: 対象の XLSX ファイルパス
    // sheet_name: ページ設定を適用するシートの名前
    // setup: 適用するページ設定
    void apply_page_setup(const std::string& xlsx_path,
                          const std::string& sheet_name,
                          const PageSetup&   setup);

    // ---- PDF 出力 ----
    // Excel COM 経由で PDF を生成する (Windows かつ Excel インストール済みが必要)
    // xlsx_path: 入力 XLSX ファイルパス
    // opts: PDF 出力オプション (output_path が必須)
    // sheet_name: 出力するシートの名前 (空文字列 = アクティブシート)
    // setup: ページ設定 (nullptr = ファイルに焼き込まれた設定を使用)
    // cb: 進捗コールバック (nullptr = 進捗通知なし)
    PrintResult to_pdf(const std::string&     xlsx_path,
                       const PdfOptions&       opts,
                       const std::string&      sheet_name = "",
                       const PageSetup*        setup      = nullptr,
                       PrintProgressCallback   cb         = nullptr);

    // ---- プリンター送信 ----
    // Excel COM 経由でプリンターに印刷ジョブを送信する (Windows 限定)
    // xlsx_path: 入力 XLSX ファイルパス
    // printer_opts: プリンター設定 (printer_name が空 = デフォルトプリンター)
    // sheet_name: 印刷するシートの名前 (空文字列 = アクティブシート)
    // setup: ページ設定 (nullptr = ファイルに焼き込まれた設定を使用)
    // cb: 進捗コールバック
    PrintResult print(const std::string&     xlsx_path,
                      const PrinterOptions&   printer_opts = {},
                      const std::string&      sheet_name   = "",
                      const PageSetup*        setup        = nullptr,
                      PrintProgressCallback   cb           = nullptr);

    // ---- プリンター一覧 ----
    // システムにインストールされているプリンターの名前一覧を返す (Windows 限定)
    std::vector<std::string> list_printers()   const;
    // システムのデフォルトプリンター名を返す (Windows 限定)
    std::string              default_printer() const;

    // 進捗コールバックをセットする
    void set_progress_callback(PrintProgressCallback cb) { progress_cb_ = cb; }

    // 内部公開（batch_printer.cpp / xlsx_page_setup.cpp から使用）
    // XLSX ファイルにページ設定を XML レベルで書き込む静的メソッド
    // static メンバ関数なのでインスタンスなしで呼べる
    static void write_page_setup_to_xlsx(const std::string& xlsx_path,
                                          const std::string& sheet_name,
                                          const PageSetup&   setup);
private:
    // 進捗コールバックを保持するメンバ変数
    PrintProgressCallback progress_cb_;
    // ログメッセージを progress_cb_ に送るプライベートヘルパー
    void log(const std::string& msg);

    // Excel COM 経由で PDF を出力するプライベートメソッド
    PrintResult excel_to_pdf (const std::string&, const PdfOptions&,
                               const std::string&, const PageSetup*);
    // Excel COM 経由でプリンターに印刷するプライベートメソッド
    PrintResult excel_print  (const std::string&, const PrinterOptions&,
                               const std::string&, const PageSetup*);

    // COMGuard: COM の初期化・解放を RAII で管理する内部クラス
    // RAII = Resource Acquisition Is Initialization (リソースの確保と解放を自動化)
    struct COMGuard;
    // unique_ptr で COMGuard のライフタイムを管理する
    std::unique_ptr<COMGuard> com_;
};

} // namespace excellib
