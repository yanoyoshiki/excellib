#pragma once
/**
 * case1 — ui_app.hpp
 *
 * メインウィンドウとアプリケーション状態。
 * 各ステージのパネルがこのウィンドウに乗る。
 */
#ifdef _WIN32
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include "ui_theme.hpp"
#include "workflow.hpp"
#include "config.hpp"
#include <string>
#include <vector>
#include <memory>
#include <thread>
#include <atomic>
#include <optional>
#include <cstdint>

namespace case1 {

// ============================================================
//  ステージ定義
// ============================================================
enum class Stage {
    Input    = 0,    // 1. 入力
    Master   = 1,    // 2. マスター読込
    Jobs     = 2,    // 3. ジョブ調整
    Run      = 3,    // 4. 実行・結果
};
constexpr int STAGE_COUNT = 4;

// ============================================================
//  ステージのステータス（インジケータ用）
// ============================================================
enum class StageStatus { Pending, Active, Done, Error };

// ============================================================
//  カスタムメッセージ (ワーカー → UI)
// ============================================================
constexpr UINT WM_APP_PROGRESS    = WM_APP + 1;  // wParam=done, lParam=msg ptr (heap)
constexpr UINT WM_APP_BATCH_DONE  = WM_APP + 2;  // BatchResult* in lParam
constexpr UINT WM_APP_LOG         = WM_APP + 3;  // lParam=msg ptr (heap)

// ============================================================
//  アプリ状態
// ============================================================
struct AppState {
    // ユーザー入力
    std::string  input_string;
    OutputMode   output_mode{OutputMode::PdfIndividual};
    std::string  output_dir;
    std::string  output_pdf_path;
    std::string  printer_name;

    // 解決済み
    std::string  resolved_path;        ///< construct_path → resolve_file 後

    // 各ステージの結果
    std::vector<MasterRecord>          records;
    std::vector<excellib::PrintJob>    jobs;
    std::optional<excellib::BatchResult> last_result;

    // 世代カウンタ — Stage N の commit で N+1 を bump し、
    // 後段は記憶した世代と比較して入力変化を検出する。
    uint32_t input_gen{0};       ///< Stage 1 commit で +1
    uint32_t records_gen{0};     ///< Stage 2 (load_master 成功) で +1
    uint32_t jobs_gen{0};        ///< Stage 3 commit で +1

    // ステージステータス
    StageStatus status[STAGE_COUNT]{
        StageStatus::Active,  StageStatus::Pending,
        StageStatus::Pending, StageStatus::Pending,
    };

    // 実行中フラグ
    std::atomic<bool> running{false};

    Workflow workflow;
};

// ============================================================
//  ステージパネル抽象
// ============================================================
class StagePanel {
public:
    virtual ~StagePanel() = default;
    virtual HWND  hwnd() const = 0;
    virtual void  on_show() {}
    virtual void  on_hide() {}
    virtual void  layout(int width, int height) = 0;
    /// 「進む」が有効か
    virtual bool  can_advance() const { return true; }
    /// 進む時の検証・コミット (false なら進めない)
    virtual bool  commit() { return true; }
    /// やり直し（このステージの状態をクリア）
    virtual void  reset() {}
    /// 子コントロールからの WM_COMMAND
    virtual void  on_command(int id, int notif_code, HWND ctrl) {}
    /// 子コントロールからの WM_NOTIFY
    virtual LRESULT on_notify(NMHDR* hdr) { return 0; }
    /// 子コントロールからのアプリカスタムメッセージ
    virtual void  on_app_message(UINT msg, WPARAM w, LPARAM l) {}
};

// ============================================================
//  メインウィンドウ
// ============================================================
class MainWindow {
public:
    MainWindow(HINSTANCE hInst, AppState& state);
    ~MainWindow();

    HWND create();
    int  run_message_loop();

    // ステージ切替
    void show_stage(Stage s);
    Stage current_stage() const { return current_; }

    void set_status(Stage s, StageStatus st);
    void invalidate_step_bar();
    void refresh_nav();

    // 共通アクセサ
    HINSTANCE hinst()  const { return hInst_; }
    HWND      hwnd()   const { return hwnd_; }
    AppState& state()        { return state_; }
    theme::Fonts& fonts()    { return fonts_; }
    int  dpi() const { return dpi_; }

private:
    static LRESULT CALLBACK wndproc_static(HWND, UINT, WPARAM, LPARAM);
    LRESULT wndproc(UINT, WPARAM, LPARAM);

    void on_create();
    void on_size(int w, int h);
    void on_command(int id, HWND ctrl);

    void create_step_bar();
    void create_stage_panels();
    void create_nav_buttons();
    void update_nav_buttons();

    void on_back();
    void on_redo();
    void on_advance();

    HINSTANCE hInst_{};
    HWND      hwnd_{};
    int       dpi_{96};
    AppState& state_;
    theme::Fonts fonts_;

    // 子ウィンドウ
    HWND step_bar_{};        // カスタム描画
    HWND nav_back_{};
    HWND nav_redo_{};
    HWND nav_next_{};

    std::unique_ptr<StagePanel> panels_[STAGE_COUNT];
    Stage current_{Stage::Input};
};

// ============================================================
//  各ステージのファクトリ
// ============================================================
std::unique_ptr<StagePanel> make_stage_input (MainWindow& app);
std::unique_ptr<StagePanel> make_stage_master(MainWindow& app);
std::unique_ptr<StagePanel> make_stage_jobs  (MainWindow& app);
std::unique_ptr<StagePanel> make_stage_run   (MainWindow& app);

}  // namespace case1
#endif
