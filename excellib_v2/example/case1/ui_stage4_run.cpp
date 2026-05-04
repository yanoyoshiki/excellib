#ifdef _WIN32
#include "ui_stages.hpp"
#include "workflow.hpp"
#include "config.hpp"
#include <commctrl.h>
#include <thread>
#include <atomic>
#include <memory>
#include <string>
#include <vector>

namespace case1 {

namespace {

constexpr int ID_RUN          = 500;
constexpr int ID_CANCEL       = 501;
constexpr int ID_RETRY_FAILED = 502;
constexpr int ID_RESTART      = 503;
constexpr int ID_LOG          = 510;
constexpr int ID_RESULT_LIST  = 511;

class StageRun : public ui::PanelBase {
public:
    StageRun(MainWindow& app) : PanelBase(app) {
        make_panel(L"Case1StageRun");

        title_ = make_label(L"ステップ 4 — 実行 & 結果", app_.fonts().title);

        run_btn_    = make_button(ID_RUN,          theme::BtnKind::Success, L"▶  実行");
        cancel_btn_ = make_button(ID_CANCEL,       theme::BtnKind::Danger,  L"■  キャンセル");
        retry_btn_  = make_button(ID_RETRY_FAILED, theme::BtnKind::Warning, L"↻  失敗のみ再実行");
        restart_btn_= make_button(ID_RESTART,      theme::BtnKind::Neutral, L"はじめからやり直す");

        progress_ = make_progress();
        SendMessageW(progress_, PBM_SETRANGE32, 0, 100);

        current_label_ = make_label(L"", app_.fonts().heading);
        summary_       = make_label(L"", app_.fonts().bodyBold);

        // ログ ListBox
        log_ = CreateWindowExW(WS_EX_CLIENTEDGE, L"LISTBOX", L"",
            WS_CHILD|WS_VISIBLE|WS_VSCROLL|LBS_NOINTEGRALHEIGHT|LBS_DISABLENOSCROLL,
            0,0,300,200, panel_, (HMENU)(intptr_t)ID_LOG, app_.hinst(), nullptr);
        SendMessageW(log_, WM_SETFONT, (WPARAM)app_.fonts().mono, TRUE);

        // 結果 ListView
        result_list_ = make_listview();
        ListView_InsertColumn(result_list_, 0, _col(L"#",       60));
        ListView_InsertColumn(result_list_, 1, _col(L"状態",    100));
        ListView_InsertColumn(result_list_, 2, _col(L"ファイル",320));
        ListView_InsertColumn(result_list_, 3, _col(L"シート",  140));
        ListView_InsertColumn(result_list_, 4, _col(L"メッセージ", 260));

        // 初期状態は結果なし → 結果系コントロールは隠す
        show_result_controls(false);
        EnableWindow(cancel_btn_, FALSE);
    }

    ~StageRun() override {
        // 実行中ならキャンセルを要求して join する。
        // 未 join のままウィンドウが破棄されると WM_APP_BATCH_DONE の
        // PostMessage 先が無効ハンドルになり UB になる。
        if (worker_.joinable()) {
            app_.state().workflow.request_cancel();
            worker_.join();
        }
    }

    void on_show() override {
        // jobs が変わっていたら結果表示をクリア (前回の結果は古い)
        if (app_.state().jobs_gen != seen_jobs_gen_) {
            seen_jobs_gen_ = app_.state().jobs_gen;
            ListView_DeleteAllItems(result_list_);
            SendMessageW(log_, LB_RESETCONTENT, 0, 0);
            app_.state().last_result.reset();
            show_result_controls(false);
        }
        wchar_t buf[128];
        swprintf(buf, 128, L"  実行待ち:  %zu 件のジョブ",
                 app_.state().jobs.size());
        SetWindowTextW(summary_, buf);
        SetWindowTextW(current_label_, L"");
        SendMessageW(progress_, PBM_SETPOS, 0, 0);
    }

    void layout(int w, int h) override {
        int pad = theme::scale(theme::Padding, app_.dpi());
        int gap = theme::scale(theme::Gap, app_.dpi());
        int line = theme::scale(28, app_.dpi());
        int x = pad, y = pad;

        SetWindowPos(title_, nullptr, x, y, w-pad*2, theme::scale(36, app_.dpi()), SWP_NOZORDER);
        y += theme::scale(40, app_.dpi());

        // ボタン行 (実行 / キャンセル)
        int btn_h = theme::scale(theme::BtnHeightTall, app_.dpi());
        int run_w = theme::scale(180, app_.dpi());
        int can_w = theme::scale(220, app_.dpi());
        SetWindowPos(run_btn_,    nullptr, x, y, run_w, btn_h, SWP_NOZORDER);
        SetWindowPos(cancel_btn_, nullptr, x + run_w + gap, y, can_w, btn_h, SWP_NOZORDER);
        // 結果ボタン (右側)
        int re_w = theme::scale(220, app_.dpi());
        int rs_w = theme::scale(220, app_.dpi());
        SetWindowPos(retry_btn_,  nullptr, w - pad - re_w - rs_w - gap, y, re_w, btn_h, SWP_NOZORDER);
        SetWindowPos(restart_btn_,nullptr, w - pad - rs_w, y, rs_w, btn_h, SWP_NOZORDER);
        y += btn_h + gap;

        // プログレスバー
        int pb_h = theme::scale(28, app_.dpi());
        SetWindowPos(progress_, nullptr, x, y, w-pad*2, pb_h, SWP_NOZORDER);
        y += pb_h + gap/2;

        SetWindowPos(current_label_, nullptr, x, y, w-pad*2, line, SWP_NOZORDER);
        y += line;
        SetWindowPos(summary_, nullptr, x, y, w-pad*2, line, SWP_NOZORDER);
        y += line + gap;

        // ログとリストを上下に並べる (1:1)
        int rest_h = h - y - pad;
        if (rest_h < theme::scale(160, app_.dpi())) rest_h = theme::scale(160, app_.dpi());
        int half = rest_h / 2 - gap/2;

        SetWindowPos(log_, nullptr, x, y, w-pad*2, half, SWP_NOZORDER);
        y += half + gap;
        SetWindowPos(result_list_, nullptr, x, y, w-pad*2, half, SWP_NOZORDER);
    }

    void on_command(int id, int notif, HWND) override {
        if (notif != BN_CLICKED) return;
        switch (id) {
            case ID_RUN:          start_run(app_.state().jobs);            break;
            case ID_CANCEL:       app_.state().workflow.request_cancel();
                                  log_line(L"  ⏸  キャンセル要求しました…");  break;
            case ID_RETRY_FAILED: retry_failed();                          break;
            case ID_RESTART:      app_.show_stage(Stage::Input);           break;
        }
    }

    void on_app_message(UINT msg, WPARAM w, LPARAM l) override {
        if (msg == WM_APP_PROGRESS) {
            // wParam = done, lParam = std::wstring*
            auto* ws = reinterpret_cast<std::wstring*>(l);
            int pct = (int)w;
            SendMessageW(progress_, PBM_SETPOS, pct, 0);
            if (ws) {
                SetWindowTextW(current_label_, ws->c_str());
                delete ws;
            }
        }
        else if (msg == WM_APP_LOG) {
            auto* ws = reinterpret_cast<std::wstring*>(l);
            if (ws) { log_line(ws->c_str()); delete ws; }
        }
        else if (msg == WM_APP_BATCH_DONE) {
            // l = BatchResult* (heap)
            auto* r = reinterpret_cast<excellib::BatchResult*>(l);
            on_batch_done(*r);
            delete r;
        }
    }

    bool can_advance() const override { return true; }   // ステージ4 では Next = 閉じる

private:
    static LVCOLUMNW _col(const wchar_t* t, int w) {
        LVCOLUMNW c{}; c.mask = LVCF_TEXT|LVCF_WIDTH;
        c.cx = w; c.pszText = const_cast<wchar_t*>(t);
        return c;
    }

    void show_result_controls(bool show) {
        ShowWindow(retry_btn_,   show ? SW_SHOW : SW_HIDE);
        ShowWindow(restart_btn_, show ? SW_SHOW : SW_HIDE);
        ShowWindow(result_list_, show ? SW_SHOW : SW_HIDE);
    }

    void log_line(const wchar_t* line) {
        int idx = (int)SendMessageW(log_, LB_ADDSTRING, 0, (LPARAM)line);
        SendMessageW(log_, LB_SETTOPINDEX, idx, 0);
    }

    void start_run(const std::vector<excellib::PrintJob>& jobs) {
        if (app_.state().running.load()) return;
        if (jobs.empty()) { log_line(L"  ⚠  実行するジョブがありません"); return; }

        SendMessageW(log_, LB_RESETCONTENT, 0, 0);
        ListView_DeleteAllItems(result_list_);
        SendMessageW(progress_, PBM_SETPOS, 0, 0);
        SetWindowTextW(current_label_, L"  準備中…");
        EnableWindow(run_btn_, FALSE);
        EnableWindow(cancel_btn_, TRUE);
        show_result_controls(false);

        app_.state().running.store(true);
        app_.state().workflow.reset_cancel();

        HWND hwnd = app_.hwnd();
        std::vector<excellib::PrintJob> jobs_copy = jobs;

        // 実行コンテキストをコピーしてワーカースレッドへ
        RunContext ctx;
        ctx.mode             = app_.state().output_mode;
        ctx.output_dir       = app_.state().output_dir;
        ctx.output_pdf_path  = app_.state().output_pdf_path;
        ctx.printer_name     = app_.state().printer_name;
        size_t total = jobs_copy.size();

        ctx.on_progress = [hwnd, total](size_t done, size_t /*total*/, const std::string& msg) {
            int pct = total>0 ? int((done*100)/total) : 0;
            auto* ws = new std::wstring(theme::to_wide(msg));
            PostMessageW(hwnd, WM_APP_PROGRESS, (WPARAM)pct, (LPARAM)ws);
            auto* lg = new std::wstring(L"  • [" + std::to_wstring(done) + L"/" +
                                        std::to_wstring(total) + L"] " + theme::to_wide(msg));
            PostMessageW(hwnd, WM_APP_LOG, 0, (LPARAM)lg);
        };

        // 既に走っているスレッドが join 待ち状態 (再実行) なら片付ける
        if (worker_.joinable()) worker_.join();

        worker_ = std::thread([this, hwnd, jobs_copy, ctx]() mutable {
            auto r = std::make_unique<excellib::BatchResult>(
                app_.state().workflow.run_batch(jobs_copy, ctx));
            PostMessageW(hwnd, WM_APP_BATCH_DONE, 0, (LPARAM)r.release());
        });
        // detach せず join 可能なまま保持 (DTOR or 再実行で安全に処理)
    }

    void on_batch_done(const excellib::BatchResult& r) {
        app_.state().running.store(false);
        app_.state().last_result = r;

        EnableWindow(run_btn_, TRUE);
        EnableWindow(cancel_btn_, FALSE);
        SendMessageW(progress_, PBM_SETPOS, 100, 0);

        // 結果 ListView 更新
        ListView_DeleteAllItems(result_list_);
        for (size_t i = 0; i < r.jobs.size(); ++i) {
            auto& jr = r.jobs[i];
            std::wstring n = std::to_wstring(i+1);
            LVITEMW it{}; it.mask = LVIF_TEXT; it.iItem = (int)i;
            it.pszText = (LPWSTR)n.c_str();
            ListView_InsertItem(result_list_, &it);

            const wchar_t* status =
                jr.cancelled ? L"⏸ キャンセル" :
                jr.success   ? L"✓ 成功" :
                                L"✗ 失敗";
            ListView_SetItemText(result_list_, (int)i, 1, (LPWSTR)status);

            auto wf = theme::to_wide(jr.file_path);
            auto ws = theme::to_wide(jr.sheet_name);
            auto we = theme::to_wide(jr.error_message);
            ListView_SetItemText(result_list_, (int)i, 2, (LPWSTR)wf.c_str());
            ListView_SetItemText(result_list_, (int)i, 3, (LPWSTR)ws.c_str());
            ListView_SetItemText(result_list_, (int)i, 4, (LPWSTR)we.c_str());
        }

        // サマリー
        wchar_t buf[256];
        swprintf(buf, 256, L"  結果:  ✓ 成功 %zu  /  ✗ 失敗 %zu  /  ⏸ キャンセル %zu  /  合計 %zu",
                 r.succeeded, r.failed, r.cancelled_count, r.total);
        SetWindowTextW(summary_, buf);

        SetWindowTextW(current_label_, r.cancelled
            ? L"  ⏸  キャンセルされました"
            : (r.failed == 0 ? L"  ✓  すべて成功しました" : L"  ⚠  失敗があります"));

        log_line(buf);
        show_result_controls(true);

        // ステージステータスを Done/Error に
        bool ok = (r.failed == 0 && !r.cancelled);
        app_.set_status(Stage::Run, ok ? StageStatus::Done : StageStatus::Error);

        // 失敗があれば「失敗のみ再実行」ボタンを有効化
        EnableWindow(retry_btn_, r.failed > 0 || r.cancelled_count > 0);
    }

    void retry_failed() {
        if (!app_.state().last_result) return;
        // BatchPrinter::from_failures で失敗ジョブのみ取り出して再実行
        auto retry = excellib::BatchPrinter::from_failures(*app_.state().last_result);
        // BatchPrinter から元 PrintJob 一覧を抽出する手段がないため、
        // last_result.jobs[].original_job からジョブ一覧を作り直す
        std::vector<excellib::PrintJob> retry_jobs;
        for (auto& jr : app_.state().last_result->jobs)
            if (!jr.success) retry_jobs.push_back(jr.original_job);
        if (retry_jobs.empty()) {
            log_line(L"  再実行対象がありません");
            return;
        }
        log_line((L"  ↻  失敗ジョブ " + std::to_wstring(retry_jobs.size()) +
                  L" 件を再実行します…").c_str());
        start_run(retry_jobs);
    }

    HWND title_{};
    HWND run_btn_{}, cancel_btn_{}, retry_btn_{}, restart_btn_{};
    HWND progress_{}, current_label_{}, summary_{};
    HWND log_{}, result_list_{};
    uint32_t seen_jobs_gen_{0};

    std::thread worker_;
};

}  // namespace

std::unique_ptr<StagePanel> make_stage_run(MainWindow& app) {
    return std::make_unique<StageRun>(app);
}

}  // namespace case1
#endif
