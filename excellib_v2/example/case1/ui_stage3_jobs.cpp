#ifdef _WIN32
#include "ui_stages.hpp"
#include "workflow.hpp"
#include "config.hpp"
#include <commctrl.h>

namespace case1 {

namespace {

constexpr int ID_LIST       = 400;
constexpr int ID_SELECT_ALL = 401;
constexpr int ID_DESELECT   = 402;
constexpr int ID_LANDSCAPE  = 403;

class StageJobs : public ui::PanelBase {
public:
    StageJobs(MainWindow& app) : PanelBase(app) {
        make_panel(L"Case1StageJobs");

        title_ = make_label(L"ステップ 3 — 印刷するジョブを選びます",
                             app_.fonts().title);
        desc_  = make_label(L"対象を外したい行はチェックを外してください。"
                             L"ダブルクリックで個別の印刷列範囲を編集できます。",
                             app_.fonts().body);

        select_all_  = make_button(ID_SELECT_ALL, theme::BtnKind::Neutral, L"全部チェック");
        deselect_all_= make_button(ID_DESELECT,   theme::BtnKind::Neutral, L"全部はずす");

        landscape_ = CreateWindowExW(0, L"BUTTON",
            L"横向き (Landscape) で印刷",
            WS_CHILD|WS_VISIBLE|WS_TABSTOP|BS_AUTOCHECKBOX,
            0,0,200,30, panel_, (HMENU)(intptr_t)ID_LANDSCAPE,
            app_.hinst(), nullptr);
        SendMessageW(landscape_, WM_SETFONT, (WPARAM)app_.fonts().heading, TRUE);

        list_ = make_listview(LVS_EX_CHECKBOXES);
        ListView_SetExtendedListViewStyle(list_,
            LVS_EX_FULLROWSELECT|LVS_EX_GRIDLINES|LVS_EX_DOUBLEBUFFER|LVS_EX_CHECKBOXES);

        ListView_InsertColumn(list_, 0, _col(L"#",       60));
        ListView_InsertColumn(list_, 1, _col(L"ファイル", 320));
        ListView_InsertColumn(list_, 2, _col(L"シート",   140));
        ListView_InsertColumn(list_, 3, _col(L"列範囲",   120));
        ListView_InsertColumn(list_, 4, _col(L"ラベル",   200));

        summary_ = make_label(L"", app_.fonts().bodyBold);
    }

    void on_show() override {
        // Stage 2 の records が更新されていたら一覧を作り直す
        if (app_.state().records_gen != seen_records_gen_) {
            seen_records_gen_ = app_.state().records_gen;
            // 全部チェックの初期状態に戻す
            for (auto& r : app_.state().records) r.included = true;
        }
        repopulate();
        update_summary();
    }

    void layout(int w, int h) override {
        int pad = theme::scale(theme::Padding, app_.dpi());
        int gap = theme::scale(theme::Gap, app_.dpi());
        int line = theme::scale(28, app_.dpi());
        int x = pad, y = pad;

        SetWindowPos(title_, nullptr, x, y, w-pad*2, theme::scale(36, app_.dpi()), SWP_NOZORDER);
        y += theme::scale(40, app_.dpi());
        SetWindowPos(desc_,  nullptr, x, y, w-pad*2, line*2, SWP_NOZORDER);
        y += line*2 + gap/2;

        // ボタン群
        int btn_h = theme::scale(theme::BtnHeightSmall, app_.dpi());
        int sb_w  = theme::scale(160, app_.dpi());
        SetWindowPos(select_all_,  nullptr, x, y, sb_w, btn_h, SWP_NOZORDER);
        SetWindowPos(deselect_all_,nullptr, x + sb_w + gap, y, sb_w, btn_h, SWP_NOZORDER);
        SetWindowPos(landscape_,   nullptr, x + sb_w*2 + gap*2, y, theme::scale(280, app_.dpi()), btn_h, SWP_NOZORDER);
        y += btn_h + gap;

        // List
        int sum_h = line;
        int lv_h = h - y - sum_h - pad - gap;
        if (lv_h < theme::scale(120, app_.dpi())) lv_h = theme::scale(120, app_.dpi());
        SetWindowPos(list_, nullptr, x, y, w-pad*2, lv_h, SWP_NOZORDER);
        y += lv_h + gap;

        SetWindowPos(summary_, nullptr, x, y, w-pad*2, sum_h, SWP_NOZORDER);
    }

    void on_command(int id, int notif, HWND) override {
        if (id == ID_SELECT_ALL && notif == BN_CLICKED) { check_all(true);  update_summary(); }
        if (id == ID_DESELECT   && notif == BN_CLICKED) { check_all(false); update_summary(); }
        if (id == ID_LANDSCAPE  && notif == BN_CLICKED) { /* 値は commit 時に読む */ }
    }

    LRESULT on_notify(NMHDR* hdr) override {
        if (hdr->idFrom != 0 && hdr->hwndFrom != list_) return 0;
        if (hdr->code == LVN_ITEMCHANGED) {
            auto* nm = (NMLISTVIEW*)hdr;
            if (nm->uChanged & LVIF_STATE)
                update_summary();
        }
        if (hdr->code == NM_DBLCLK) {
            edit_print_cols();
        }
        return 0;
    }

    bool can_advance() const override {
        return checked_count() > 0;
    }

    bool commit() override {
        size_t n = app_.state().records.size();
        for (size_t i = 0; i < n; ++i) {
            app_.state().records[i].included =
                ListView_GetCheckState(list_, (int)i) != 0;
            wchar_t buf[64]{};
            ListView_GetItemText(list_, (int)i, 3, buf, 64);
            std::string val = theme::to_utf8(buf);
            if (val == "(全列)") val.clear();
            app_.state().records[i].print_cols = val;
        }
        bool landscape = SendMessageW(landscape_, BM_GETCHECK, 0, 0) == BST_CHECKED;
        app_.state().jobs = Workflow::records_to_jobs(app_.state().records, landscape);
        ++app_.state().jobs_gen;
        return !app_.state().jobs.empty();
    }

    void reset() override {
        app_.state().jobs.clear();
        repopulate();
        update_summary();
    }

private:
    static LVCOLUMNW _col(const wchar_t* t, int w) {
        LVCOLUMNW c{}; c.mask = LVCF_TEXT|LVCF_WIDTH;
        c.cx = w; c.pszText = const_cast<wchar_t*>(t);
        return c;
    }

    void repopulate() {
        ListView_DeleteAllItems(list_);
        int i = 0;
        for (auto& r : app_.state().records) {
            std::wstring n = std::to_wstring(++i);
            LVITEMW it{}; it.mask = LVIF_TEXT; it.iItem = i-1;
            it.pszText = (LPWSTR)n.c_str();
            ListView_InsertItem(list_, &it);

            auto wfile = theme::to_wide(r.file_path);
            auto wsh   = theme::to_wide(r.sheet_name);
            auto wcol  = theme::to_wide(r.print_cols);
            auto wlbl  = theme::to_wide(r.label);

            ListView_SetItemText(list_, i-1, 1, (LPWSTR)wfile.c_str());
            ListView_SetItemText(list_, i-1, 2, (LPWSTR)wsh.c_str());
            ListView_SetItemText(list_, i-1, 3,
                                  wcol.empty() ? (LPWSTR)L"(全列)" : (LPWSTR)wcol.c_str());
            ListView_SetItemText(list_, i-1, 4, (LPWSTR)wlbl.c_str());

            ListView_SetCheckState(list_, i-1, r.included);
        }
    }

    void check_all(bool checked) {
        int n = ListView_GetItemCount(list_);
        for (int i = 0; i < n; ++i)
            ListView_SetCheckState(list_, i, checked);
    }

    size_t checked_count() const {
        int n = ListView_GetItemCount(list_);
        size_t c = 0;
        for (int i = 0; i < n; ++i)
            if (ListView_GetCheckState(list_, i)) ++c;
        return c;
    }

    void update_summary() {
        int total = ListView_GetItemCount(list_);
        size_t c = checked_count();
        wchar_t buf[128];
        swprintf(buf, 128, L"  選択中:  %zu / %d 件", c, total);
        SetWindowTextW(summary_, buf);
        app_.refresh_nav();
    }

    void edit_print_cols() {
        int sel = ListView_GetNextItem(list_, -1, LVNI_SELECTED);
        if (sel < 0) return;
        wchar_t cur[64]{};
        ListView_GetItemText(list_, sel, 3, cur, 64);

        std::wstring v(cur);
        if (v == L"(全列)") v.clear();

        std::wstring out;
        if (run_edit_popup(L"列範囲を入力 (例: A:D / 空欄なら全列)", v, out)) {
            if (out.empty()) out = L"(全列)";
            ListView_SetItemText(list_, sel, 3, (LPWSTR)out.c_str());
        }
    }

    // -------- 自前のミニ入力ポップアップ --------
    // モーダルな小窓を 1 つ作って、OK / Cancel / Enter / Esc を受ける。
    // DLGTEMPLATE を手組みする代わりに普通の WS_POPUP ウィンドウ。
    static constexpr int IDOK_BTN     = 0x7001;
    static constexpr int IDCANCEL_BTN = 0x7002;
    static constexpr int ID_EDIT_FLD  = 0x7003;

    struct EditPopupCtx {
        std::wstring* out{};
        bool          accepted{false};
    };

    static LRESULT CALLBACK popup_proc(HWND h, UINT m, WPARAM w, LPARAM l) {
        auto* ctx = reinterpret_cast<EditPopupCtx*>(GetWindowLongPtrW(h, GWLP_USERDATA));
        switch (m) {
            case WM_NCCREATE: {
                auto cs = (CREATESTRUCTW*)l;
                SetWindowLongPtrW(h, GWLP_USERDATA, (LONG_PTR)cs->lpCreateParams);
                return TRUE;
            }
            case WM_CLOSE: DestroyWindow(h); return 0;
            case WM_COMMAND: {
                int id = LOWORD(w);
                if (id == IDOK_BTN || id == IDCANCEL_BTN) {
                    if (id == IDOK_BTN && ctx) {
                        wchar_t buf[256]{};
                        GetDlgItemTextW(h, ID_EDIT_FLD, buf, 256);
                        *ctx->out = buf;
                        ctx->accepted = true;
                    }
                    DestroyWindow(h);
                }
                return 0;
            }
        }
        return DefWindowProcW(h, m, w, l);
    }

    bool run_edit_popup(const wchar_t* prompt, const std::wstring& initial,
                         std::wstring& out) {
        out = initial;
        EditPopupCtx ctx{ &out, false };

        static bool reg = false;
        const wchar_t* cls = L"Case1EditPopup";
        if (!reg) {
            WNDCLASSW wc{};
            wc.lpfnWndProc   = popup_proc;
            wc.hInstance     = app_.hinst();
            wc.hCursor       = LoadCursor(nullptr, IDC_ARROW);
            wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
            wc.lpszClassName = cls;
            RegisterClassW(&wc);
            reg = true;
        }

        // 親ウィンドウの中央に配置
        RECT pr; GetWindowRect(app_.hwnd(), &pr);
        int W = theme::scale(420, app_.dpi());
        int H = theme::scale(150, app_.dpi());
        int x = pr.left + ((pr.right - pr.left) - W) / 2;
        int y = pr.top  + ((pr.bottom - pr.top) - H) / 2;

        HWND popup = CreateWindowExW(WS_EX_DLGMODALFRAME|WS_EX_TOPMOST,
            cls, L"列範囲の編集",
            WS_POPUP|WS_CAPTION|WS_SYSMENU,
            x, y, W, H, app_.hwnd(), nullptr, app_.hinst(), &ctx);
        if (!popup) return false;

        int pad = theme::scale(12, app_.dpi());
        int line = theme::scale(28, app_.dpi());
        HWND lbl = CreateWindowExW(0, L"STATIC", prompt,
            WS_CHILD|WS_VISIBLE|SS_LEFT,
            pad, pad, W - pad*2, line, popup, nullptr, app_.hinst(), nullptr);
        SendMessageW(lbl, WM_SETFONT, (WPARAM)app_.fonts().body, TRUE);

        HWND ed = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", initial.c_str(),
            WS_CHILD|WS_VISIBLE|WS_TABSTOP|ES_AUTOHSCROLL,
            pad, pad + line + 4, W - pad*2, theme::scale(28, app_.dpi()),
            popup, (HMENU)(intptr_t)ID_EDIT_FLD, app_.hinst(), nullptr);
        SendMessageW(ed, WM_SETFONT, (WPARAM)app_.fonts().large_input, TRUE);

        int by = pad + line + 4 + theme::scale(34, app_.dpi());
        int bw = theme::scale(96, app_.dpi());
        int bh = theme::scale(32, app_.dpi());
        HWND ok = CreateWindowExW(0, L"BUTTON", L"OK",
            WS_CHILD|WS_VISIBLE|WS_TABSTOP|BS_DEFPUSHBUTTON,
            W - pad - bw*2 - 8, by, bw, bh, popup,
            (HMENU)(intptr_t)IDOK_BTN, app_.hinst(), nullptr);
        SendMessageW(ok, WM_SETFONT, (WPARAM)app_.fonts().body, TRUE);
        HWND cn = CreateWindowExW(0, L"BUTTON", L"キャンセル",
            WS_CHILD|WS_VISIBLE|WS_TABSTOP,
            W - pad - bw, by, bw, bh, popup,
            (HMENU)(intptr_t)IDCANCEL_BTN, app_.hinst(), nullptr);
        SendMessageW(cn, WM_SETFONT, (WPARAM)app_.fonts().body, TRUE);

        EnableWindow(app_.hwnd(), FALSE);
        ShowWindow(popup, SW_SHOW);
        SetFocus(ed);
        SendMessageW(ed, EM_SETSEL, 0, -1);

        // モーダルループ
        MSG msg;
        while (IsWindow(popup) && GetMessageW(&msg, nullptr, 0, 0)) {
            // Enter/Esc を OK/Cancel に変換
            if (msg.message == WM_KEYDOWN && IsChild(popup, msg.hwnd)) {
                if (msg.wParam == VK_RETURN) {
                    SendMessageW(popup, WM_COMMAND, MAKEWPARAM(IDOK_BTN, BN_CLICKED), 0);
                    continue;
                }
                if (msg.wParam == VK_ESCAPE) {
                    SendMessageW(popup, WM_COMMAND, MAKEWPARAM(IDCANCEL_BTN, BN_CLICKED), 0);
                    continue;
                }
            }
            if (!IsDialogMessageW(popup, &msg)) {
                TranslateMessage(&msg);
                DispatchMessageW(&msg);
            }
        }
        EnableWindow(app_.hwnd(), TRUE);
        SetForegroundWindow(app_.hwnd());
        return ctx.accepted;
    }

    HWND title_{}, desc_{};
    HWND select_all_{}, deselect_all_{}, landscape_{};
    HWND list_{}, summary_{};
    uint32_t seen_records_gen_{0};
};

}  // namespace

std::unique_ptr<StagePanel> make_stage_jobs(MainWindow& app) {
    return std::make_unique<StageJobs>(app);
}

}  // namespace case1
#endif
