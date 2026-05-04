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
        // チェック状態を records に反映
        size_t n = app_.state().records.size();
        for (size_t i = 0; i < n; ++i) {
            app_.state().records[i].included =
                ListView_GetCheckState(list_, (int)i) != 0;
            // 印刷列範囲 (編集された値) を読み戻す
            wchar_t buf[64]{};
            ListView_GetItemText(list_, (int)i, 3, buf, 64);
            std::string val = theme::to_utf8(buf);
            if (val == "(全列)") val.clear();
            app_.state().records[i].print_cols = val;
        }
        bool landscape = SendMessageW(landscape_, BM_GETCHECK, 0, 0) == BST_CHECKED;
        app_.state().jobs = Workflow::records_to_jobs(app_.state().records, landscape);
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

        // 簡易ダイアログ: InputBox 風
        std::wstring v(cur);
        if (v == L"(全列)") v.clear();

        // 自前ダイアログ生成
        InputDialogState st{ app_.hwnd(), L"列範囲を入力", L"例: A:D / 空欄なら全列",
                              v };
        if (show_input_dialog(st)) {
            std::wstring out = st.value;
            if (out.empty()) out = L"(全列)";
            ListView_SetItemText(list_, sel, 3, (LPWSTR)out.c_str());
        }
    }

    // -------- ミニ入力ダイアログ --------
    struct InputDialogState {
        HWND parent; const wchar_t* title; const wchar_t* prompt;
        std::wstring value;
    };

    static INT_PTR CALLBACK input_dlg_proc(HWND h, UINT m, WPARAM w, LPARAM l) {
        InputDialogState* st = nullptr;
        if (m == WM_INITDIALOG) {
            st = (InputDialogState*)l;
            SetWindowLongPtrW(h, GWLP_USERDATA, (LONG_PTR)st);
            SetWindowTextW(h, st->title);
            SetDlgItemTextW(h, 1001, st->prompt);
            SetDlgItemTextW(h, 1002, st->value.c_str());
            SetFocus(GetDlgItem(h, 1002));
            return FALSE;
        }
        st = (InputDialogState*)GetWindowLongPtrW(h, GWLP_USERDATA);
        if (m == WM_COMMAND) {
            int id = LOWORD(w);
            if (id == IDOK) {
                wchar_t buf[256]{};
                GetDlgItemTextW(h, 1002, buf, 256);
                if (st) st->value = buf;
                EndDialog(h, IDOK);
                return TRUE;
            }
            if (id == IDCANCEL) { EndDialog(h, IDCANCEL); return TRUE; }
        }
        return FALSE;
    }

    bool show_input_dialog(InputDialogState& st) {
        // メモリ上に DLGTEMPLATE を構築
        struct DlgItem {
            DLGITEMTEMPLATE t;
            // 続き: クラス/タイトル/データを WORD 配列で
        };
        // 簡易版: リソース不要なダイアログを動的生成
        // -- WORD 配列で組み立て --
        std::vector<WORD> tpl;
        auto push_w = [&](WORD v){ tpl.push_back(v); };
        auto push_dw = [&](DWORD v){ tpl.push_back(LOWORD(v)); tpl.push_back(HIWORD(v)); };
        auto push_str = [&](const wchar_t* s){
            for (; *s; ++s) tpl.push_back((WORD)*s);
            tpl.push_back(0);
        };
        auto align_dw = [&](){
            while (tpl.size() % 2) tpl.push_back(0);
        };

        // DLGTEMPLATE
        push_dw(DS_SETFONT|DS_CENTER|WS_POPUP|WS_CAPTION|WS_SYSMENU);  // style
        push_dw(0); // dwExtendedStyle
        push_w(4);  // 4 controls
        push_w(0); push_w(0); // x,y
        push_w(220); push_w(80); // cx,cy
        push_w(0); // menu
        push_w(0); // class
        push_str(L""); // title
        push_w(9); push_str(L"Segoe UI"); // font

        // 1. STATIC prompt
        align_dw();
        push_dw(WS_CHILD|WS_VISIBLE|SS_LEFT);
        push_dw(0);
        push_w(8); push_w(8); push_w(204); push_w(20);
        push_w(1001);
        push_w(0xFFFF); push_w(0x0082); // STATIC class
        push_str(L"");
        push_w(0);

        // 2. EDIT input
        align_dw();
        push_dw(WS_CHILD|WS_VISIBLE|WS_TABSTOP|WS_BORDER|ES_AUTOHSCROLL);
        push_dw(0);
        push_w(8); push_w(30); push_w(204); push_w(14);
        push_w(1002);
        push_w(0xFFFF); push_w(0x0081); // EDIT class
        push_str(L"");
        push_w(0);

        // 3. OK button
        align_dw();
        push_dw(WS_CHILD|WS_VISIBLE|WS_TABSTOP|BS_DEFPUSHBUTTON);
        push_dw(0);
        push_w(108); push_w(54); push_w(50); push_w(18);
        push_w(IDOK);
        push_w(0xFFFF); push_w(0x0080); // BUTTON class
        push_str(L"OK");
        push_w(0);

        // 4. Cancel button
        align_dw();
        push_dw(WS_CHILD|WS_VISIBLE|WS_TABSTOP);
        push_dw(0);
        push_w(162); push_w(54); push_w(50); push_w(18);
        push_w(IDCANCEL);
        push_w(0xFFFF); push_w(0x0080);
        push_str(L"キャンセル");
        push_w(0);

        INT_PTR r = DialogBoxIndirectParamW(GetModuleHandleW(nullptr),
            (LPCDLGTEMPLATE)tpl.data(), st.parent, input_dlg_proc,
            (LPARAM)&st);
        return r == IDOK;
    }

    HWND title_{}, desc_{};
    HWND select_all_{}, deselect_all_{}, landscape_{};
    HWND list_{}, summary_{};
};

}  // namespace

std::unique_ptr<StagePanel> make_stage_jobs(MainWindow& app) {
    return std::make_unique<StageJobs>(app);
}

}  // namespace case1
#endif
