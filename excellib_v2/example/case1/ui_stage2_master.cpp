#ifdef _WIN32
#include "ui_stages.hpp"
#include "workflow.hpp"
#include "config.hpp"
#include <commctrl.h>

namespace case1 {

namespace {

constexpr int ID_LOAD = 300;

class StageMaster : public ui::PanelBase {
public:
    StageMaster(MainWindow& app) : PanelBase(app) {
        make_panel(L"Case1StageMaster");

        title_   = make_label(L"ステップ 2 — マスターファイルを読み込みます",
                              app_.fonts().title);
        path_lbl_ = make_label(L"対象ファイル", app_.fonts().heading);
        path_     = make_label(L"", app_.fonts().mono);

        load_btn_ = make_button(ID_LOAD, theme::BtnKind::Primary, L"📥 読み込む");

        banner_ = make_banner();

        list_ = make_listview();
        insert_col(list_, 0, L"#",       60);
        insert_col(list_, 1, L"ファイル", 320);
        insert_col(list_, 2, L"シート",   140);
        insert_col(list_, 3, L"列範囲",   100);
        insert_col(list_, 4, L"ラベル",   180);

        summary_ = make_label(L"", app_.fonts().bodyBold);
    }

    void on_show() override {
        std::string base = Workflow::construct_path_base(app_.state().input_string);
        std::wstring p = theme::to_wide(base);
        SetWindowTextW(path_, (L"  " + p + L"  (.xlsx → .xls)").c_str());

        // Stage 1 が変わっていたら、表示中のレコードはもう古い → 自動リセット
        if (app_.state().input_gen != seen_input_gen_) {
            reset();
            seen_input_gen_ = app_.state().input_gen;
        }
    }

    void layout(int w, int h) override {
        int pad = theme::scale(theme::Padding, app_.dpi());
        int gap = theme::scale(theme::Gap, app_.dpi());
        int line = theme::scale(28, app_.dpi());
        int x = pad, y = pad;

        SetWindowPos(title_, nullptr, x, y, w-pad*2, theme::scale(36, app_.dpi()), SWP_NOZORDER);
        y += theme::scale(40, app_.dpi());

        SetWindowPos(path_lbl_, nullptr, x, y, w-pad*2, line, SWP_NOZORDER);
        y += line;
        SetWindowPos(path_, nullptr, x, y, w-pad*2, line, SWP_NOZORDER);
        y += line + gap;

        // 読み込みボタン
        int btn_h = theme::scale(theme::BtnHeightTall, app_.dpi());
        int btn_w = theme::scale(220, app_.dpi());
        SetWindowPos(load_btn_, nullptr, x, y, btn_w, btn_h, SWP_NOZORDER);
        y += btn_h + gap;

        // バナー
        int banner_h = theme::scale(48, app_.dpi());
        SetWindowPos(banner_, nullptr, x, y, w-pad*2, banner_h, SWP_NOZORDER);
        if (banner_visible_) y += banner_h + gap;

        // ListView (残り全部)
        int sum_h = line;
        int lv_h = h - y - sum_h - pad - gap;
        if (lv_h < theme::scale(80, app_.dpi())) lv_h = theme::scale(80, app_.dpi());
        SetWindowPos(list_, nullptr, x, y, w-pad*2, lv_h, SWP_NOZORDER);
        y += lv_h + gap;

        SetWindowPos(summary_, nullptr, x, y, w-pad*2, sum_h, SWP_NOZORDER);
    }

    void on_command(int id, int notif, HWND) override {
        if (id == ID_LOAD && notif == BN_CLICKED) load();
    }

    bool can_advance() const override {
        return !app_.state().records.empty();
    }

    void reset() override {
        app_.state().resolved_path.clear();
        app_.state().records.clear();
        ListView_DeleteAllItems(list_);
        SetWindowTextW(summary_, L"");
        banner_visible_ = false;
        ShowWindow(banner_, SW_HIDE);
        layout_dirty();
    }

    LRESULT on_notify(NMHDR* hdr) override {
        if (hdr->idFrom == 0 && hdr->code == NM_CUSTOMDRAW) {
            // バナーは SS_OWNERDRAW なので WM_DRAWITEM で描画される（後述）
        }
        return 0;
    }

    void on_app_message(UINT, WPARAM, LPARAM) override {}

private:
    static LVCOLUMNW _col(const wchar_t* t, int w) {
        LVCOLUMNW c{};
        c.mask = LVCF_TEXT|LVCF_WIDTH;
        c.cx = w;
        c.pszText = const_cast<wchar_t*>(t);
        return c;
    }

    static void insert_col(HWND list, int index, const wchar_t* text, int width) {
        LVCOLUMNW col = _col(text, width);
        ListView_InsertColumn(list, index, &col);
    }

    void load() {
        reset();
        try {
            std::string base = Workflow::construct_path_base(app_.state().input_string);
            auto resolved = Workflow::resolve_file(base);
            if (!resolved) {
                show_banner_error(
                    (L"ファイルが見つかりません: " + theme::to_wide(base) +
                     L".xlsx も .xls も存在しません").c_str());
                return;
            }
            app_.state().resolved_path = *resolved;

            auto records = Workflow::load_master(*resolved);
            if (records.empty()) {
                show_banner_error(L"マスターからレコードを 1 件も読み取れませんでした。"
                                  L"設定ファイルの開始行・列番号を確認してください。");
                return;
            }
            app_.state().records = std::move(records);
            ++app_.state().records_gen;
            populate_list();

            wchar_t buf[128];
            swprintf(buf, 128, L"  ✓ %zu 件のレコードを読み込みました",
                     app_.state().records.size());
            SetWindowTextW(summary_, buf);
            show_banner_ok(theme::to_wide(*resolved).c_str());
        }
        catch (const std::exception& e) {
            show_banner_error(theme::to_wide(e.what()).c_str());
        }
    }

    void populate_list() {
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
        }
    }

    void show_banner_ok(const wchar_t* path) {
        std::wstring msg = L"  ✓ 読み込み成功:  " + std::wstring(path);
        SetWindowTextW(banner_, msg.c_str());
        banner_kind_ = 0;
        banner_visible_ = true;
        ShowWindow(banner_, SW_SHOW);
        InvalidateRect(banner_, nullptr, TRUE);
        layout_dirty();
    }

    void show_banner_error(const wchar_t* msg) {
        std::wstring m = L"  ⚠  " + std::wstring(msg);
        SetWindowTextW(banner_, m.c_str());
        banner_kind_ = 1;
        banner_visible_ = true;
        ShowWindow(banner_, SW_SHOW);
        InvalidateRect(banner_, nullptr, TRUE);
        layout_dirty();
    }

    void layout_dirty() {
        RECT rc; GetClientRect(panel_, &rc);
        layout(rc.right, rc.bottom);
    }

public:
    HWND banner_handle()   const { return banner_; }
    int  banner_kind_state() const { return banner_kind_; }

private:
    HWND title_{}, path_lbl_{}, path_{}, load_btn_{}, banner_{}, list_{}, summary_{};
    int  banner_kind_{0};   // 0 = ok / 1 = error
    bool banner_visible_{false};
    uint32_t seen_input_gen_{0};
};

}  // namespace

std::unique_ptr<StagePanel> make_stage_master(MainWindow& app) {
    return std::make_unique<StageMaster>(app);
}

}  // namespace case1
#endif
