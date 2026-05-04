#ifdef _WIN32
#include "ui_app.hpp"
#include "ui_theme.hpp"
#include <commctrl.h>
#include <string>
#include <vector>

namespace case1 {

// ============================================================
//  コントロール ID
// ============================================================
enum : int {
    ID_NAV_BACK = 100,
    ID_NAV_REDO,
    ID_NAV_NEXT,
};

// ============================================================
//  ステップバーのウィンドウクラス
// ============================================================
constexpr wchar_t kStepBarClass[] = L"Case1StepBar";

struct StepBarData {
    MainWindow* app{};
};

static const wchar_t* stage_label(Stage s) {
    switch (s) {
        case Stage::Input:  return L"1. 入力";
        case Stage::Master: return L"2. マスター読込";
        case Stage::Jobs:   return L"3. ジョブ調整";
        case Stage::Run:    return L"4. 実行・結果";
    }
    return L"";
}

static COLORREF status_color(StageStatus st) {
    using namespace theme;
    switch (st) {
        case StageStatus::Pending: return Neutral;
        case StageStatus::Active:  return Primary;
        case StageStatus::Done:    return Success;
        case StageStatus::Error:   return Danger;
    }
    return Neutral;
}

static LRESULT CALLBACK stepbar_proc(HWND h, UINT m, WPARAM w, LPARAM l) {
    auto* d = reinterpret_cast<StepBarData*>(GetWindowLongPtrW(h, GWLP_USERDATA));
    switch (m) {
        case WM_NCCREATE: {
            auto cs = reinterpret_cast<CREATESTRUCTW*>(l);
            SetWindowLongPtrW(h, GWLP_USERDATA, (LONG_PTR)cs->lpCreateParams);
            return TRUE;
        }
        case WM_ERASEBKGND: return 1;
        case WM_LBUTTONDOWN: {
            if (!d || !d->app) break;
            RECT rc; GetClientRect(h, &rc);
            int x = LOWORD(l);
            int seg = (rc.right - rc.left) / STAGE_COUNT;
            int idx = x / seg;
            if (idx < 0 || idx >= STAGE_COUNT) return 0;

            // ジャンプ可否: 現在地は何もしない / 既に Done のステージは双方向 OK /
            // それ以外 (Pending) は「進む」を強制
            Stage target = (Stage)idx;
            Stage cur    = d->app->current_stage();
            if (target == cur) return 0;
            StageStatus st = d->app->state().status[idx];
            if ((int)target < (int)cur || st == StageStatus::Done || st == StageStatus::Error)
                d->app->show_stage(target);
            return 0;
        }
        case WM_PAINT: {
            PAINTSTRUCT ps; HDC dc = BeginPaint(h, &ps);
            RECT rc; GetClientRect(h, &rc);

            // 背景
            HBRUSH bg = CreateSolidBrush(theme::BgWindow);
            FillRect(dc, &rc, bg); DeleteObject(bg);

            if (!d || !d->app) { EndPaint(h, &ps); return 0; }

            int dpi = d->app->dpi();
            int seg_w = (rc.right - rc.left) / STAGE_COUNT;
            int cy = rc.bottom / 2;
            int radius = theme::scale(theme::StepCircleR, dpi);

            // 接続線
            HPEN line = CreatePen(PS_SOLID, theme::scale(3, dpi), theme::Border);
            HPEN op = (HPEN)SelectObject(dc, line);
            MoveToEx(dc, seg_w/2, cy, nullptr);
            LineTo(dc, seg_w*STAGE_COUNT - seg_w/2, cy);
            SelectObject(dc, op); DeleteObject(line);

            // 各ステップ
            for (int i = 0; i < STAGE_COUNT; ++i) {
                int cx = seg_w*i + seg_w/2;
                StageStatus st = d->app->state().status[i];
                COLORREF clr = status_color(st);

                // 円の縁
                int outer = radius + theme::scale(3, dpi);
                theme::draw_status_dot(dc, cx, cy, outer, theme::BgWindow);
                theme::draw_status_dot(dc, cx, cy, outer-1, clr);
                // 内側白抜き＋数字 or チェック
                theme::draw_status_dot(dc, cx, cy, radius - theme::scale(4, dpi),
                                       (st==StageStatus::Active||st==StageStatus::Done||st==StageStatus::Error)
                                            ? clr : theme::BgWindow);

                wchar_t buf[8];
                if (st == StageStatus::Done) lstrcpynW(buf, L"✓", 8);
                else                          swprintf(buf, 8, L"%d", i+1);

                RECT cr;
                cr.left = cx-radius; cr.right = cx+radius;
                cr.top  = cy-radius; cr.bottom= cy+radius;
                COLORREF fg = (st == StageStatus::Pending) ? theme::TextMuted : RGB(255,255,255);
                HFONT f = (st == StageStatus::Done) ? d->app->fonts().heading : d->app->fonts().bodyBold;
                theme::draw_text_centered(dc, cr, buf, fg, f);

                // ラベル
                RECT lr;
                lr.left = seg_w*i; lr.right = seg_w*(i+1);
                lr.top = cy + radius + theme::scale(8, dpi);
                lr.bottom = rc.bottom - theme::scale(4, dpi);
                COLORREF lbl_clr = (st == StageStatus::Pending) ? theme::TextMuted : theme::TextNormal;
                HFONT lbl_font = (st == StageStatus::Active) ? d->app->fonts().bodyBold : d->app->fonts().body;
                theme::draw_text_centered(dc, lr, stage_label((Stage)i), lbl_clr, lbl_font);
            }

            // 下端の薄いボーダー
            RECT bd = {0, rc.bottom-1, rc.right, rc.bottom};
            HBRUSH bdb = CreateSolidBrush(theme::Border);
            FillRect(dc, &bd, bdb); DeleteObject(bdb);

            EndPaint(h, &ps);
            return 0;
        }
        case WM_NCDESTROY:
            delete d;
            return 0;
    }
    return DefWindowProcW(h, m, w, l);
}

static void register_stepbar() {
    WNDCLASSW wc{};
    wc.lpfnWndProc = stepbar_proc;
    wc.hInstance = GetModuleHandleW(nullptr);
    wc.hCursor = LoadCursor(nullptr, IDC_HAND);
    wc.hbrBackground = nullptr;
    wc.lpszClassName = kStepBarClass;
    static bool done = false;
    if (!done) { RegisterClassW(&wc); done = true; }
}

// ============================================================
//  メインウィンドウ
// ============================================================
constexpr wchar_t kMainClass[] = L"Case1MainWindow";

MainWindow::MainWindow(HINSTANCE h, AppState& s)
    : hInst_(h), state_(s) {}

MainWindow::~MainWindow() { fonts_.destroy(); }

LRESULT CALLBACK MainWindow::wndproc_static(HWND h, UINT m, WPARAM w, LPARAM l) {
    MainWindow* self = nullptr;
    if (m == WM_NCCREATE) {
        auto cs = reinterpret_cast<CREATESTRUCTW*>(l);
        self = reinterpret_cast<MainWindow*>(cs->lpCreateParams);
        SetWindowLongPtrW(h, GWLP_USERDATA, (LONG_PTR)self);
        self->hwnd_ = h;
    } else {
        self = reinterpret_cast<MainWindow*>(GetWindowLongPtrW(h, GWLP_USERDATA));
    }
    if (!self) return DefWindowProcW(h, m, w, l);
    return self->wndproc(m, w, l);
}

LRESULT MainWindow::wndproc(UINT m, WPARAM w, LPARAM l) {
    switch (m) {
        case WM_CREATE:
            on_create();
            return 0;
        case WM_SIZE:
            on_size(LOWORD(l), HIWORD(l));
            return 0;
        case WM_ERASEBKGND: {
            HDC dc = (HDC)w;
            RECT rc; GetClientRect(hwnd_, &rc);
            HBRUSH br = CreateSolidBrush(theme::BgWindow);
            FillRect(dc, &rc, br); DeleteObject(br);
            return 1;
        }
        case WM_CTLCOLORSTATIC: {
            HDC dc = (HDC)w;
            SetBkMode(dc, TRANSPARENT);
            SetTextColor(dc, theme::TextNormal);
            return (LRESULT)GetStockObject(NULL_BRUSH);
        }
        case WM_DRAWITEM: {
            auto* dis = reinterpret_cast<LPDRAWITEMSTRUCT>(l);
            // 1) kind は window prop から (make_button が SetProp で格納)
            theme::BtnKind kind = theme::BtnKind::Primary;
            HANDLE p = GetPropW(dis->hwndItem, L"case1.btn.kind");
            if (p) kind = (theme::BtnKind)((intptr_t)p - 1);
            // 2) text は実際のボタンタイトルから
            wchar_t text[256] = L"";
            GetWindowTextW(dis->hwndItem, text, 256);
            theme::draw_owner_button(dis, kind, text, fonts_.heading);
            return TRUE;
        }
        case WM_COMMAND: {
            int id    = LOWORD(w);
            int notif = HIWORD(w);
            HWND ctrl = (HWND)l;
            on_command(id, ctrl);
            // 自分のナビ ID 以外は現在のステージへ
            if (id != ID_NAV_BACK && id != ID_NAV_REDO && id != ID_NAV_NEXT) {
                if (panels_[(int)current_])
                    panels_[(int)current_]->on_command(id, notif, ctrl);
            }
            // 入力変化に応じてナビゲーション再評価
            if (notif == EN_CHANGE || notif == BN_CLICKED)
                update_nav_buttons();
            return 0;
        }
        case WM_NOTIFY: {
            auto* hdr = (NMHDR*)l;
            if (panels_[(int)current_])
                return panels_[(int)current_]->on_notify(hdr);
            return 0;
        }
        case WM_APP_PROGRESS:
        case WM_APP_BATCH_DONE:
        case WM_APP_LOG:
            // 進捗・結果メッセージは常に Run ステージへ。
            // 現在ステージにルートすると、ユーザーがステップバーで別ステージへ
            // 移動した際に heap allocate したペイロードがリークする。
            if (panels_[(int)Stage::Run])
                panels_[(int)Stage::Run]->on_app_message(m, w, l);
            return 0;
        case WM_DPICHANGED: {
            dpi_ = HIWORD(w);
            fonts_.destroy(); fonts_.create(dpi_);
            RECT* nr = reinterpret_cast<RECT*>(l);
            SetWindowPos(hwnd_, nullptr, nr->left, nr->top,
                         nr->right-nr->left, nr->bottom-nr->top,
                         SWP_NOZORDER|SWP_NOACTIVATE);
            return 0;
        }
        case WM_DESTROY:
            PostQuitMessage(0);
            return 0;
    }
    return DefWindowProcW(hwnd_, m, w, l);
}

HWND MainWindow::create() {
    INITCOMMONCONTROLSEX icc{ sizeof(icc), ICC_LISTVIEW_CLASSES|ICC_PROGRESS_CLASS|ICC_TAB_CLASSES };
    InitCommonControlsEx(&icc);

    register_stepbar();

    WNDCLASSW wc{};
    wc.lpfnWndProc   = wndproc_static;
    wc.hInstance     = hInst_;
    wc.hCursor       = LoadCursor(nullptr, IDC_ARROW);
    wc.hbrBackground = nullptr;
    wc.lpszClassName = kMainClass;
    RegisterClassW(&wc);

    // GetDpiForSystem は Win10 1607+。古い環境向けにフォールバック
    typedef UINT (WINAPI* PFN_GDFS)();
    if (HMODULE u32 = GetModuleHandleW(L"user32.dll")) {
        if (auto fn = (PFN_GDFS)GetProcAddress(u32, "GetDpiForSystem")) {
            dpi_ = (int)fn();
        }
    }
    if (dpi_ <= 0) {
        HDC sdc = GetDC(nullptr);
        dpi_ = GetDeviceCaps(sdc, LOGPIXELSX);
        ReleaseDC(nullptr, sdc);
    }
    fonts_.create(dpi_);

    auto title = theme::to_wide(config::APP_TITLE);

    int W = theme::scale(theme::WindowWidth, dpi_);
    int H = theme::scale(theme::WindowHeight, dpi_);

    HWND h = CreateWindowExW(
        0, kMainClass, title.c_str(),
        WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, CW_USEDEFAULT, W, H,
        nullptr, nullptr, hInst_, this);

    return h;
}

void MainWindow::on_create() {
    create_step_bar();
    create_stage_panels();
    create_nav_buttons();
    show_stage(Stage::Input);
}

void MainWindow::create_step_bar() {
    auto* sd = new StepBarData{ this };
    step_bar_ = CreateWindowExW(
        0, kStepBarClass, L"",
        WS_CHILD|WS_VISIBLE,
        0, 0, 100, 100,
        hwnd_, nullptr, hInst_, sd);
}

void MainWindow::create_stage_panels() {
    panels_[(int)Stage::Input ] = make_stage_input (*this);
    panels_[(int)Stage::Master] = make_stage_master(*this);
    panels_[(int)Stage::Jobs  ] = make_stage_jobs  (*this);
    panels_[(int)Stage::Run   ] = make_stage_run   (*this);
}

static void make_nav_btn(HWND parent, HWND& out, int id, theme::BtnKind kind,
                          const wchar_t* text, HINSTANCE inst) {
    out = CreateWindowExW(0, L"BUTTON", text,
        WS_CHILD|WS_VISIBLE|BS_OWNERDRAW,
        0,0,100,40, parent, (HMENU)(intptr_t)id, inst, nullptr);
    SetPropW(out, L"case1.btn.kind", (HANDLE)(intptr_t)((int)kind + 1));
}

void MainWindow::create_nav_buttons() {
    make_nav_btn(hwnd_, nav_back_, ID_NAV_BACK, theme::BtnKind::Neutral,
                 L"◀ 戻る", hInst_);
    make_nav_btn(hwnd_, nav_redo_, ID_NAV_REDO, theme::BtnKind::Warning,
                 L"↻ このステージをやり直す", hInst_);
    make_nav_btn(hwnd_, nav_next_, ID_NAV_NEXT, theme::BtnKind::Primary,
                 L"進む ▶", hInst_);
    update_nav_buttons();
}

void MainWindow::update_nav_buttons() {
    EnableWindow(nav_back_, current_ != Stage::Input);
    EnableWindow(nav_redo_, current_ != Stage::Input);
    bool can_advance = panels_[(int)current_] && panels_[(int)current_]->can_advance();
    EnableWindow(nav_next_, can_advance);

    // ステージ4 では NEXT を「閉じる」(Success) に切り替える
    if (current_ == Stage::Run) {
        SetWindowTextW(nav_next_, L"閉じる");
        SetPropW(nav_next_, L"case1.btn.kind",
                 (HANDLE)(intptr_t)((int)theme::BtnKind::Success + 1));
    } else {
        SetWindowTextW(nav_next_, L"進む ▶");
        SetPropW(nav_next_, L"case1.btn.kind",
                 (HANDLE)(intptr_t)((int)theme::BtnKind::Primary + 1));
    }
    InvalidateRect(nav_back_, nullptr, TRUE);
    InvalidateRect(nav_redo_, nullptr, TRUE);
    InvalidateRect(nav_next_, nullptr, TRUE);
}

void MainWindow::on_size(int w, int h) {
    int pad = theme::scale(theme::Padding, dpi_);
    int step_h = theme::scale(theme::StepBarHeight, dpi_);
    int btn_h = theme::scale(theme::BtnHeightTall, dpi_);
    int btn_h2= theme::scale(theme::BtnHeightMid, dpi_);

    SetWindowPos(step_bar_, nullptr, 0, 0, w, step_h, SWP_NOZORDER);

    int content_top = step_h;
    int nav_top = h - pad - btn_h;
    int content_h = nav_top - content_top - pad;

    for (int i = 0; i < STAGE_COUNT; ++i) {
        if (panels_[i]) {
            HWND ph = panels_[i]->hwnd();
            SetWindowPos(ph, nullptr, pad, content_top + pad/2,
                         w - pad*2, content_h, SWP_NOZORDER);
            panels_[i]->layout(w - pad*2, content_h);
        }
    }

    // ナビゲーションボタン
    int back_w = theme::scale(140, dpi_);
    int redo_w = theme::scale(280, dpi_);
    int next_w = theme::scale(220, dpi_);

    SetWindowPos(nav_back_, nullptr,
                 pad, nav_top, back_w, btn_h2, SWP_NOZORDER);
    SetWindowPos(nav_redo_, nullptr,
                 pad + back_w + theme::scale(12, dpi_), nav_top,
                 redo_w, btn_h2, SWP_NOZORDER);
    SetWindowPos(nav_next_, nullptr,
                 w - pad - next_w, nav_top, next_w, btn_h, SWP_NOZORDER);
}

void MainWindow::on_command(int id, HWND) {
    switch (id) {
        case ID_NAV_BACK:    on_back();    break;
        case ID_NAV_REDO:    on_redo();    break;
        case ID_NAV_NEXT:    on_advance(); break;
    }
}

void MainWindow::show_stage(Stage s) {
    if (panels_[(int)current_]) {
        panels_[(int)current_]->on_hide();
        ShowWindow(panels_[(int)current_]->hwnd(), SW_HIDE);
    }
    current_ = s;
    // 「現在地」は常に Active。それより前の Active は Done として残す
    // (commit 経由で Done にしたものはそのまま、未確定で前進した場合の表示安定化)。
    for (int i = 0; i < STAGE_COUNT; ++i) {
        if (i == (int)s) {
            state_.status[i] = StageStatus::Active;
        } else if (state_.status[i] == StageStatus::Active) {
            // 別の Active が残っていたら、現在地より前なら Done、後なら Pending
            state_.status[i] = (i < (int)s) ? StageStatus::Done : StageStatus::Pending;
        }
    }
    if (panels_[(int)s]) {
        ShowWindow(panels_[(int)s]->hwnd(), SW_SHOW);
        panels_[(int)s]->on_show();
    }
    invalidate_step_bar();
    update_nav_buttons();
}

void MainWindow::set_status(Stage s, StageStatus st) {
    state_.status[(int)s] = st;
    invalidate_step_bar();
    update_nav_buttons();
}

void MainWindow::invalidate_step_bar() {
    InvalidateRect(step_bar_, nullptr, TRUE);
}

void MainWindow::refresh_nav() { update_nav_buttons(); }

void MainWindow::on_back() {
    if (current_ == Stage::Input) return;
    show_stage((Stage)((int)current_ - 1));
}

void MainWindow::on_redo() {
    panels_[(int)current_]->reset();
    set_status(current_, StageStatus::Active);
    // 後段はすべて pending に戻す
    for (int i = (int)current_ + 1; i < STAGE_COUNT; ++i)
        set_status((Stage)i, StageStatus::Pending);
}

void MainWindow::on_advance() {
    if (current_ == Stage::Run) {
        DestroyWindow(hwnd_);
        return;
    }
    if (!panels_[(int)current_]->commit()) return;
    set_status(current_, StageStatus::Done);
    show_stage((Stage)((int)current_ + 1));
}

int MainWindow::run_message_loop() {
    ShowWindow(hwnd_, SW_SHOW);
    UpdateWindow(hwnd_);

    MSG msg;
    while (GetMessageW(&msg, nullptr, 0, 0) > 0) {
        if (!IsDialogMessageW(hwnd_, &msg)) {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }
    return (int)msg.wParam;
}

}  // namespace case1
#endif
