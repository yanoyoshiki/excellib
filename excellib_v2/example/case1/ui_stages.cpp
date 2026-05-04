#ifdef _WIN32
#include "ui_stages.hpp"

namespace case1::ui {

LRESULT CALLBACK PanelBase::panel_forward_proc(HWND h, UINT m, WPARAM w, LPARAM l) {
    // 子コントロールからの通知を MainWindow (=GetParent) に転送
    switch (m) {
        case WM_COMMAND:
        case WM_NOTIFY:
        case WM_DRAWITEM:
        case WM_CTLCOLORSTATIC:
        case WM_CTLCOLORBTN:
        case WM_CTLCOLOREDIT:
            return SendMessageW(GetParent(h), m, w, l);
    }
    return DefWindowProcW(h, m, w, l);
}

HWND PanelBase::make_panel(const wchar_t* class_name) {
    WNDCLASSW existing{};
    if (!GetClassInfoW(app_.hinst(), class_name, &existing)) {
        WNDCLASSW wc{};
        wc.lpfnWndProc   = panel_forward_proc;
        wc.hInstance     = app_.hinst();
        wc.hCursor       = LoadCursor(nullptr, IDC_ARROW);
        wc.hbrBackground = nullptr;
        wc.lpszClassName = class_name;
        RegisterClassW(&wc);
    }
    panel_ = CreateWindowExW(
        0, class_name, L"",
        WS_CHILD|WS_CLIPCHILDREN,
        0,0,100,100, app_.hwnd(), nullptr, app_.hinst(), nullptr);
    return panel_;
}

HWND PanelBase::make_label(const wchar_t* text, HFONT font) {
    HWND h = CreateWindowExW(0, L"STATIC", text,
        WS_CHILD|WS_VISIBLE|SS_LEFT,
        0,0,100,20, panel_, nullptr, app_.hinst(), nullptr);
    SendMessageW(h, WM_SETFONT, (WPARAM)font, TRUE);
    return h;
}

HWND PanelBase::make_banner() {
    HWND h = CreateWindowExW(0, L"STATIC", L"",
        WS_CHILD|SS_LEFT,
        0,0,100,40, panel_, nullptr, app_.hinst(), nullptr);
    SendMessageW(h, WM_SETFONT, (WPARAM)app_.fonts().bodyBold, TRUE);
    return h;
}

HWND PanelBase::make_group(const wchar_t* title) {
    HWND h = CreateWindowExW(0, L"BUTTON", title,
        WS_CHILD|WS_VISIBLE|BS_GROUPBOX,
        0,0,100,100, panel_, nullptr, app_.hinst(), nullptr);
    SendMessageW(h, WM_SETFONT, (WPARAM)app_.fonts().bodyBold, TRUE);
    return h;
}

HWND PanelBase::make_edit(int extra_styles, int id) {
    HWND h = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
        WS_CHILD|WS_VISIBLE|WS_TABSTOP|ES_AUTOHSCROLL|extra_styles,
        0,0,200,32, panel_, (HMENU)(intptr_t)id, app_.hinst(), nullptr);
    SendMessageW(h, WM_SETFONT, (WPARAM)app_.fonts().large_input, TRUE);
    return h;
}

HWND PanelBase::make_button(int id, theme::BtnKind kind, const wchar_t* text) {
    HWND h = CreateWindowExW(0, L"BUTTON", text,
        WS_CHILD|WS_VISIBLE|WS_TABSTOP|BS_OWNERDRAW,
        0,0,160,40, panel_, (HMENU)(intptr_t)id, app_.hinst(), nullptr);
    // BtnKind を window prop に格納 (WM_DRAWITEM 時に取り出す)
    SetPropW(h, L"case1.btn.kind", (HANDLE)(intptr_t)((int)kind + 1));
    return h;
}

HWND PanelBase::make_listview(int extra_styles) {
    HWND h = CreateWindowExW(WS_EX_CLIENTEDGE, WC_LISTVIEWW, L"",
        WS_CHILD|WS_VISIBLE|WS_TABSTOP|LVS_REPORT|LVS_SHOWSELALWAYS|extra_styles,
        0,0,200,200, panel_, nullptr, app_.hinst(), nullptr);
    ListView_SetExtendedListViewStyle(h,
        LVS_EX_FULLROWSELECT|LVS_EX_GRIDLINES|LVS_EX_DOUBLEBUFFER);
    SendMessageW(h, WM_SETFONT, (WPARAM)app_.fonts().body, TRUE);
    return h;
}

HWND PanelBase::make_progress() {
    HWND h = CreateWindowExW(0, PROGRESS_CLASSW, L"",
        WS_CHILD|WS_VISIBLE|PBS_SMOOTH,
        0,0,200,20, panel_, nullptr, app_.hinst(), nullptr);
    return h;
}

}  // namespace case1::ui
#endif
