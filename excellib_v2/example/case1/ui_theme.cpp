#ifdef _WIN32
#include "ui_theme.hpp"
#include <string>
#include <vector>

namespace case1::theme {

// ============================================================
//  Fonts
// ============================================================
static HFONT make_font(int height_pt, int dpi, int weight = FW_NORMAL,
                       const wchar_t* face = L"Segoe UI") {
    LOGFONTW lf{};
    lf.lfHeight = -MulDiv(height_pt, dpi, 72);
    lf.lfWeight = weight;
    lf.lfCharSet = DEFAULT_CHARSET;
    lf.lfQuality = CLEARTYPE_QUALITY;
    lstrcpynW(lf.lfFaceName, face, LF_FACESIZE);
    return CreateFontIndirectW(&lf);
}

void Fonts::create(int dpi) {
    body         = make_font(11, dpi, FW_NORMAL);
    bodyBold     = make_font(11, dpi, FW_SEMIBOLD);
    heading      = make_font(14, dpi, FW_SEMIBOLD);
    title        = make_font(18, dpi, FW_BOLD);
    large_input  = make_font(16, dpi, FW_NORMAL);
    mono         = make_font(10, dpi, FW_NORMAL, L"Consolas");
}

void Fonts::destroy() {
    for (HFONT* h : {&body,&bodyBold,&heading,&title,&large_input,&mono})
        if (*h) { DeleteObject(*h); *h = nullptr; }
}

// ============================================================
//  Drawing helpers
// ============================================================
int scale(int px, int dpi) { return MulDiv(px, dpi, 96); }

void fill_rounded(HDC dc, RECT r, COLORREF bg, int radius) {
    HBRUSH br = CreateSolidBrush(bg);
    HPEN   pn = CreatePen(PS_NULL, 0, 0);
    HBRUSH ob = (HBRUSH)SelectObject(dc, br);
    HPEN   op = (HPEN)  SelectObject(dc, pn);
    RoundRect(dc, r.left, r.top, r.right, r.bottom, radius*2, radius*2);
    SelectObject(dc, ob); SelectObject(dc, op);
    DeleteObject(br); DeleteObject(pn);
}

void draw_text_centered(HDC dc, const RECT& r, const wchar_t* text,
                         COLORREF color, HFONT font) {
    int oldMode = SetBkMode(dc, TRANSPARENT);
    COLORREF oldClr = SetTextColor(dc, color);
    HFONT oldFont = (HFONT)SelectObject(dc, font);
    DrawTextW(dc, text, -1, (RECT*)&r,
              DT_CENTER|DT_VCENTER|DT_SINGLELINE|DT_NOPREFIX);
    SelectObject(dc, oldFont);
    SetTextColor(dc, oldClr);
    SetBkMode(dc, oldMode);
}

void draw_banner(HDC dc, RECT r, COLORREF bg, COLORREF accent,
                  const wchar_t* text, HFONT font) {
    fill_rounded(dc, r, bg, 6);

    // 左側 6px のアクセントバー
    RECT bar = r; bar.right = r.left + 6;
    HBRUSH br = CreateSolidBrush(accent);
    FillRect(dc, &bar, br);
    DeleteObject(br);

    RECT tr = r; tr.left += 16; tr.right -= 12;
    int oldMode = SetBkMode(dc, TRANSPARENT);
    COLORREF oldClr = SetTextColor(dc, RGB(0x21,0x25,0x29));
    HFONT oldFont = (HFONT)SelectObject(dc, font);
    DrawTextW(dc, text, -1, &tr, DT_LEFT|DT_VCENTER|DT_WORDBREAK);
    SelectObject(dc, oldFont);
    SetTextColor(dc, oldClr);
    SetBkMode(dc, oldMode);
}

void draw_status_dot(HDC dc, int cx, int cy, int radius, COLORREF color) {
    HBRUSH br = CreateSolidBrush(color);
    HPEN   pn = CreatePen(PS_NULL, 0, 0);
    HBRUSH ob = (HBRUSH)SelectObject(dc, br);
    HPEN   op = (HPEN)  SelectObject(dc, pn);
    Ellipse(dc, cx-radius, cy-radius, cx+radius, cy+radius);
    SelectObject(dc, ob); SelectObject(dc, op);
    DeleteObject(br); DeleteObject(pn);
}

// ============================================================
//  Custom-drawn buttons
// ============================================================
static void kind_colors(BtnKind k, bool hover, COLORREF& bg, COLORREF& fg) {
    fg = RGB(0xFF,0xFF,0xFF);
    switch (k) {
        case BtnKind::Primary: bg = hover ? PrimaryHover : Primary; break;
        case BtnKind::Success: bg = hover ? SuccessHover : Success; break;
        case BtnKind::Warning: bg = hover ? WarningHover : Warning; break;
        case BtnKind::Danger:  bg = hover ? DangerHover  : Danger;  break;
        case BtnKind::Neutral: bg = hover ? NeutralHover : Neutral; break;
    }
}

void draw_owner_button(LPDRAWITEMSTRUCT dis, BtnKind kind,
                        const wchar_t* text, HFONT font) {
    bool pressed = (dis->itemState & ODS_SELECTED) != 0;
    bool focused = (dis->itemState & ODS_FOCUS)    != 0;
    bool disabled= (dis->itemState & ODS_DISABLED) != 0;

    COLORREF bg, fg;
    kind_colors(kind, pressed, bg, fg);
    if (disabled) {
        bg = RGB(0xCE,0xD4,0xDA);
        fg = RGB(0x6C,0x75,0x7D);
    }

    RECT r = dis->rcItem;
    fill_rounded(dis->hDC, r, bg, 8);

    // Focus 時の縁取り
    if (focused && !disabled) {
        HPEN pn = CreatePen(PS_SOLID, 2, RGB(0xFF,0xFF,0xFF));
        HPEN op = (HPEN)SelectObject(dis->hDC, pn);
        HBRUSH ob = (HBRUSH)SelectObject(dis->hDC, GetStockObject(NULL_BRUSH));
        RoundRect(dis->hDC, r.left+3, r.top+3, r.right-3, r.bottom-3, 12, 12);
        SelectObject(dis->hDC, op); SelectObject(dis->hDC, ob);
        DeleteObject(pn);
    }

    draw_text_centered(dis->hDC, r, text, fg, font);
}

// ============================================================
//  Encoding
// ============================================================
std::wstring to_wide(std::string_view s) {
    if (s.empty()) return {};
    int n = MultiByteToWideChar(CP_UTF8, 0, s.data(), (int)s.size(), nullptr, 0);
    std::wstring w(n, 0);
    MultiByteToWideChar(CP_UTF8, 0, s.data(), (int)s.size(), w.data(), n);
    return w;
}

std::string to_utf8(const wchar_t* w) {
    if (!w || !*w) return {};
    int n = WideCharToMultiByte(CP_UTF8, 0, w, -1, nullptr, 0, nullptr, nullptr);
    std::string s(n>0 ? n-1 : 0, 0);
    WideCharToMultiByte(CP_UTF8, 0, w, -1, s.data(), n, nullptr, nullptr);
    return s;
}

}  // namespace case1::theme
#endif
