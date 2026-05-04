#pragma once
/**
 * case1 — ui_theme.hpp
 *
 * 配色・フォント・サイズの一元管理。
 * 「片手で操作」「マウス主導」を意識した大きめ寸法。
 */
#ifdef _WIN32
#define WIN32_LEAN_AND_MEAN
#include <windows.h>

namespace case1::theme {

// ============================================================
//  配色 (Bootstrap 風 / Material 風 寄せ)
// ============================================================
// 主役の青 (進む / 実行)
constexpr COLORREF Primary       = RGB(0x00, 0x66, 0xCC);
constexpr COLORREF PrimaryHover  = RGB(0x00, 0x52, 0xA3);
constexpr COLORREF PrimaryText   = RGB(0xFF, 0xFF, 0xFF);

// 成功・実行系の緑
constexpr COLORREF Success       = RGB(0x28, 0xA7, 0x45);
constexpr COLORREF SuccessHover  = RGB(0x1E, 0x7E, 0x34);

// 警告 (戻る / やり直す)
constexpr COLORREF Warning       = RGB(0xFD, 0x7E, 0x14);
constexpr COLORREF WarningHover  = RGB(0xE3, 0x6A, 0x09);

// 危険 (中止 / キャンセル)
constexpr COLORREF Danger        = RGB(0xDC, 0x35, 0x45);
constexpr COLORREF DangerHover   = RGB(0xC8, 0x2A, 0x37);

// 中立
constexpr COLORREF Neutral       = RGB(0x6C, 0x75, 0x7D);
constexpr COLORREF NeutralHover  = RGB(0x54, 0x5B, 0x62);

// 背景系
constexpr COLORREF BgWindow      = RGB(0xF8, 0xF9, 0xFA);  // 全体背景
constexpr COLORREF BgPanel       = RGB(0xFF, 0xFF, 0xFF);  // パネル背景
constexpr COLORREF BgError       = RGB(0xF8, 0xD7, 0xDA);  // エラーバナー
constexpr COLORREF BgInfo        = RGB(0xD1, 0xEC, 0xF1);  // 情報バナー

// テキスト
constexpr COLORREF TextNormal    = RGB(0x21, 0x25, 0x29);
constexpr COLORREF TextMuted     = RGB(0x6C, 0x75, 0x7D);
constexpr COLORREF TextOnError   = RGB(0x72, 0x1C, 0x24);

// 罫線
constexpr COLORREF Border        = RGB(0xDE, 0xE2, 0xE6);
constexpr COLORREF BorderActive  = Primary;

// ============================================================
//  寸法 (DPI 100% 基準 / DPI スケールは手動で適用)
// ============================================================
constexpr int WindowWidth   = 980;
constexpr int WindowHeight  = 720;

constexpr int Padding       = 16;
constexpr int PaddingLarge  = 24;
constexpr int Gap           = 12;

// 大きなボタン (片手操作前提で 48px 以上を確保)
constexpr int BtnHeightTall = 56;   // 主要アクション (実行・進む)
constexpr int BtnHeightMid  = 44;   // 二次アクション (戻る・やり直す)
constexpr int BtnHeightSmall = 32;  // 補助 (全選択・全解除など)

// 入力欄
constexpr int InputHeight   = 48;

// ステップインジケータ
constexpr int StepBarHeight = 96;
constexpr int StepCircleR   = 22;   // 円の半径

// ============================================================
//  フォント
// ============================================================
struct Fonts {
    HFONT body        {nullptr};   // Segoe UI 11pt
    HFONT bodyBold    {nullptr};
    HFONT heading     {nullptr};   // Segoe UI 14pt
    HFONT title       {nullptr};   // Segoe UI 18pt Bold
    HFONT large_input {nullptr};   // Segoe UI 16pt (入力欄用)
    HFONT mono        {nullptr};   // Consolas 10pt (パスプレビュー)

    void create(int dpi);
    void destroy();
};

// ============================================================
//  描画ヘルパー
// ============================================================
/// 角丸塗りつぶし
void fill_rounded(HDC dc, RECT r, COLORREF bg, int radius = 6);

/// テキスト中央描画
void draw_text_centered(HDC dc, const RECT& r, const wchar_t* text, COLORREF color, HFONT font);

/// バナー（背景色 + 左にアクセントカラー）
void draw_banner(HDC dc, RECT r, COLORREF bg, COLORREF accent, const wchar_t* text, HFONT font);

/// ステータス丸ドット
void draw_status_dot(HDC dc, int cx, int cy, int radius, COLORREF color);

/// DPI スケール (現在のモニタの DPI から論理 px → 物理 px へ)
int scale(int px, int dpi);

/// UTF-8 → UTF-16
std::wstring to_wide(std::string_view s);
std::string  to_utf8(const wchar_t* w);

// ============================================================
//  カスタム描画ボタン
// ============================================================
enum class BtnKind {
    Primary,    // 青 (進む / 実行)
    Success,    // 緑 (成功時の主アクション)
    Warning,    // 橙 (やり直す)
    Danger,     // 赤 (キャンセル / 中止)
    Neutral,    // グレー (戻る / 補助)
};

/// オーナー描画ボタンを WM_DRAWITEM 内で描画する
void draw_owner_button(LPDRAWITEMSTRUCT dis, BtnKind kind, const wchar_t* text, HFONT font);

}  // namespace case1::theme

#endif  // _WIN32
