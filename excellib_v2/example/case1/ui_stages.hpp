#pragma once
/**
 * case1 — ui_stages.hpp
 * 各ステージで共通利用する小道具
 */
#ifdef _WIN32
#include "ui_app.hpp"
#include "ui_theme.hpp"
#include <commctrl.h>
#include <string>

namespace case1::ui {

// ============================================================
//  パネル基本クラス（子コントロールを乗せる空のチャイルドウィンドウ）
// ============================================================
class PanelBase : public StagePanel {
public:
    PanelBase(MainWindow& app) : app_(app) {}
    HWND hwnd() const override { return panel_; }

protected:
    /// 子コントロールからの WM_COMMAND 等を MainWindow に転送する WNDPROC。
    /// パネルクラス登録時にこれを使う。
    static LRESULT CALLBACK panel_forward_proc(HWND, UINT, WPARAM, LPARAM);

    /// チャイルドウィンドウクラスを (重複なく) 登録してパネルを生成
    HWND make_panel(const wchar_t* class_name);

    /// 静的テキスト
    HWND make_label(const wchar_t* text, HFONT font);

    /// バナー (背景色付きの STATIC) — テキストは WM_SETTEXT で更新
    HWND make_banner();

    /// グループボックス (装飾)
    HWND make_group(const wchar_t* title);

    /// 単行入力欄
    HWND make_edit(int extra_styles = 0, int id = 0);

    /// オーナー描画ボタン (kind を window prop に格納)
    HWND make_button(int id, theme::BtnKind kind, const wchar_t* text);

    /// ListView (報告系)
    HWND make_listview(int extra_styles = 0);

    /// 進捗バー
    HWND make_progress();

    MainWindow& app_;
    HWND        panel_{};
};

}  // namespace case1::ui
#endif
