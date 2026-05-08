#ifdef _WIN32
#include "ui_stages.hpp"
#include "workflow.hpp"
#include "config.hpp"
#include <commctrl.h>
#include <commdlg.h>     // GetSaveFileNameW / OPENFILENAMEW
#include <shlobj.h>      // SHBrowseForFolderW / BROWSEINFOW
#include <objbase.h>     // CoTaskMemFree
#include <string>

namespace case1 {

namespace {

constexpr int ID_INPUT       = 200;
constexpr int ID_RADIO_PRN   = 210;
constexpr int ID_RADIO_INDIV = 211;
constexpr int ID_RADIO_MERGE = 212;
constexpr int ID_OUT_DIR     = 220;
constexpr int ID_BROWSE_DIR  = 221;
constexpr int ID_OUT_PDF     = 222;
constexpr int ID_BROWSE_PDF  = 223;

class StageInput : public ui::PanelBase {
public:
    StageInput(MainWindow& app) : PanelBase(app) {
        make_panel(L"Case1StageInput");

        // タイトル
        title_ = make_label(L"ステップ 1 — 可変文字列を入力", app_.fonts().title);

        // 説明
        desc_ = make_label(
            L"マスター Excel のパスを生成するための文字列を入力してください。\n"
            L"拡張子 (.xlsx / .xls) は自動判定します。",
            app_.fonts().body);

        // 入力欄
        input_label_ = make_label(theme::to_wide(config::INPUT_LABEL).c_str(),
                                  app_.fonts().heading);
        input_ = make_edit(0, ID_INPUT);
        SendMessageW(input_, EM_SETCUEBANNER, TRUE,
            (LPARAM)theme::to_wide(config::INPUT_HINT).c_str());

        // パスプレビュー
        preview_label_ = make_label(L"パスプレビュー", app_.fonts().heading);
        preview_       = make_label(L"", app_.fonts().mono);

        // 出力モード
        mode_label_ = make_label(L"出力モード", app_.fonts().heading);

        radio_prn_ = CreateWindowExW(0, L"BUTTON", L"プリンターに送信",
            WS_CHILD|WS_VISIBLE|WS_TABSTOP|BS_AUTORADIOBUTTON|WS_GROUP,
            0,0,200,32, panel_, (HMENU)(intptr_t)ID_RADIO_PRN,
            app_.hinst(), nullptr);
        SendMessageW(radio_prn_, WM_SETFONT, (WPARAM)app_.fonts().heading, TRUE);

        radio_indiv_ = CreateWindowExW(0, L"BUTTON", L"個別 PDF として出力",
            WS_CHILD|WS_VISIBLE|WS_TABSTOP|BS_AUTORADIOBUTTON,
            0,0,200,32, panel_, (HMENU)(intptr_t)ID_RADIO_INDIV,
            app_.hinst(), nullptr);
        SendMessageW(radio_indiv_, WM_SETFONT, (WPARAM)app_.fonts().heading, TRUE);

        radio_merge_ = CreateWindowExW(0, L"BUTTON", L"1つの PDF に結合して出力",
            WS_CHILD|WS_VISIBLE|WS_TABSTOP|BS_AUTORADIOBUTTON,
            0,0,200,32, panel_, (HMENU)(intptr_t)ID_RADIO_MERGE,
            app_.hinst(), nullptr);
        SendMessageW(radio_merge_, WM_SETFONT, (WPARAM)app_.fonts().heading, TRUE);

        // 出力先
        out_dir_label_ = make_label(L"出力フォルダ", app_.fonts().body);
        out_dir_ = make_edit();
        out_dir_browse_ = CreateWindowExW(0, L"BUTTON", L"参照…",
            WS_CHILD|WS_VISIBLE|WS_TABSTOP, 0,0,80,32,
            panel_, (HMENU)(intptr_t)ID_BROWSE_DIR, app_.hinst(), nullptr);
        SendMessageW(out_dir_browse_, WM_SETFONT, (WPARAM)app_.fonts().body, TRUE);

        out_pdf_label_ = make_label(L"出力 PDF パス", app_.fonts().body);
        out_pdf_ = make_edit();
        out_pdf_browse_ = CreateWindowExW(0, L"BUTTON", L"参照…",
            WS_CHILD|WS_VISIBLE|WS_TABSTOP, 0,0,80,32,
            panel_, (HMENU)(intptr_t)ID_BROWSE_PDF, app_.hinst(), nullptr);
        SendMessageW(out_pdf_browse_, WM_SETFONT, (WPARAM)app_.fonts().body, TRUE);

        // デフォルト
        SendMessageW(radio_indiv_, BM_SETCHECK, BST_CHECKED, 0);
        update_mode_visibility();
        update_preview();
    }

    ~StageInput() override = default;

    void layout(int w, int h) override {
        int pad = theme::scale(theme::Padding, app_.dpi());
        int gap = theme::scale(theme::Gap, app_.dpi());
        int x = pad, y = pad;
        int line = theme::scale(28, app_.dpi());
        int input_h = theme::scale(theme::InputHeight, app_.dpi());

        SetWindowPos(title_, nullptr, x, y, w-pad*2, theme::scale(36, app_.dpi()), SWP_NOZORDER);
        y += theme::scale(40, app_.dpi());
        SetWindowPos(desc_, nullptr, x, y, w-pad*2, theme::scale(48, app_.dpi()), SWP_NOZORDER);
        y += theme::scale(56, app_.dpi());

        SetWindowPos(input_label_, nullptr, x, y, w-pad*2, line, SWP_NOZORDER);
        y += line + gap/2;
        SetWindowPos(input_, nullptr, x, y, w-pad*2, input_h, SWP_NOZORDER);
        y += input_h + gap;

        SetWindowPos(preview_label_, nullptr, x, y, w-pad*2, line, SWP_NOZORDER);
        y += line;
        SetWindowPos(preview_, nullptr, x, y, w-pad*2, line, SWP_NOZORDER);
        y += line + gap;

        SetWindowPos(mode_label_, nullptr, x, y, w-pad*2, line, SWP_NOZORDER);
        y += line + gap/2;

        int rh = theme::scale(34, app_.dpi());
        SetWindowPos(radio_prn_,   nullptr, x, y, w-pad*2, rh, SWP_NOZORDER); y += rh;
        SetWindowPos(radio_indiv_, nullptr, x, y, w-pad*2, rh, SWP_NOZORDER); y += rh;
        SetWindowPos(radio_merge_, nullptr, x, y, w-pad*2, rh, SWP_NOZORDER); y += rh + gap;

        // 出力先
        int browse_w = theme::scale(100, app_.dpi());
        SetWindowPos(out_dir_label_, nullptr, x, y, w-pad*2, line, SWP_NOZORDER);
        SetWindowPos(out_pdf_label_, nullptr, x, y, w-pad*2, line, SWP_NOZORDER);
        y += line;
        int eh = theme::scale(36, app_.dpi());
        SetWindowPos(out_dir_, nullptr, x, y, w - pad*2 - browse_w - gap, eh, SWP_NOZORDER);
        SetWindowPos(out_dir_browse_, nullptr, x + (w - pad*2 - browse_w), y, browse_w, eh, SWP_NOZORDER);
        SetWindowPos(out_pdf_, nullptr, x, y, w - pad*2 - browse_w - gap, eh, SWP_NOZORDER);
        SetWindowPos(out_pdf_browse_, nullptr, x + (w - pad*2 - browse_w), y, browse_w, eh, SWP_NOZORDER);
    }

    bool can_advance() const override {
        if (current_input().empty()) return false;
        switch (current_mode()) {
            case OutputMode::PdfIndividual:
                return GetWindowTextLengthW(out_dir_) > 0;
            case OutputMode::PdfMerged:
                return GetWindowTextLengthW(out_pdf_) > 0;
            case OutputMode::Printer:
                return true;
        }
        return false;
    }

    bool commit() override {
        auto& s = app_.state();
        std::string in    = current_input();
        OutputMode  mode  = current_mode();
        std::string odir  = current_text(out_dir_);
        std::string opdf  = current_text(out_pdf_);
        // 内容が変化していれば世代カウンタを進めて後段の自動リセットを誘発
        bool changed = (in != s.input_string) || (mode != s.output_mode)
                    || (odir != s.output_dir) || (opdf != s.output_pdf_path);
        if (changed) ++s.input_gen;
        s.input_string    = std::move(in);
        s.output_mode     = mode;
        s.output_dir      = std::move(odir);
        s.output_pdf_path = std::move(opdf);
        return true;
    }

    void reset() override {
        SetWindowTextW(input_, L"");
        SetWindowTextW(out_dir_, L"");
        SetWindowTextW(out_pdf_, L"");
        SendMessageW(radio_indiv_, BM_SETCHECK, BST_CHECKED, 0);
        update_mode_visibility();
        update_preview();
    }

    void on_command(int id, int notif, HWND ctrl) override {
        switch (id) {
            case ID_RADIO_PRN:
            case ID_RADIO_INDIV:
            case ID_RADIO_MERGE:
                if (notif == BN_CLICKED) update_mode_visibility();
                return;
            case ID_BROWSE_DIR: if (notif == BN_CLICKED) browse_dir(); return;
            case ID_BROWSE_PDF: if (notif == BN_CLICKED) browse_pdf(); return;
            case ID_INPUT:
                if (notif == EN_CHANGE) update_preview();
                return;
        }
    }

private:
    std::string current_input() const {
        wchar_t buf[1024]{};
        GetWindowTextW(input_, buf, 1024);
        return theme::to_utf8(buf);
    }
    std::string current_text(HWND h) const {
        wchar_t buf[1024]{};
        GetWindowTextW(h, buf, 1024);
        return theme::to_utf8(buf);
    }
    OutputMode current_mode() const {
        if (SendMessageW(radio_prn_,   BM_GETCHECK, 0,0) == BST_CHECKED) return OutputMode::Printer;
        if (SendMessageW(radio_merge_, BM_GETCHECK, 0,0) == BST_CHECKED) return OutputMode::PdfMerged;
        return OutputMode::PdfIndividual;
    }
    void update_mode_visibility() {
        OutputMode m = current_mode();
        ShowWindow(out_dir_label_,   m == OutputMode::PdfIndividual ? SW_SHOW : SW_HIDE);
        ShowWindow(out_dir_,         m == OutputMode::PdfIndividual ? SW_SHOW : SW_HIDE);
        ShowWindow(out_dir_browse_,  m == OutputMode::PdfIndividual ? SW_SHOW : SW_HIDE);
        ShowWindow(out_pdf_label_,   m == OutputMode::PdfMerged ? SW_SHOW : SW_HIDE);
        ShowWindow(out_pdf_,         m == OutputMode::PdfMerged ? SW_SHOW : SW_HIDE);
        ShowWindow(out_pdf_browse_,  m == OutputMode::PdfMerged ? SW_SHOW : SW_HIDE);
    }
    void update_preview() {
        std::string in = current_input();
        std::string base = Workflow::construct_path_base(in);
        std::wstring text = L"  " + theme::to_wide(base) + L"  (.xlsx → .xls の順に試行)";
        SetWindowTextW(preview_, text.c_str());
    }

    void browse_dir() {
        wchar_t path[MAX_PATH]{};
        BROWSEINFOW bi{};
        bi.hwndOwner = app_.hwnd();
        bi.lpszTitle = L"出力フォルダを選んでください";
        bi.ulFlags   = BIF_NEWDIALOGSTYLE|BIF_RETURNONLYFSDIRS;
        if (LPITEMIDLIST pidl = SHBrowseForFolderW(&bi)) {
            SHGetPathFromIDListW(pidl, path);
            CoTaskMemFree(pidl);
            SetWindowTextW(out_dir_, path);
        }
    }
    void browse_pdf() {
        wchar_t path[MAX_PATH]{ L"output.pdf" };
        OPENFILENAMEW ofn{ sizeof(ofn) };
        ofn.hwndOwner = app_.hwnd();
        ofn.lpstrFilter = L"PDF\0*.pdf\0All\0*.*\0";
        ofn.lpstrFile = path; ofn.nMaxFile = MAX_PATH;
        ofn.lpstrDefExt = L"pdf";
        ofn.Flags = OFN_OVERWRITEPROMPT|OFN_PATHMUSTEXIST;
        if (GetSaveFileNameW(&ofn))
            SetWindowTextW(out_pdf_, path);
    }

    HWND title_{}, desc_{};
    HWND input_label_{}, input_{};
    HWND preview_label_{}, preview_{};
    HWND mode_label_{};
    HWND radio_prn_{}, radio_indiv_{}, radio_merge_{};
    HWND out_dir_label_{}, out_dir_{}, out_dir_browse_{};
    HWND out_pdf_label_{}, out_pdf_{}, out_pdf_browse_{};
};

}  // namespace

std::unique_ptr<StagePanel> make_stage_input(MainWindow& app) {
    return std::make_unique<StageInput>(app);
}

}  // namespace case1
#endif
