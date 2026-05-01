#pragma once
/**
 * excellib/print_settings.hpp
 * Page layout and print configuration types.
 */

// 文字列型 (std::string) を使うため
#include <string>
// std::optional: 値があるかもしれない型
#include <optional>
// uint16_t など固定幅整数型
#include <cstdint>
// std::runtime_error など標準例外クラス
#include <stdexcept>

namespace excellib {

// PrintError: 印刷・PDF 出力に関するエラー
// using ... で基底クラスのコンストラクタを継承する
struct PrintError : std::runtime_error { using std::runtime_error::runtime_error; };

// ---- Paper size (matches OOXML / BIFF8 paperSize values) ----
// 用紙サイズを表す列挙型
// : uint16_t を付けることで基底型を uint16_t (2バイト符号なし整数) に指定
// 括弧内の数値は OOXML 仕様で定義された paperSize 属性の値
enum class PaperSize : uint16_t {
    Letter  = 1,    // レターサイズ (米国標準)
    A3      = 8,    // A3 (297×420mm)
    A4      = 9,    // A4 (210×297mm) ← 日本で最も一般的
    A5      = 11,   // A5 (148×210mm)
    B4      = 12,   // B4 (257×364mm)
    B5      = 13,   // B5 (182×257mm)
    A6      = 70,   // A6 (105×148mm) ハガキサイズ
    Custom  = 0,    // カスタムサイズ (custom_paper_width/height_mm で指定)
};

// Orientation: 印刷方向を表す列挙型
enum class Orientation {
    Portrait,   // 縦向き (縦長)
    Landscape   // 横向き (横長)
};

/// How to scale content to fit pages.
/// ページに収めるためのスケール方法を指定する列挙型
enum class FitTo {
    None,           ///< Use scale_percent instead. (スケールなし。scale_percent で倍率指定)
    Width,          ///< All columns fit on one page width; rows flow freely. (全列を1ページ幅に収める)
    Height,         ///< All rows fit on one page height; columns flow freely. (全行を1ページ高さに収める)
    WidthAndHeight, ///< Everything on one page (fit_to_pages_wide × fit_to_pages_tall). (全体を1ページに収める)
};

// ---- Print area (0-based, inclusive) ----
// 印刷範囲を表す構造体 (すべて 0-indexed・包括的)
struct PrintArea {
    // first_row: 印刷範囲の開始行 (0-indexed)
    uint32_t first_row{0};
    // first_col: 印刷範囲の開始列 (0-indexed)
    uint32_t first_col{0};
    // last_row: 印刷範囲の終了行 (0-indexed)
    uint32_t last_row{0};
    // last_col: 印刷範囲の終了列 (0-indexed)
    uint32_t last_col{0};

    /// Parse "A1:Z50" range string.  Normalises reversed coordinates.
    /// "A1:Z50" のような範囲文字列をパースして PrintArea を作成する
    /// 逆順の座標 (例: "Z50:A1") は自動的に正規化される
    static PrintArea from_range(const std::string& range);

    /// Returns "A1:Z50" style string.
    /// 印刷範囲を "A1:Z50" のような範囲文字列に変換して返す
    std::string to_range() const;

    // is_valid(): 有効な範囲かどうかを確認する (終了 >= 開始であることをチェック)
    bool is_valid() const { return last_row >= first_row && last_col >= first_col; }
};

// ---- Rows/columns to repeat on every page (0-based, inclusive) ----
// 各ページに繰り返し印刷するタイトル行・列を指定する構造体
struct RepeatTitles {
    // row_start, row_end: 繰り返す行の範囲 (0-indexed)
    // optional なので「繰り返し行なし」= std::nullopt を表現できる
    std::optional<uint32_t> row_start, row_end;   ///< Rows to repeat (e.g. header row)
    // col_start, col_end: 繰り返す列の範囲 (0-indexed)
    std::optional<uint32_t> col_start, col_end;   ///< Columns to repeat (e.g. label col)
};

// ---- Margins in inches ----
// ページ余白をインチ単位で指定する構造体
struct PageMargins {
    double left  {0.70};    // 左余白 (デフォルト: 0.70 インチ)
    double right {0.70};    // 右余白 (デフォルト: 0.70 インチ)
    double top   {0.75};    // 上余白 (デフォルト: 0.75 インチ)
    double bottom{0.75};    // 下余白 (デフォルト: 0.75 インチ)
    double header{0.30};    // ヘッダー余白 (デフォルト: 0.30 インチ)
    double footer{0.30};    // フッター余白 (デフォルト: 0.30 インチ)
};

/**
 * Header/footer strings use Excel codes:
 *   &L left  &C centre  &R right
 *   &P page number  &N total pages  &D date  &T time
 *   &F filename  &A sheet name
 *   &"FontName,Style"  &NN font size
 *
 * ヘッダー/フッター文字列の Excel コード説明:
 *   &L=左寄せ  &C=中央揃え  &R=右寄せ
 *   &P=ページ番号  &N=総ページ数  &D=日付  &T=時刻
 *   &F=ファイル名  &A=シート名
 *   &"フォント名,スタイル"  &NN=フォントサイズ(NN は数値)
 */
struct HeaderFooter {
    // odd_header: 奇数ページのヘッダー文字列
    std::string odd_header;
    // odd_footer: 奇数ページのフッター文字列
    std::string odd_footer;
    // even_header: 偶数ページのヘッダー (different_odd_even=true の時に有効)
    std::string even_header;
    // even_footer: 偶数ページのフッター (different_odd_even=true の時に有効)
    std::string even_footer;
    // different_odd_even: 奇数・偶数ページで異なるヘッダー/フッターを使うか
    bool different_odd_even{false};
    // different_first: 1ページ目に異なるヘッダー/フッターを使うか
    bool different_first   {false};
};

// ---- Full page setup (ページ設定の全オプションをまとめた構造体) ----
struct PageSetup {
    // 用紙サイズ (デフォルト: A4)
    PaperSize   paper_size {PaperSize::A4};
    // 印刷方向 (デフォルト: 縦向き)
    Orientation orientation{Orientation::Portrait};

    // Scaling — mutually exclusive with fit_to
    // スケーリング設定 (fit_to と scale_percent は排他的)
    // ページに収める方法 (デフォルト: None = scale_percent で倍率指定)
    FitTo    fit_to          {FitTo::None};
    // fit_to が Width / WidthAndHeight の時の横方向のページ数
    uint16_t fit_to_pages_wide{1};
    // fit_to が Height / WidthAndHeight の時の縦方向のページ数
    uint16_t fit_to_pages_tall{1};
    // fit_to == None の時に使う印刷倍率 (10〜400 %)
    uint16_t scale_percent   {100};   ///< Used when fit_to == FitTo::None (10–400)

    // 余白設定
    PageMargins  margins;
    // ヘッダー/フッター設定
    HeaderFooter header_footer;

    // 印刷範囲 (optional: 未設定なら全体を印刷)
    std::optional<PrintArea>    print_area;
    // タイトル行・列の繰り返し設定 (optional: 未設定なら繰り返しなし)
    std::optional<RepeatTitles> repeat_titles;

    // グリッド線を印刷するか (デフォルト: false)
    bool print_gridlines      {false};
    // 行・列のヘッダー (行番号・列番号) を印刷するか
    bool print_row_col_headings{false};
    // 白黒印刷にするか
    bool black_and_white      {false};
    // 下書き品質 (低解像度・高速) で印刷するか
    bool draft_quality        {false};
    // ページ順序: 上から下→横に進む (false) か 左から右→縦に進む (true) か
    bool page_order_over_then_down{false};

    // 開始ページ番号 (0 = 自動)
    uint16_t first_page_number{0};   ///< 0 = automatic
    // 印刷部数
    uint16_t copies           {1};

    // カスタム用紙サイズ (paper_size == PaperSize::Custom の時のみ有効)
    double custom_paper_width_mm {0};    // 幅 (mm)
    double custom_paper_height_mm{0};   // 高さ (mm)
};

// ---- Printer options (プリンター送信のオプション) ----
struct PrinterOptions {
    // プリンター名 (空文字列 = システムの既定プリンターを使用)
    std::string printer_name;   ///< Empty = use system default printer
    // 丁合い (複数部数を印刷するとき、1部ずつまとめて出力するか)
    bool collate{true};
    // 両面印刷にするか
    bool duplex {false};
};

// ---- PDF options (PDF 出力のオプション) ----
struct PdfOptions {
    // 出力先ファイルパス (必須)
    std::string output_path;         ///< Required: destination file path
    // PDF 生成後にデフォルトビューアで自動的に開くか
    bool        open_after{false};   ///< Open in default viewer when done
};

} // namespace excellib
