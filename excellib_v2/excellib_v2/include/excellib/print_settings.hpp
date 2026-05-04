#pragma once
/**
 * excellib/print_settings.hpp
 * Page layout and print configuration types.
 */
#include <string>
#include <optional>
#include <cstdint>
#include <stdexcept>

namespace excellib {

struct PrintError : std::runtime_error { using std::runtime_error::runtime_error; };

// ---- Paper size (matches OOXML / BIFF8 paperSize values) ----
enum class PaperSize : uint16_t {
    Letter  = 1,
    A3      = 8,
    A4      = 9,
    A5      = 11,
    B4      = 12,
    B5      = 13,
    A6      = 70,
    Custom  = 0,
};

enum class Orientation { Portrait, Landscape };

/// How to scale content to fit pages.
enum class FitTo {
    None,           ///< Use scale_percent instead.
    Width,          ///< All columns fit on one page width; rows flow freely.
    Height,         ///< All rows fit on one page height; columns flow freely.
    WidthAndHeight, ///< Everything on one page (fit_to_pages_wide × fit_to_pages_tall).
};

// ---- Print area (0-based, inclusive) ----
struct PrintArea {
    uint32_t first_row{0}, first_col{0};
    uint32_t last_row{0},  last_col{0};
    bool col_range_only{false};   ///< true = entire columns (no row bounds), e.g. A:D

    /// Parse "A1:Z50" range string.  Normalises reversed coordinates.
    static PrintArea from_range(const std::string& range);

    /// Columns-only print area: "A:D" or "A" selects entire columns.
    static PrintArea columns_only(std::string_view col_range);
    static PrintArea columns_only(uint32_t col_start, uint32_t col_end);

    /// "A1:Z50" or "$A:$D" depending on col_range_only.
    std::string to_range() const;

    bool is_columns_only() const noexcept { return col_range_only; }
    bool is_valid() const {
        return last_col >= first_col && (col_range_only || last_row >= first_row);
    }
};

// ---- Rows/columns to repeat on every page (0-based, inclusive) ----
struct RepeatTitles {
    std::optional<uint32_t> row_start, row_end;   ///< Rows to repeat (e.g. header row)
    std::optional<uint32_t> col_start, col_end;   ///< Columns to repeat (e.g. label col)
};

// ---- Margins in inches ----
struct PageMargins {
    double left  {0.70};
    double right {0.70};
    double top   {0.75};
    double bottom{0.75};
    double header{0.30};
    double footer{0.30};
};

/**
 * Header/footer strings use Excel codes:
 *   &L left  &C centre  &R right
 *   &P page number  &N total pages  &D date  &T time
 *   &F filename  &A sheet name
 *   &"FontName,Style"  &NN font size
 */
struct HeaderFooter {
    std::string odd_header;
    std::string odd_footer;
    std::string even_header;
    std::string even_footer;
    bool different_odd_even{false};
    bool different_first   {false};
};

// ---- Full page setup ----
struct PageSetup {
    PaperSize   paper_size {PaperSize::A4};
    Orientation orientation{Orientation::Portrait};

    // Scaling — mutually exclusive with fit_to
    FitTo    fit_to          {FitTo::None};
    uint16_t fit_to_pages_wide{1};
    uint16_t fit_to_pages_tall{1};
    uint16_t scale_percent   {100};   ///< Used when fit_to == FitTo::None (10–400)

    PageMargins  margins;
    HeaderFooter header_footer;

    std::optional<PrintArea>    print_area;
    std::optional<RepeatTitles> repeat_titles;

    bool print_gridlines      {false};
    bool print_row_col_headings{false};
    bool black_and_white      {false};
    bool draft_quality        {false};
    bool page_order_over_then_down{false};

    uint16_t first_page_number{0};   ///< 0 = automatic
    uint16_t copies           {1};

    // Used only when paper_size == PaperSize::Custom
    double custom_paper_width_mm {0};
    double custom_paper_height_mm{0};
};

// ---- Printer options ----
struct PrinterOptions {
    std::string printer_name;   ///< Empty = use system default printer
    bool collate{true};
    bool duplex {false};
};

// ---- PDF options ----
struct PdfOptions {
    std::string output_path;         ///< Required: destination file path
    bool        open_after{false};   ///< Open in default viewer when done
};

} // namespace excellib
