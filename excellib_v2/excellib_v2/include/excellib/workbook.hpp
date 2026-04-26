#pragma once
/**
 * excellib/workbook.hpp
 * Core types: CellValue, Cell, Sheet, Workbook, WorkbookFactory
 */
#include <string>
#include <vector>
#include <memory>
#include <variant>
#include <optional>
#include <stdexcept>
#include <cstdint>
#include <functional>

namespace excellib {

// ============================================================
//  Error hierarchy
// ============================================================
struct ExcelError  : std::runtime_error { using std::runtime_error::runtime_error; };
struct ParseError  : ExcelError         { using ExcelError::ExcelError; };
struct FormatError : ExcelError         { using ExcelError::ExcelError; };
struct IOError     : ExcelError         { using ExcelError::ExcelError; };
struct RangeError  : ExcelError         { using ExcelError::ExcelError; };
struct WriteError  : ExcelError         { using ExcelError::ExcelError; };

// ============================================================
//  CellValue — strict typed variant, no silent coercion
// ============================================================
struct BlankValue {};
struct ErrorValue { std::string code; };   // "#REF!", "#DIV/0!", etc.

using CellValue = std::variant<
    BlankValue,
    bool,
    int64_t,
    double,
    std::string,
    ErrorValue
>;

// ---- Type predicates ----
inline bool is_blank  (const CellValue& v) { return std::holds_alternative<BlankValue>(v); }
inline bool is_bool   (const CellValue& v) { return std::holds_alternative<bool>(v); }
inline bool is_int    (const CellValue& v) { return std::holds_alternative<int64_t>(v); }
inline bool is_double (const CellValue& v) { return std::holds_alternative<double>(v); }
inline bool is_string (const CellValue& v) { return std::holds_alternative<std::string>(v); }
inline bool is_error  (const CellValue& v) { return std::holds_alternative<ErrorValue>(v); }
inline bool is_numeric(const CellValue& v) { return is_int(v) || is_double(v); }

// ---- Strict getters (throw std::bad_variant_access on type mismatch) ----
inline bool        get_bool  (const CellValue& v) { return std::get<bool>(v); }
inline int64_t     get_int   (const CellValue& v) { return std::get<int64_t>(v); }
inline double      get_double(const CellValue& v) { return std::get<double>(v); }
inline std::string get_string(const CellValue& v) { return std::get<std::string>(v); }
inline ErrorValue  get_error (const CellValue& v) { return std::get<ErrorValue>(v); }

// ---- Numeric coercion (explicit) ----
inline double as_double(const CellValue& v) {
    if (is_int(v))    return static_cast<double>(get_int(v));
    if (is_double(v)) return get_double(v);
    throw FormatError("Cell value is not numeric");
}

// ---- Convert any non-error value to string ----
std::string to_string(const CellValue& v);

// ============================================================
//  CellAddress (0-based, row & col)
// ============================================================
struct CellAddress {
    uint32_t row{0};
    uint32_t col{0};

    bool operator==(const CellAddress& o) const { return row==o.row && col==o.col; }

    /// Parse "B3" → {2, 1}.  Throws FormatError on invalid input.
    static CellAddress from_a1(const std::string& a1);

    /// Convert back to "B3" form.
    std::string to_a1() const;
};

// ============================================================
//  Cell metadata
// ============================================================
enum class CellType { Blank, Number, String, Boolean, Error, Formula };

struct FormatInfo {
    uint16_t    format_index{0};
    std::string format_string;      // e.g. "YYYY-MM-DD", "0.00%"
    bool        is_date_format{false};
    bool        is_time_format{false};
};

struct CellStyle {
    std::optional<FormatInfo> format;
};

struct Cell {
    CellAddress                address;
    CellType                   type{CellType::Blank};
    CellValue                  value;
    std::optional<std::string> formula;   // raw formula string if present
    CellStyle                  style;
    bool                       merged{false};

    bool is_blank()    const { return type == CellType::Blank; }
    bool has_formula() const { return formula.has_value(); }
};

// ============================================================
//  Sheet interface
// ============================================================
class Sheet {
public:
    virtual ~Sheet() = default;

    virtual std::string name()      const = 0;
    virtual uint32_t    row_count() const = 0;
    virtual uint32_t    col_count() const = 0;

    // ---- Read ----
    /// Returns a blank Cell (not an error) if the address is empty or out of range.
    virtual Cell cell(uint32_t row, uint32_t col)  const = 0;
    virtual Cell cell(const CellAddress& addr)      const = 0;
    virtual Cell cell(const std::string& a1)        const = 0;

    /// Returns nullopt for blank or out-of-range cells.
    virtual std::optional<Cell> try_cell(uint32_t row, uint32_t col) const = 0;

    /// All cells in one row (sorted by column).
    virtual std::vector<Cell> row(uint32_t row_idx) const = 0;

    /// All non-blank cells (row-major order).
    virtual std::vector<Cell> cells() const = 0;

    virtual void for_each_cell(std::function<void(const Cell&)> fn) const = 0;

    // ---- Write ----
    virtual void set_cell   (uint32_t row, uint32_t col, const CellValue& v) = 0;
    virtual void set_cell   (const std::string& a1,      const CellValue& v) = 0;
    virtual void set_formula(const std::string& a1, const std::string& formula) = 0;

    // ---- Bulk helpers ----
    /// Write a 2-D table starting at top_left.  Each inner vector is one row.
    void set_table(const std::string& top_left,
                   const std::vector<std::vector<CellValue>>& data);
};

// ============================================================
//  Workbook interface
// ============================================================
enum class FileFormat { XLS, XLSX, Auto };

struct OpenOptions {
    bool strict_validation  {true};   // throw on any format violation
    bool preserve_formulas  {true};
    bool preserve_styles    {true};
    bool fail_on_unsupported{false};  // throw on unrecognised features
};

struct SaveOptions {
    FileFormat format    {FileFormat::Auto}; // Auto = keep original format
    bool       recalculate{false};
};

class Workbook {
public:
    virtual ~Workbook() = default;

    virtual FileFormat   format()      const = 0;
    virtual std::string  format_name() const = 0;
    virtual size_t       sheet_count() const = 0;

    virtual Sheet&       sheet(size_t index)             = 0;
    virtual const Sheet& sheet(size_t index)       const = 0;
    virtual Sheet&       sheet(const std::string& name)  = 0;
    virtual std::vector<std::string> sheet_names() const = 0;

    virtual Sheet& add_sheet   (const std::string& name)              = 0;
    virtual void   remove_sheet(size_t index)                         = 0;
    virtual void   rename_sheet(size_t index, const std::string& name)= 0;

    virtual void                 save    (const std::string& path,
                                          const SaveOptions& opts = {}) const = 0;
    virtual std::vector<uint8_t> to_bytes(FileFormat fmt,
                                          const SaveOptions& opts = {}) const = 0;
};

// ============================================================
//  WorkbookFactory
// ============================================================
class WorkbookFactory {
public:
    /// Open from file path.  Format is detected from magic bytes, never from extension.
    static std::unique_ptr<Workbook> open(const std::string& path,
                                          const OpenOptions& opts = {});

    /// Open from an in-memory buffer.
    static std::unique_ptr<Workbook> open(const std::vector<uint8_t>& data,
                                          const OpenOptions& opts = {});

    /// Create a new empty workbook.
    static std::unique_ptr<Workbook> create(FileFormat fmt = FileFormat::XLSX);
};

} // namespace excellib
