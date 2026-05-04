#pragma once
#include "excellib/workbook.hpp"
#include <vector>
#include <unordered_map>
#include <map>
#include <string>
#include <memory>
#include <functional>
#include <iostream>

namespace excellib::xlsx {

// ============================================================
//  ZIP entry
// ============================================================
struct ZipEntry {
    std::string              path;
    std::vector<uint8_t>     data;
};

// ============================================================
//  Minimal ZIP reader (no external dependency)
// ============================================================
class ZipReader {
public:
    explicit ZipReader(const std::vector<uint8_t>& data);

    bool              has(const std::string& path) const;
    const ZipEntry&   get(const std::string& path) const;
    std::vector<std::string> paths() const;

    // Get as string (decompresses)
    std::string text(const std::string& path) const;

private:
    std::unordered_map<std::string, ZipEntry> entries_;

    void parse(const std::vector<uint8_t>& data);
    static std::vector<uint8_t> inflate(const uint8_t* compressed, size_t comp_size,
                                         size_t uncomp_size);
};

// ============================================================
//  Relationship (.rels) parser
// ============================================================
struct Relationship {
    std::string id;
    std::string type;
    std::string target;
};

std::vector<Relationship> parse_rels(const std::string& xml);

// ============================================================
//  Shared Strings (xl/sharedStrings.xml)
// ============================================================
class XlsxSharedStrings {
public:
    void parse(const std::string& xml);
    const std::string& get(uint32_t index) const;
    size_t size() const { return strings_.size(); }

    // For writing
    uint32_t intern(const std::string& s);
    std::string to_xml() const;

private:
    std::vector<std::string>            strings_;
    std::unordered_map<std::string, uint32_t> index_map_;  // for write dedup
};

// ============================================================
//  Number formats & styles (xl/styles.xml)
// ============================================================
struct XlsxXF {
    uint16_t num_fmt_id{0};
    uint16_t font_id{0};
    uint16_t fill_id{0};
    uint16_t border_id{0};
    bool     apply_number_format{false};
};

class XlsxStyles {
public:
    void parse(const std::string& xml);

    FormatInfo cell_format(uint32_t xf_index) const;
    size_t     xf_count() const { return cell_xfs_.size(); }
    std::string to_xml() const;

private:
    std::unordered_map<uint16_t, std::string> num_fmts_;   // id -> format string
    std::vector<XlsxXF>                        cell_xfs_;
    static bool is_date_format(uint16_t id, const std::string& fmt);
    static bool looks_like_date(const std::string& fmt);
};

// ============================================================
//  XLSX Sheet implementation
// ============================================================
class XlsxSheet : public Sheet {
public:
    explicit XlsxSheet(std::string name) : name_(std::move(name)) {}

    std::string name() const override { return name_; }
    uint32_t    row_count() const override { return row_count_; }
    uint32_t    col_count() const override { return col_count_; }

    Cell cell(uint32_t row, uint32_t col) const override;
    Cell cell(const CellAddress& addr) const override;
    Cell cell(const std::string& a1) const override;
    std::optional<Cell> try_cell(uint32_t row, uint32_t col) const override;

    std::vector<Cell> row(uint32_t row_idx) const override;
    std::vector<Cell> col(uint32_t col_idx) const override;
    std::vector<Cell> cells() const override;

    void set_cell(uint32_t row, uint32_t col, const CellValue& value) override;
    void set_cell(const std::string& a1, const CellValue& value) override;
    void set_formula(const std::string& a1, const std::string& formula) override;
    void set_row(uint32_t row_idx, const std::vector<CellValue>& values) override;

    void merge  (const CellRange& range) override;
    void unmerge(const CellRange& range) override;
    std::vector<CellRange> merged_ranges() const override;

    void for_each_cell(std::function<void(const Cell&)> fn) const override;
    std::vector<Cell> read_until_blank(const CellAddress& start, Direction dir) const override;

    void put_cell(const Cell& c);
    void set_dimensions(uint32_t rows, uint32_t cols);

    std::vector<CellRange> merges_;

private:
    std::string name_;
    uint32_t    row_count_{0};
    uint32_t    col_count_{0};
    std::map<uint32_t, std::map<uint32_t, Cell>> data_;
};

// ============================================================
//  XLSX Workbook implementation
// ============================================================
class XlsxWorkbook : public Workbook {
public:
    FileFormat   format() const override { return FileFormat::XLSX; }
    std::string  format_name() const override { return "XLSX"; }
    size_t       sheet_count() const override { return sheets_.size(); }

    Sheet&       sheet(size_t index) override;
    const Sheet& sheet(size_t index) const override;
    Sheet&       sheet(const std::string& name) override;
    std::vector<std::string> sheet_names() const override;

    Sheet&       add_sheet(const std::string& name) override;
    void         remove_sheet(size_t index) override;
    void         rename_sheet(size_t index, const std::string& name) override;

    void         save(const std::string& path, const SaveOptions& opts = {}) const override;
    std::vector<uint8_t> to_bytes(FileFormat fmt, const SaveOptions& opts = {}) const override;

    void add_parsed_sheet(std::unique_ptr<XlsxSheet> sheet);

    std::map<std::string, std::vector<uint8_t>> passthrough_entries_;
    std::string original_styles_xml_;

private:
    std::vector<std::unique_ptr<XlsxSheet>> sheets_;
    XlsxSharedStrings                        sst_;
    XlsxStyles                               styles_;
};

// ============================================================
//  XLSX Parser
// ============================================================
class XlsxParser {
public:
    explicit XlsxParser(const OpenOptions& opts) : opts_(opts) {}

    std::unique_ptr<XlsxWorkbook> parse(const std::vector<uint8_t>& data);

    void warn(ParseWarning::Kind kind, const std::string& location, const std::string& msg) const {
        ParseWarning w{kind, location, msg};
        if (opts_.on_warning) {
            opts_.on_warning(w);
        } else {
            std::cerr << "[excellib warning] " << location << ": " << msg << "\n";
        }
    }

private:
    OpenOptions opts_;

    // Parse xl/workbook.xml -> sheet names + rIds
    struct SheetMeta { std::string name, rid, path; };
    std::vector<SheetMeta> parse_workbook_xml(
        const std::string& xml,
        const std::vector<Relationship>& rels
    );

    std::unique_ptr<XlsxSheet> parse_sheet_xml(
        const std::string& name,
        const std::string& xml,
        const XlsxSharedStrings& sst,
        const XlsxStyles& styles
    );

    // Cell value parsing
    Cell parse_cell_element(
        const std::string& ref,
        const std::string& type_attr,
        const std::string& s_attr,
        const std::string& v_text,
        const std::string& f_text,
        const XlsxSharedStrings& sst,
        const XlsxStyles& styles
    );

    // A1 -> {row, col}
    static CellAddress parse_ref(const std::string& ref);
};

// ============================================================
//  XLSX Writer
// ============================================================
class XlsxWriter {
public:
    std::vector<uint8_t> write(const XlsxWorkbook& wb, const SaveOptions& opts);

private:
    struct ZipBuilder;

    std::string sheet_xml(const XlsxSheet& sheet, XlsxSharedStrings& sst);
    std::string workbook_xml(const std::vector<std::string>& sheet_names);
    std::string workbook_rels_xml(size_t sheet_count);
    std::string content_types_xml(size_t sheet_count, bool has_shared_strings);
    std::string root_rels_xml();

    static std::string cell_ref(uint32_t row, uint32_t col);
    static std::string escape_xml(const std::string& s);
};

} // namespace excellib::xlsx
