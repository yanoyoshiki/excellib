#pragma once
#include "excellib/workbook.hpp"
#include <vector>
#include <unordered_map>
#include <map>
#include <cstdint>
#include <string>
#include <memory>
#include <iostream>

namespace excellib::xls {

// ---- BIFF8 record type codes ----
enum class RecordType : uint16_t {
    BOF         = 0x0809,
    EOF_        = 0x000A,
    BOUNDSHEET  = 0x0085,
    SST         = 0x00FC,
    CONTINUE    = 0x003C,
    LABELSST    = 0x00FD,
    LABEL       = 0x0204,
    NUMBER      = 0x0203,
    BOOLERR     = 0x0205,
    BLANK       = 0x0201,
    MULBLANK    = 0x00BE,
    MULRK       = 0x00BD,
    RK          = 0x027E,
    FORMULA     = 0x0006,
    STRING_     = 0x0207,
    XF          = 0x00E0,
    FORMAT      = 0x041E,
    DIMENSION   = 0x0200,
    ROW         = 0x0208,
    MERGEDCELLS = 0x00E5,
    UNKNOWN     = 0xFFFF,
};

struct BiffRecord {
    RecordType           type{RecordType::UNKNOWN};
    uint16_t             raw_type{0};
    std::vector<uint8_t> data;

    size_t size() const { return data.size(); }

    // All bounds-checked; throw ParseError on overrun
    uint8_t  u8 (size_t offset) const;
    uint16_t u16(size_t offset) const;
    uint32_t u32(size_t offset) const;
    int16_t  i16(size_t offset) const;
    double   f64(size_t offset) const;
};

double decode_rk(uint32_t rk);

std::string decode_biff8_string(const uint8_t* data, size_t available,
                                 size_t& bytes_consumed,
                                 bool short_header = false);

class SharedStringTable {
public:
    void parse(const BiffRecord& sst_record,
               const std::vector<BiffRecord>& continues);
    const std::string& get(uint32_t index) const;
    size_t size() const { return strings_.size(); }
private:
    std::vector<std::string> strings_;
};

class FormatRegistry {
public:
    FormatRegistry();
    void       add(uint16_t index, const std::string& fmt_string);
    FormatInfo get(uint16_t index) const;
    bool       is_date_format(uint16_t index) const;
private:
    std::unordered_map<uint16_t, std::string> formats_;
    static bool looks_like_date(const std::string& fmt);
};

struct XFRecord {
    uint16_t font_index  {0};
    uint16_t format_index{0};
    bool     is_style_xf {false};
};

class XFTable {
public:
    void             add(const XFRecord& xf);
    const XFRecord&  get(uint16_t index) const;
    size_t           size() const { return records_.size(); }
private:
    std::vector<XFRecord> records_;
};

// ---- Sheet ----
class XlsSheet : public Sheet {
public:
    explicit XlsSheet(std::string name) : name_(std::move(name)) {}

    std::string name()      const override { return name_; }
    uint32_t    row_count() const override { return row_count_; }
    uint32_t    col_count() const override { return col_count_; }

    Cell                cell(uint32_t row, uint32_t col)  const override;
    Cell                cell(const CellAddress& addr)      const override;
    Cell                cell(const std::string& a1)        const override;
    std::optional<Cell> try_cell(uint32_t row, uint32_t col) const override;
    std::vector<Cell>   row(uint32_t row_idx)              const override;
    std::vector<Cell>   col(uint32_t col_idx)              const override;
    std::vector<Cell>   cells()                            const override;
    void for_each_cell(std::function<void(const Cell&)> fn) const override;

    void set_cell   (uint32_t row, uint32_t col, const CellValue& v) override;
    void set_cell   (const std::string& a1,      const CellValue& v) override;
    void set_formula(const std::string& a1, const std::string& formula) override;
    void set_row    (uint32_t row_idx, const std::vector<CellValue>& values) override;

    void merge  (const CellRange& range) override;
    void unmerge(const CellRange& range) override;
    std::vector<CellRange> merged_ranges() const override;

    void put_cell       (const Cell& c);
    void set_dimensions (uint32_t rows, uint32_t cols);

    std::vector<CellRange> merges_;

private:
    std::string name_;
    uint32_t    row_count_{0};
    uint32_t    col_count_{0};
    std::map<uint32_t, std::map<uint32_t, Cell>> data_;

    static void fill_type(Cell& c);
};

// ---- Workbook ----
class XlsWorkbook : public Workbook {
public:
    FileFormat   format()      const override { return FileFormat::XLS; }
    std::string  format_name() const override { return "XLS"; }
    size_t       sheet_count() const override { return sheets_.size(); }

    Sheet&       sheet(size_t index)            override;
    const Sheet& sheet(size_t index)      const override;
    Sheet&       sheet(const std::string& name) override;
    std::vector<std::string> sheet_names() const override;

    Sheet& add_sheet   (const std::string& name)              override;
    void   remove_sheet(size_t index)                         override;
    void   rename_sheet(size_t index, const std::string& name)override;

    void                 save    (const std::string& path,
                                  const SaveOptions& opts = {}) const override;
    std::vector<uint8_t> to_bytes(FileFormat fmt,
                                  const SaveOptions& opts = {}) const override;

    void add_parsed_sheet(std::unique_ptr<XlsSheet> s);
private:
    std::vector<std::unique_ptr<XlsSheet>> sheets_;
};

// ---- Parser ----
class XlsParser {
public:
    explicit XlsParser(const OpenOptions& opts) : opts_(opts) {}
    std::unique_ptr<XlsWorkbook> parse(const std::vector<uint8_t>& data);

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

    std::vector<uint8_t> extract_workbook_stream(const std::vector<uint8_t>& raw);
    std::vector<BiffRecord> read_records(const std::vector<uint8_t>& stream);

    struct GlobalContext {
        SharedStringTable sst;
        XFTable           xf_table;
        FormatRegistry    formats;
        std::vector<std::pair<std::string, uint32_t>> sheets;
    };

    GlobalContext parse_globals(const std::vector<BiffRecord>& records);

    CellStyle xf_to_style(uint16_t xf_index, const XFTable& xf,
                           const FormatRegistry& fmt);

    Cell handle_number  (const BiffRecord& r, const GlobalContext& ctx);
    Cell handle_labelsst(const BiffRecord& r, const GlobalContext& ctx);
    Cell handle_label   (const BiffRecord& r);
    Cell handle_boolerr (const BiffRecord& r);
    Cell handle_rk      (const BiffRecord& r, const GlobalContext& ctx);
    std::vector<Cell> handle_mulrk(const BiffRecord& r, const GlobalContext& ctx);
    Cell handle_formula (const BiffRecord& r, const GlobalContext& ctx);
};

} // namespace excellib::xls
