#include "excellib/workbook.hpp"
#include "excellib/print_settings.hpp"
#include <cctype>
#include <algorithm>
#include <sstream>

namespace excellib {

// ============================================================
//  CellAddress
// ============================================================
CellAddress CellAddress::from_a1(const std::string& a1) {
    if (a1.empty())
        throw FormatError("Empty A1 reference");

    uint32_t col = 0;
    size_t   i   = 0;

    while (i < a1.size() && std::isalpha(static_cast<unsigned char>(a1[i]))) {
        char c = static_cast<char>(std::toupper(static_cast<unsigned char>(a1[i])));
        col = col * 26 + static_cast<uint32_t>(c - 'A' + 1);
        ++i;
    }
    if (i == 0)
        throw FormatError("No column letters in A1 reference: " + a1);
    if (col == 0 || col > 16384)
        throw FormatError("Column out of Excel range in: " + a1);

    if (i == a1.size())
        throw FormatError("No row digits in A1 reference: " + a1);

    uint32_t row = 0;
    while (i < a1.size()) {
        unsigned char c = static_cast<unsigned char>(a1[i]);
        if (!std::isdigit(c))
            throw FormatError("Non-digit in row part of A1 reference: " + a1);
        row = row * 10 + static_cast<uint32_t>(c - '0');
        ++i;
    }
    if (row == 0 || row > 1048576)
        throw FormatError("Row out of Excel range in: " + a1);

    return CellAddress{row - 1, col - 1};
}

std::string CellAddress::to_a1() const {
    std::string col_str;
    uint32_t c = col + 1;
    while (c > 0) {
        --c;
        col_str += static_cast<char>('A' + (c % 26));
        c /= 26;
    }
    std::reverse(col_str.begin(), col_str.end());
    return col_str + std::to_string(row + 1);
}

// ============================================================
//  CellValue to_string
// ============================================================
std::string to_string(const CellValue& v) {
    return std::visit([](auto&& val) -> std::string {
        using T = std::decay_t<decltype(val)>;
        if constexpr (std::is_same_v<T, BlankValue>)
            return "";
        else if constexpr (std::is_same_v<T, bool>)
            return val ? "TRUE" : "FALSE";
        else if constexpr (std::is_same_v<T, int64_t>)
            return std::to_string(val);
        else if constexpr (std::is_same_v<T, double>) {
            std::ostringstream os;
            os << val;
            return os.str();
        }
        else if constexpr (std::is_same_v<T, std::string>)
            return val;
        else if constexpr (std::is_same_v<T, ErrorValue>)
            return val.code;
    }, v);
}

// ============================================================
//  Sheet::set_table helper
// ============================================================
void Sheet::set_table(const std::string& top_left,
                      const std::vector<std::vector<CellValue>>& data) {
    CellAddress origin = CellAddress::from_a1(top_left);
    for (size_t r = 0; r < data.size(); ++r)
        for (size_t c = 0; c < data[r].size(); ++c)
            set_cell(static_cast<uint32_t>(origin.row + r),
                     static_cast<uint32_t>(origin.col + c),
                     data[r][c]);
}

// ============================================================
//  CellRange
// ============================================================
CellRange CellRange::from_a1(std::string_view range_str) {
    std::string s(range_str);
    auto colon = s.find(':');
    if (colon == std::string::npos) {
        auto addr = CellAddress::from_a1(s);
        return CellRange{addr.row, addr.col, addr.row, addr.col};
    }
    auto a = CellAddress::from_a1(s.substr(0, colon));
    auto b = CellAddress::from_a1(s.substr(colon + 1));
    return CellRange{
        std::min(a.row, b.row), std::min(a.col, b.col),
        std::max(a.row, b.row), std::max(a.col, b.col)
    };
}

std::string CellRange::to_a1() const {
    return CellAddress{row1, col1}.to_a1() + ":" + CellAddress{row2, col2}.to_a1();
}

// ============================================================
//  PrintArea
// ============================================================
PrintArea PrintArea::from_range(const std::string& range) {
    auto colon = range.find(':');
    PrintArea pa;
    if (colon == std::string::npos) {
        auto addr = CellAddress::from_a1(range);
        pa.first_row = pa.last_row = addr.row;
        pa.first_col = pa.last_col = addr.col;
    } else {
        auto s = CellAddress::from_a1(range.substr(0, colon));
        auto e = CellAddress::from_a1(range.substr(colon + 1));
        pa.first_row = std::min(s.row, e.row);
        pa.first_col = std::min(s.col, e.col);
        pa.last_row  = std::max(s.row, e.row);
        pa.last_col  = std::max(s.col, e.col);
    }
    return pa;
}

std::string PrintArea::to_range() const {
    return CellAddress{first_row, first_col}.to_a1()
         + ":" +
           CellAddress{last_row, last_col}.to_a1();
}

} // namespace excellib
