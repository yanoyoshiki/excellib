#include "excellib/workbook.hpp"
#include "excellib/xls_parser.hpp"
#include "excellib/xlsx_parser.hpp"
#include <fstream>
#include <iterator>
#include <cstring>

namespace excellib {

static FileFormat detect_format(const std::vector<uint8_t>& data) {
    if (data.size() < 8)
        throw ParseError("File too small to detect format (< 8 bytes)");

    static constexpr uint8_t OLE2[8] = {0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1};
    if (std::memcmp(data.data(), OLE2, 8) == 0)
        return FileFormat::XLS;

    if (data[0]==0x50 && data[1]==0x4B && data[2]==0x03 && data[3]==0x04)
        return FileFormat::XLSX;

    throw FormatError("Unknown file format. "
                      "First 4 bytes: "
                      + std::to_string(int(data[0])) + " "
                      + std::to_string(int(data[1])) + " "
                      + std::to_string(int(data[2])) + " "
                      + std::to_string(int(data[3])));
}

std::unique_ptr<Workbook> WorkbookFactory::open(
    const std::string& path, const OpenOptions& opts)
{
    std::ifstream f(path, std::ios::binary);
    if (!f) throw IOError("Cannot open file: " + path);
    std::vector<uint8_t> data(
        (std::istreambuf_iterator<char>(f)), std::istreambuf_iterator<char>());
    if (data.empty()) throw IOError("File is empty: " + path);
    return open(data, opts);
}

std::unique_ptr<Workbook> WorkbookFactory::open(
    const std::vector<uint8_t>& data, const OpenOptions& opts)
{
    switch (detect_format(data)) {
        case FileFormat::XLS: {
            xls::XlsParser p(opts); return p.parse(data);
        }
        case FileFormat::XLSX: {
            xlsx::XlsxParser p(opts); return p.parse(data);
        }
        default: throw FormatError("Unsupported format");
    }
}

std::unique_ptr<Workbook> WorkbookFactory::create(FileFormat fmt) {
    switch (fmt) {
        case FileFormat::XLSX: return std::make_unique<xlsx::XlsxWorkbook>();
        case FileFormat::XLS:  return std::make_unique<xls::XlsWorkbook>();
        default:               return std::make_unique<xlsx::XlsxWorkbook>();
    }
}

} // namespace excellib
