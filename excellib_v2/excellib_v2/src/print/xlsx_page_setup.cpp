#include "excellib/excel_printer.hpp"
#include "excellib/xlsx_parser.hpp"
#include <vector>
#include <string>
#include <sstream>
#include <fstream>
#include <cstring>
#include <algorithm>
#include <stdexcept>
#include <zlib.h>

namespace excellib {

// ============================================================
//  ZIP read/write helpers (reuse ZipReader internals)
// ============================================================
static uint32_t read_le32(const uint8_t* p) {
    return uint32_t(p[0])|(uint32_t(p[1])<<8)|(uint32_t(p[2])<<16)|(uint32_t(p[3])<<24);
}
static uint16_t read_le16(const uint8_t* p) {
    return uint16_t(p[0])|(uint16_t(p[1])<<8);
}
static void append_le16(std::vector<uint8_t>& v, uint16_t x) {
    v.push_back(x&0xFF); v.push_back((x>>8)&0xFF);
}
static void append_le32(std::vector<uint8_t>& v, uint32_t x) {
    v.push_back(x&0xFF); v.push_back((x>>8)&0xFF);
    v.push_back((x>>16)&0xFF); v.push_back((x>>24)&0xFF);
}

struct ZipEntry2 {
    std::string           name;
    std::vector<uint8_t>  data;
    uint32_t local_offset{0};
    uint32_t crc{0};
};

static std::vector<uint8_t> read_file(const std::string& path) {
    std::ifstream f(path, std::ios::binary);
    if (!f) throw IOError("Cannot read: " + path);
    return {std::istreambuf_iterator<char>(f), {}};
}

static void write_file(const std::string& path, const std::vector<uint8_t>& data) {
    std::ofstream f(path, std::ios::binary);
    if (!f) throw IOError("Cannot write: " + path);
    f.write(reinterpret_cast<const char*>(data.data()),
            static_cast<std::streamsize>(data.size()));
}

// Parse entire ZIP into entries map
static std::vector<ZipEntry2> zip_read_all(const std::vector<uint8_t>& data) {
    std::vector<ZipEntry2> entries;
    if (data.size() < 22) throw ParseError("ZIP too small");

    // Find EOCD
    size_t eocd = std::string::npos;
    for (size_t i = data.size()-22; i > 0; --i)
        if (data[i]==0x50&&data[i+1]==0x4B&&data[i+2]==0x05&&data[i+3]==0x06)
            { eocd=i; break; }
    if (eocd==std::string::npos) throw ParseError("ZIP EOCD not found");

    uint32_t cd_off = read_le32(data.data()+eocd+16);
    uint16_t count  = read_le16(data.data()+eocd+10);
    size_t pos = cd_off;

    for (uint16_t i=0; i<count; ++i) {
        if (pos+46>data.size()) throw ParseError("CD truncated");
        uint16_t method   = read_le16(data.data()+pos+10);
        uint32_t comp_sz  = read_le32(data.data()+pos+20);
        uint32_t uncomp_sz= read_le32(data.data()+pos+24);
        uint16_t fn_len   = read_le16(data.data()+pos+28);
        uint16_t ex_len   = read_le16(data.data()+pos+30);
        uint16_t cm_len   = read_le16(data.data()+pos+32);
        uint32_t lhdr_off = read_le32(data.data()+pos+42);
        std::string name(reinterpret_cast<const char*>(data.data()+pos+46), fn_len);
        pos += 46+fn_len+ex_len+cm_len;

        uint16_t lfn = read_le16(data.data()+lhdr_off+26);
        uint16_t lex = read_le16(data.data()+lhdr_off+28);
        size_t doff  = lhdr_off+30+lfn+lex;

        ZipEntry2 e; e.name=name;
        if (method==0) {
            e.data.assign(data.data()+doff, data.data()+doff+comp_sz);
        } else if (method==8) {
            e.data.resize(uncomp_sz);
            z_stream zs{}; zs.next_in=const_cast<Bytef*>(data.data()+doff);
            zs.avail_in=comp_sz; zs.next_out=e.data.data(); zs.avail_out=uncomp_sz;
            inflateInit2(&zs,-MAX_WBITS); ::inflate(&zs,Z_FINISH); inflateEnd(&zs);
            e.data.resize(zs.total_out);
        }
        entries.push_back(std::move(e));
    }
    return entries;
}

static std::vector<uint8_t> zip_write_all(std::vector<ZipEntry2>& entries) {
    std::vector<uint8_t> out;
    for (auto& e : entries) {
        e.local_offset = uint32_t(out.size());
        e.crc = uint32_t(crc32(0L, e.data.data(), uInt(e.data.size())));
        append_le32(out,0x04034B50); append_le16(out,20); append_le16(out,0);
        append_le16(out,0); append_le16(out,0); append_le16(out,0);
        append_le32(out,e.crc);
        append_le32(out,uint32_t(e.data.size()));
        append_le32(out,uint32_t(e.data.size()));
        append_le16(out,uint16_t(e.name.size())); append_le16(out,0);
        out.insert(out.end(),e.name.begin(),e.name.end());
        out.insert(out.end(),e.data.begin(),e.data.end());
    }
    uint32_t cd_off = uint32_t(out.size());
    for (auto& e : entries) {
        append_le32(out,0x02014B50); append_le16(out,20); append_le16(out,20);
        append_le16(out,0); append_le16(out,0); append_le16(out,0); append_le16(out,0);
        append_le32(out,e.crc);
        append_le32(out,uint32_t(e.data.size()));
        append_le32(out,uint32_t(e.data.size()));
        append_le16(out,uint16_t(e.name.size()));
        append_le16(out,0); append_le16(out,0); append_le16(out,0);
        append_le16(out,0); append_le32(out,0); append_le32(out,e.local_offset);
        out.insert(out.end(),e.name.begin(),e.name.end());
    }
    uint32_t cd_sz = uint32_t(out.size())-cd_off;
    append_le32(out,0x06054B50); append_le16(out,0); append_le16(out,0);
    append_le16(out,uint16_t(entries.size()));
    append_le16(out,uint16_t(entries.size()));
    append_le32(out,cd_sz); append_le32(out,cd_off); append_le16(out,0);
    return out;
}

// ============================================================
//  XML helpers
// ============================================================
static std::string xml_attr(const std::string& el, const std::string& attr) {
    auto pos = el.find(attr+"=\"");
    if (pos==std::string::npos) return {};
    pos += attr.size()+2;
    auto end = el.find('"',pos);
    return end==std::string::npos ? "" : el.substr(pos,end-pos);
}

static std::string to_str(double v) {
    // Format double without trailing zeros
    char buf[32]; std::snprintf(buf,sizeof(buf),"%.4f",v);
    std::string s(buf);
    // trim trailing zeros
    auto dot = s.find('.');
    if (dot!=std::string::npos) {
        size_t last = s.find_last_not_of('0');
        if (last==dot) ++last;
        s = s.substr(0,last+1);
    }
    return s;
}

// ============================================================
//  Build pageSetup XML element from PageSetup struct
// ============================================================
static std::string build_page_setup_xml(const PageSetup& ps) {
    std::ostringstream o;
    o << "<pageSetup";
    o << " paperSize=\"" << uint16_t(ps.paper_size) << "\"";

    if (ps.orientation == Orientation::Landscape)
        o << " orientation=\"landscape\"";
    else
        o << " orientation=\"portrait\"";

    if (ps.fit_to != FitTo::None) {
        o << " fitToPage=\"1\"";
        if (ps.fit_to==FitTo::Width || ps.fit_to==FitTo::WidthAndHeight)
            o << " fitToWidth=\"" << ps.fit_to_pages_wide << "\"";
        else
            o << " fitToWidth=\"0\"";
        if (ps.fit_to==FitTo::Height || ps.fit_to==FitTo::WidthAndHeight)
            o << " fitToHeight=\"" << ps.fit_to_pages_tall << "\"";
        else
            o << " fitToHeight=\"0\"";
    } else {
        o << " scale=\"" << ps.scale_percent << "\"";
        o << " fitToPage=\"0\"";
    }

    if (ps.first_page_number > 0)
        o << " firstPageNumber=\"" << ps.first_page_number
          << "\" useFirstPageNumber=\"1\"";

    if (ps.black_and_white)    o << " blackAndWhite=\"1\"";
    if (ps.draft_quality)      o << " draft=\"1\"";
    if (ps.page_order_over_then_down) o << " pageOrder=\"overThenDown\"";
    if (ps.copies > 1)         o << " copies=\"" << ps.copies << "\"";

    o << "/>";
    return o.str();
}

static std::string build_page_margins_xml(const PageMargins& m) {
    std::ostringstream o;
    o << "<pageMargins"
      << " left=\""   << to_str(m.left)   << "\""
      << " right=\""  << to_str(m.right)  << "\""
      << " top=\""    << to_str(m.top)    << "\""
      << " bottom=\"" << to_str(m.bottom) << "\""
      << " header=\"" << to_str(m.header) << "\""
      << " footer=\"" << to_str(m.footer) << "\""
      << "/>";
    return o.str();
}

static std::string build_print_options_xml(const PageSetup& ps) {
    if (!ps.print_gridlines && !ps.print_row_col_headings) return "";
    std::ostringstream o;
    o << "<printOptions";
    if (ps.print_gridlines)         o << " gridLines=\"1\" gridLinesSet=\"1\"";
    if (ps.print_row_col_headings)  o << " headings=\"1\"";
    o << "/>";
    return o.str();
}

static std::string build_header_footer_xml(const HeaderFooter& hf) {
    if (hf.odd_header.empty() && hf.odd_footer.empty() &&
        hf.even_header.empty() && hf.even_footer.empty()) return "";
    std::ostringstream o;
    o << "<headerFooter";
    if (hf.different_odd_even) o << " differentOddEven=\"1\"";
    if (hf.different_first)    o << " differentFirst=\"1\"";
    o << ">";
    if (!hf.odd_header.empty())  o << "<oddHeader>"  << hf.odd_header  << "</oddHeader>";
    if (!hf.odd_footer.empty())  o << "<oddFooter>"  << hf.odd_footer  << "</oddFooter>";
    if (!hf.even_header.empty()) o << "<evenHeader>" << hf.even_header << "</evenHeader>";
    if (!hf.even_footer.empty()) o << "<evenFooter>" << hf.even_footer << "</evenFooter>";
    o << "</headerFooter>";
    return o.str();
}

// ============================================================
//  シートXMLにページ設定を注入
// ============================================================
static std::string inject_page_setup(const std::string& sheet_xml, const PageSetup& ps) {
    std::string xml = sheet_xml;

    // 既存のpageSetup / pageMargins / printOptions / headerFooter を削除
    auto remove_element = [&](const std::string& tag) {
        // self-closing: <tag ... />
        while (true) {
            auto s = xml.find("<" + tag);
            if (s == std::string::npos) break;
            auto e = xml.find("/>", s);
            auto e2 = xml.find("</" + tag + ">", s);
            size_t end_pos;
            if (e != std::string::npos && (e2 == std::string::npos || e < e2))
                end_pos = e + 2;
            else if (e2 != std::string::npos)
                end_pos = e2 + tag.size() + 3;
            else break;
            xml.erase(s, end_pos - s);
        }
    };

    remove_element("pageSetup");
    remove_element("pageMargins");
    remove_element("printOptions");
    remove_element("headerFooter");

    // 印刷範囲 (sheetView/定義名ではなく printArea として)
    // → workbook.xml の definedNames に書くのが正式だが
    //    sheet.xml の sheetView に sqref として入れる方法もある
    // ここでは sheet_xml にマーカーを挿入し、workbook 側で definedNames を設定

    // </sheetData> の直後に挿入
    std::string insert_point = "</sheetData>";
    auto ins = xml.find(insert_point);
    if (ins == std::string::npos) {
        // sheetData が self-closing の場合
        insert_point = "<sheetData/>";
        ins = xml.find(insert_point);
        if (ins != std::string::npos) ins += insert_point.size() - 2;
    } else {
        ins += insert_point.size();
    }

    std::string additions;
    additions += build_print_options_xml(ps);
    additions += build_page_margins_xml(ps.margins);
    additions += build_page_setup_xml(ps);
    additions += build_header_footer_xml(ps.header_footer);

    if (ins != std::string::npos)
        xml.insert(ins, additions);
    else
        xml += additions;  // fallback

    return xml;
}

// ============================================================
//  シート名 → ZIP内パス解決
// ============================================================
static std::string resolve_sheet_path(
    const std::vector<ZipEntry2>& entries,
    const std::string& sheet_name)
{
    // workbook.xml を探す
    std::string wb_xml;
    for (auto& e : entries)
        if (e.name == "xl/workbook.xml")
            { wb_xml = std::string(e.data.begin(), e.data.end()); break; }
    if (wb_xml.empty()) throw ParseError("workbook.xml not found");

    // rels を探す
    std::string rels_xml;
    for (auto& e : entries)
        if (e.name == "xl/_rels/workbook.xml.rels")
            { rels_xml = std::string(e.data.begin(), e.data.end()); break; }

    // sheet名 → rId
    std::string target_rid;
    size_t pos = 0;
    while (true) {
        auto sh = wb_xml.find("<sheet ", pos);
        if (sh == std::string::npos) break;
        auto she = wb_xml.find('>', sh);
        std::string elem = wb_xml.substr(sh, she-sh+1);
        std::string name = xml_attr(elem, "name");
        if (name == sheet_name || sheet_name.empty()) {
            target_rid = xml_attr(elem, "r:id");
            break;
        }
        pos = she;
    }
    if (target_rid.empty() && !sheet_name.empty())
        throw RangeError("Sheet not found: " + sheet_name);
    if (target_rid.empty()) {
        // アクティブシート = 最初のシート
        auto sh = wb_xml.find("<sheet ");
        if (sh == std::string::npos) throw ParseError("No sheets in workbook");
        auto she = wb_xml.find('>', sh);
        target_rid = xml_attr(wb_xml.substr(sh, she-sh+1), "r:id");
    }

    // rId → target path
    pos = 0;
    while (true) {
        auto r = rels_xml.find("<Relationship ", pos);
        if (r == std::string::npos) break;
        auto re = rels_xml.find('>', r);
        std::string elem = rels_xml.substr(r, re-r+1);
        if (xml_attr(elem, "Id") == target_rid) {
            std::string target = xml_attr(elem, "Target");
            if (!target.empty() && target[0] == '/')
                return "xl" + target;
            return "xl/" + target;
        }
        pos = re;
    }
    throw ParseError("Could not resolve sheet path for rId=" + target_rid);
}

// ============================================================
//  workbook.xml に definedNames (印刷範囲・タイトル行) を設定
// ============================================================
static void inject_defined_names(
    std::string& wb_xml,
    const std::string& sheet_name,
    const PageSetup& ps)
{
    if (!ps.print_area && !ps.repeat_titles) return;

    // 既存の definedNames ブロックを削除
    auto dns = wb_xml.find("<definedNames");
    if (dns != std::string::npos) {
        auto dne = wb_xml.find("</definedNames>", dns);
        if (dne != std::string::npos)
            wb_xml.erase(dns, dne - dns + 15);
    }

    std::ostringstream dn;
    dn << "<definedNames>";

    // 印刷範囲
    if (ps.print_area) {
        std::string ref = "'" + sheet_name + "'!" +
            "$" + CellAddress{ps.print_area->first_row,
                              ps.print_area->first_col}.to_a1() +
            ":$" + CellAddress{ps.print_area->last_row,
                               ps.print_area->last_col}.to_a1();
        // A1 -> $A$1 変換
        // (簡易: to_a1 は "A1" を返すので $ を挿入)
        auto make_absolute = [](const std::string& a1) -> std::string {
            std::string r;
            size_t i=0;
            while (i<a1.size() && std::isalpha((unsigned char)a1[i]))
                { r += '$'; r += a1[i++]; }
            r += '$';
            while (i<a1.size()) r += a1[i++];
            return r;
        };
        std::string abs_start = make_absolute(
            CellAddress{ps.print_area->first_row, ps.print_area->first_col}.to_a1());
        std::string abs_end = make_absolute(
            CellAddress{ps.print_area->last_row, ps.print_area->last_col}.to_a1());
        dn << "<definedName name=\"_xlnm.Print_Area\">"
           << "'" << sheet_name << "'!" << abs_start << ":" << abs_end
           << "</definedName>";
    }

    // 繰り返し行・列タイトル
    if (ps.repeat_titles) {
        auto& rt = *ps.repeat_titles;
        std::string val;
        if (rt.row_start && rt.row_end) {
            val += "'" + sheet_name + "'!$" +
                   std::to_string(*rt.row_start + 1) + ":$" +
                   std::to_string(*rt.row_end + 1);
        }
        if (rt.col_start && rt.col_end) {
            if (!val.empty()) val += ",";
            std::string cs = CellAddress{0, *rt.col_start}.to_a1();
            std::string ce = CellAddress{0, *rt.col_end}.to_a1();
            // 列文字だけ取り出す
            auto col_letters = [](const std::string& a1) {
                std::string r;
                for (char c : a1) if (std::isalpha((unsigned char)c)) r+=c;
                return r;
            };
            val += "'" + sheet_name + "'!$" + col_letters(cs)
                + ":$" + col_letters(ce);
        }
        if (!val.empty())
            dn << "<definedName name=\"_xlnm.Print_Titles\">"
               << val << "</definedName>";
    }

    dn << "</definedNames>";

    // </workbook> の直前に挿入
    auto end = wb_xml.rfind("</workbook>");
    if (end != std::string::npos)
        wb_xml.insert(end, dn.str());
}

// ============================================================
//  ExcelPrinter::write_page_setup_to_xlsx  (static)
// ============================================================
void ExcelPrinter::write_page_setup_to_xlsx(
    const std::string& xlsx_path,
    const std::string& sheet_name,
    const PageSetup&   setup)
{
    auto raw     = read_file(xlsx_path);
    auto entries = zip_read_all(raw);

    // 対象シートパスを解決
    std::string sheet_path = resolve_sheet_path(entries, sheet_name);

    // シートXMLを更新
    for (auto& e : entries) {
        if (e.name == sheet_path) {
            std::string xml(e.data.begin(), e.data.end());
            xml = inject_page_setup(xml, setup);
            e.data = std::vector<uint8_t>(xml.begin(), xml.end());
        }
        // workbook.xml に definedNames を注入
        if (e.name == "xl/workbook.xml") {
            std::string xml(e.data.begin(), e.data.end());
            // 実際のシート名を取得（sheet_name が空の場合は最初のシート名）
            std::string sname = sheet_name;
            if (sname.empty()) {
                auto sh = xml.find("<sheet ");
                if (sh != std::string::npos) {
                    auto she = xml.find('>', sh);
                    sname = xml_attr(xml.substr(sh, she-sh+1), "name");
                }
            }
            inject_defined_names(xml, sname, setup);
            e.data = std::vector<uint8_t>(xml.begin(), xml.end());
        }
    }

    write_file(xlsx_path, zip_write_all(entries));
}

// ============================================================
//  apply_page_setup  (public)
// ============================================================
void ExcelPrinter::apply_page_setup(
    const std::string& xlsx_path,
    const std::string& sheet_name,
    const PageSetup&   setup)
{
    log("Applying page setup to: " + xlsx_path);
    write_page_setup_to_xlsx(xlsx_path, sheet_name, setup);
    log("Page setup applied.");
}

} // namespace excellib
