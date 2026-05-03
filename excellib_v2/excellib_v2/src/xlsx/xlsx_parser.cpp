#include "excellib/xlsx_parser.hpp"
#include "../common/deflate.hpp"
#include <cstring>
#include <algorithm>
#include <sstream>
#include <fstream>
#include <stdexcept>

namespace excellib::xlsx {

// ============================================================
//  XML helpers (no external dependency)
// ============================================================
static std::string xml_attr(const std::string& el, const std::string& attr) {
    // Match 'attr="value"' — handles namespace prefixes correctly
    std::string needle = attr + "=\"";
    auto pos = el.find(needle);
    if (pos == std::string::npos) return {};
    pos += needle.size();
    auto end = el.find('"', pos);
    return end == std::string::npos ? std::string{} : el.substr(pos, end-pos);
}

static std::string xml_text(const std::string& xml, const std::string& tag) {
    std::string open  = "<" + tag + ">";
    std::string close = "</" + tag + ">";
    auto pos = xml.find(open);
    if (pos == std::string::npos) return {};
    pos += open.size();
    auto end = xml.find(close, pos);
    return end == std::string::npos ? std::string{} : xml.substr(pos, end-pos);
}

static std::string xml_unescape(const std::string& s) {
    std::string r;
    r.reserve(s.size());
    for (size_t i = 0; i < s.size(); ) {
        if (s[i] != '&') { r += s[i++]; continue; }
        if (s.substr(i,4) == "&lt;")   { r += '<';  i += 4; }
        else if (s.substr(i,4) == "&gt;")   { r += '>';  i += 4; }
        else if (s.substr(i,5) == "&amp;")  { r += '&';  i += 5; }
        else if (s.substr(i,6) == "&apos;") { r += '\''; i += 6; }
        else if (s.substr(i,6) == "&quot;") { r += '"';  i += 6; }
        else { r += s[i++]; }
    }
    return r;
}

// ============================================================
//  ZipReader — fixed EOCD search and stoul safety
// ============================================================
static uint32_t rd32(const uint8_t* p) {
    return uint32_t(p[0])|(uint32_t(p[1])<<8)|(uint32_t(p[2])<<16)|(uint32_t(p[3])<<24);
}
static uint16_t rd16(const uint8_t* p) {
    return uint16_t(p[0])|(uint16_t(p[1])<<8);
}

std::vector<uint8_t> ZipReader::inflate(const uint8_t* comp, size_t comp_sz, size_t uncomp_sz) {
    return excellib::detail::deflate_decompress(comp, comp_sz, uncomp_sz);
}

void ZipReader::parse(const std::vector<uint8_t>& data) {
    if (data.size() < 22) throw ParseError("ZIP file too small");

    // FIX: EOCD search — use signed index to correctly handle i==0
    size_t eocd = std::string::npos;
    // Start at last possible position and scan backwards
    for (ptrdiff_t i = static_cast<ptrdiff_t>(data.size()) - 22; i >= 0; --i) {
        auto j = static_cast<size_t>(i);
        if (data[j]==0x50 && data[j+1]==0x4B && data[j+2]==0x05 && data[j+3]==0x06) {
            eocd = j; break;
        }
    }
    if (eocd == std::string::npos) throw ParseError("ZIP: EOCD record not found");

    uint32_t cd_off   = rd32(data.data() + eocd + 16);
    uint16_t cd_count = rd16(data.data() + eocd + 10);

    if (cd_off > data.size()) throw ParseError("ZIP: central directory offset beyond EOF");

    size_t pos = cd_off;
    for (uint16_t i = 0; i < cd_count; ++i) {
        if (pos + 46 > data.size()) throw ParseError("ZIP: central directory truncated");
        if (rd32(data.data()+pos) != 0x02014B50) throw ParseError("ZIP: bad CD signature");

        uint16_t method    = rd16(data.data()+pos+10);
        uint32_t comp_sz   = rd32(data.data()+pos+20);
        uint32_t uncomp_sz = rd32(data.data()+pos+24);
        uint16_t fn_len    = rd16(data.data()+pos+28);
        uint16_t ex_len    = rd16(data.data()+pos+30);
        uint16_t cm_len    = rd16(data.data()+pos+32);
        uint32_t lhdr_off  = rd32(data.data()+pos+42);

        if (pos + 46 + fn_len > data.size()) throw ParseError("ZIP: filename beyond EOF");
        std::string fname(reinterpret_cast<const char*>(data.data()+pos+46), fn_len);
        pos += 46 + fn_len + ex_len + cm_len;

        if (lhdr_off + 30 > data.size()) throw ParseError("ZIP: local header OOB");
        uint16_t lfn = rd16(data.data()+lhdr_off+26);
        uint16_t lex = rd16(data.data()+lhdr_off+28);
        size_t   doff = lhdr_off + 30 + lfn + lex;

        if (doff + comp_sz > data.size()) throw ParseError("ZIP: file data OOB: " + fname);

        ZipEntry entry;
        entry.path = fname;
        if (method == 0) {
            entry.data.assign(data.data()+doff, data.data()+doff+comp_sz);
        } else if (method == 8) {
            entry.data = inflate(data.data()+doff, comp_sz, uncomp_sz);
        } else {
            throw ParseError("ZIP: unsupported compression method " + std::to_string(method));
        }
        entries_[fname] = std::move(entry);
    }
}

ZipReader::ZipReader(const std::vector<uint8_t>& data) { parse(data); }
bool ZipReader::has(const std::string& p) const { return entries_.count(p)>0; }
const ZipEntry& ZipReader::get(const std::string& p) const {
    auto it = entries_.find(p);
    if (it == entries_.end()) throw IOError("ZIP entry not found: " + p);
    return it->second;
}
std::string ZipReader::text(const std::string& p) const {
    auto& e = get(p);
    return {reinterpret_cast<const char*>(e.data.data()), e.data.size()};
}
std::vector<std::string> ZipReader::paths() const {
    std::vector<std::string> r;
    for (auto& [k,v]:entries_) r.push_back(k);
    return r;
}

// ============================================================
//  Relationships
// ============================================================
std::vector<Relationship> parse_rels(const std::string& xml) {
    std::vector<Relationship> out;
    size_t pos = 0;
    while (true) {
        auto r = xml.find("<Relationship ", pos);
        if (r == std::string::npos) break;
        auto re = xml.find('>', r);
        if (re == std::string::npos) break;
        std::string elem = xml.substr(r, re-r+1);
        out.push_back({xml_attr(elem,"Id"), xml_attr(elem,"Type"), xml_attr(elem,"Target")});
        pos = re+1;
    }
    return out;
}

// ============================================================
//  XlsxSharedStrings
// ============================================================
void XlsxSharedStrings::parse(const std::string& xml) {
    size_t pos = 0;
    while (true) {
        auto si = xml.find("<si>", pos);
        if (si == std::string::npos) break;
        auto sie = xml.find("</si>", si);
        if (sie == std::string::npos) break;
        std::string block = xml.substr(si, sie-si+5);
        std::string text_val;
        size_t tp = 0;
        while (true) {
            auto ts = block.find("<t", tp);
            if (ts == std::string::npos) break;
            auto te = block.find('>', ts);
            if (te == std::string::npos) break;
            auto tc = block.find("</t>", te);
            if (tc == std::string::npos) break;
            text_val += xml_unescape(block.substr(te+1, tc-te-1));
            tp = tc + 4;
        }
        strings_.push_back(text_val);
        pos = sie + 5;
    }
}

const std::string& XlsxSharedStrings::get(uint32_t idx) const {
    if (idx >= strings_.size())
        throw RangeError("SharedStrings index " + std::to_string(idx)
                         + " out of range (size=" + std::to_string(strings_.size()) + ")");
    return strings_[idx];
}

uint32_t XlsxSharedStrings::intern(const std::string& s) {
    auto it = index_map_.find(s);
    if (it != index_map_.end()) return it->second;
    auto idx = static_cast<uint32_t>(strings_.size());
    strings_.push_back(s);
    index_map_[s] = idx;
    return idx;
}

// ============================================================
//  XlsxStyles
// ============================================================
void XlsxStyles::parse(const std::string& xml) {
    // numFmts
    auto nf_start = xml.find("<numFmts");
    auto nf_end   = xml.find("</numFmts>");
    if (nf_start != std::string::npos && nf_end != std::string::npos) {
        std::string blk = xml.substr(nf_start, nf_end-nf_start);
        size_t pos = 0;
        while (true) {
            auto p = blk.find("<numFmt ", pos);
            if (p == std::string::npos) break;
            auto pe = blk.find('>', p);
            if (pe == std::string::npos) break;
            std::string el = blk.substr(p, pe-p+1);
            std::string id_s  = xml_attr(el, "numFmtId");
            std::string fmt_s = xml_attr(el, "formatCode");
            // FIX: guard against empty/non-numeric strings before stoi
            if (!id_s.empty()) {
                try { num_fmts_[static_cast<uint16_t>(std::stoi(id_s))] = fmt_s; }
                catch (...) {}
            }
            pos = pe+1;
        }
    }
    // cellXfs
    auto xf_start = xml.find("<cellXfs");
    auto xf_end   = xml.find("</cellXfs>");
    if (xf_start != std::string::npos && xf_end != std::string::npos) {
        std::string blk = xml.substr(xf_start, xf_end-xf_start);
        size_t pos = 0;
        while (true) {
            auto p = blk.find("<xf ", pos);
            if (p == std::string::npos) break;
            auto pe = blk.find('>', p);
            if (pe == std::string::npos) break;
            std::string el = blk.substr(p, pe-p+1);
            XlsxXF xf;
            std::string nfid = xml_attr(el, "numFmtId");
            if (!nfid.empty()) {
                try { xf.num_fmt_id = static_cast<uint16_t>(std::stoi(nfid)); } catch (...) {}
            }
            std::string apnf = xml_attr(el, "applyNumberFormat");
            xf.apply_number_format = (apnf == "1" || apnf == "true");
            cell_xfs_.push_back(xf);
            pos = pe+1;
        }
    }
}

FormatInfo XlsxStyles::cell_format(uint32_t xf_idx) const {
    FormatInfo fi;
    if (xf_idx >= cell_xfs_.size()) return fi;
    auto& xf = cell_xfs_[xf_idx];
    fi.format_index   = xf.num_fmt_id;
    auto it = num_fmts_.find(xf.num_fmt_id);
    if (it != num_fmts_.end()) fi.format_string = it->second;
    fi.is_date_format = is_date_format(xf.num_fmt_id, fi.format_string);
    return fi;
}

bool XlsxStyles::is_date_format(uint16_t id, const std::string& fmt) {
    if ((id >= 14 && id <= 22) || (id >= 45 && id <= 47)) return true;
    return looks_like_date(fmt);
}
bool XlsxStyles::looks_like_date(const std::string& fmt) {
    std::string lo = fmt;
    std::transform(lo.begin(), lo.end(), lo.begin(), ::tolower);
    return lo.find('y') != std::string::npos
        || (lo.find('d') != std::string::npos && lo.find('h') == std::string::npos);
}

// ============================================================
//  XlsxSheet
// ============================================================
void XlsxSheet::put_cell(const Cell& c) {
    data_[c.address.row][c.address.col] = c;
    if (c.address.row + 1 > row_count_) row_count_ = c.address.row + 1;
    if (c.address.col + 1 > col_count_) col_count_ = c.address.col + 1;
}
void XlsxSheet::set_dimensions(uint32_t r, uint32_t c) { row_count_=r; col_count_=c; }

static Cell blank_cell(uint32_t r, uint32_t c) {
    Cell cell; cell.address={r,c}; cell.type=CellType::Blank; cell.value=BlankValue{}; return cell;
}

Cell XlsxSheet::cell(uint32_t r, uint32_t c) const {
    auto ri = data_.find(r);
    if (ri != data_.end()) { auto ci = ri->second.find(c); if (ci != ri->second.end()) return ci->second; }
    return blank_cell(r,c);
}
Cell XlsxSheet::cell(const CellAddress& a) const { return cell(a.row,a.col); }
Cell XlsxSheet::cell(const std::string& a1) const { return cell(CellAddress::from_a1(a1)); }

std::optional<Cell> XlsxSheet::try_cell(uint32_t r, uint32_t c) const {
    auto ri = data_.find(r);
    if (ri == data_.end()) return std::nullopt;
    auto ci = ri->second.find(c);
    if (ci == ri->second.end() || ci->second.is_blank()) return std::nullopt;
    return ci->second;
}
std::vector<Cell> XlsxSheet::row(uint32_t r) const {
    std::vector<Cell> out;
    auto ri = data_.find(r);
    if (ri != data_.end()) for (auto& [c,cell]:ri->second) out.push_back(cell);
    return out;
}
std::vector<Cell> XlsxSheet::col(uint32_t col_idx) const {
    std::vector<Cell> out;
    for (uint32_t r = 0; r < row_count_; ++r) {
        auto oc = try_cell(r, col_idx);
        if (oc) out.push_back(*oc);
    }
    return out;
}
std::vector<Cell> XlsxSheet::cells() const {
    std::vector<Cell> out;
    for (auto& [r,cols]:data_) for (auto& [c,cell]:cols) if (!cell.is_blank()) out.push_back(cell);
    return out;
}
void XlsxSheet::for_each_cell(std::function<void(const Cell&)> fn) const {
    for (auto& [r,cols]:data_) for (auto& [c,cell]:cols) fn(cell);
}

static void fill_type(Cell& c) {
    std::visit([&](auto&& v){
        using T=std::decay_t<decltype(v)>;
        if      constexpr(std::is_same_v<T,BlankValue>)   c.type=CellType::Blank;
        else if constexpr(std::is_same_v<T,bool>)         c.type=CellType::Boolean;
        else if constexpr(std::is_same_v<T,int64_t>)      c.type=CellType::Number;
        else if constexpr(std::is_same_v<T,double>)       c.type=CellType::Number;
        else if constexpr(std::is_same_v<T,std::string>)  c.type=CellType::String;
        else if constexpr(std::is_same_v<T,ErrorValue>)   c.type=CellType::Error;
    }, c.value);
}
void XlsxSheet::set_cell(uint32_t r, uint32_t c, const CellValue& v) {
    Cell cell; cell.address={r,c}; cell.value=v; fill_type(cell); put_cell(cell);
}
void XlsxSheet::set_cell(const std::string& a1, const CellValue& v) {
    auto a=CellAddress::from_a1(a1); set_cell(a.row,a.col,v);
}
void XlsxSheet::set_formula(const std::string& a1, const std::string& f) {
    auto a=CellAddress::from_a1(a1);
    Cell c; c.address=a; c.type=CellType::Formula; c.formula=f; c.value=BlankValue{}; put_cell(c);
}
void XlsxSheet::set_row(uint32_t row_idx, const std::vector<CellValue>& values) {
    for (uint32_t c = 0; c < static_cast<uint32_t>(values.size()); ++c)
        set_cell(row_idx, c, values[c]);
}
void XlsxSheet::merge(const CellRange& range) {
    merges_.push_back(range);
}
void XlsxSheet::unmerge(const CellRange& range) {
    merges_.erase(std::remove_if(merges_.begin(), merges_.end(), [&](const CellRange& r){
        return r.row1==range.row1 && r.col1==range.col1 && r.row2==range.row2 && r.col2==range.col2;
    }), merges_.end());
}
std::vector<CellRange> XlsxSheet::merged_ranges() const {
    return merges_;
}

// ============================================================
//  XlsxWorkbook
// ============================================================
void XlsxWorkbook::add_parsed_sheet(std::unique_ptr<XlsxSheet> s) { sheets_.push_back(std::move(s)); }
Sheet& XlsxWorkbook::sheet(size_t i) {
    if(i>=sheets_.size()) throw RangeError("Sheet index "+std::to_string(i)+" out of range");
    return *sheets_[i];
}
const Sheet& XlsxWorkbook::sheet(size_t i) const {
    if(i>=sheets_.size()) throw RangeError("Sheet index "+std::to_string(i)+" out of range");
    return *sheets_[i];
}
Sheet& XlsxWorkbook::sheet(const std::string& name) {
    for(auto& s:sheets_) if(s->name()==name) return *s;
    throw RangeError("Sheet not found: "+name);
}
std::vector<std::string> XlsxWorkbook::sheet_names() const {
    std::vector<std::string> n; for(auto& s:sheets_) n.push_back(s->name()); return n;
}
Sheet& XlsxWorkbook::add_sheet(const std::string& name) {
    for(auto& s:sheets_) if(s->name()==name) throw WriteError("Sheet exists: "+name);
    sheets_.push_back(std::make_unique<XlsxSheet>(name));
    return *sheets_.back();
}
void XlsxWorkbook::remove_sheet(size_t i) {
    if(i>=sheets_.size()) throw RangeError("Sheet index out of range");
    sheets_.erase(sheets_.begin()+static_cast<ptrdiff_t>(i));
}
void XlsxWorkbook::rename_sheet(size_t i, const std::string& name) {
    if(i>=sheets_.size()) throw RangeError("Sheet index out of range");
    for(size_t j=0;j<sheets_.size();++j)
        if(j!=i && sheets_[j]->name()==name) throw WriteError("Sheet name exists: "+name);
    auto old = std::move(sheets_[i]);
    auto ns  = std::make_unique<XlsxSheet>(name);
    old->for_each_cell([&](const Cell& c){ ns->put_cell(c); });
    ns->set_dimensions(old->row_count(), old->col_count());
    sheets_[i] = std::move(ns);
}
void XlsxWorkbook::save(const std::string& path, const SaveOptions& opts) const {
    FileFormat fmt = opts.format == FileFormat::Auto ? FileFormat::XLSX : opts.format;
    if (fmt == FileFormat::XLS)
        throw WriteError("XLS write is not supported. Open the file and re-save as XLSX using FileFormat::XLSX.");
    auto bytes = to_bytes(fmt, opts);
    std::ofstream f(path, std::ios::binary);
    if(!f) throw IOError("Cannot write: "+path);
    f.write(reinterpret_cast<const char*>(bytes.data()),
            static_cast<std::streamsize>(bytes.size()));
}
std::vector<uint8_t> XlsxWorkbook::to_bytes(FileFormat fmt, const SaveOptions&) const {
    if (fmt == FileFormat::XLS)
        throw WriteError("XLS write is not supported. Open the file and re-save as XLSX using FileFormat::XLSX.");
    XlsxWriter w; return w.write(*this, {});
}

// ============================================================
//  XlsxParser
// ============================================================
CellAddress XlsxParser::parse_ref(const std::string& ref) {
    return CellAddress::from_a1(ref);
}

static bool is_known_entry(const std::string& path) {
    if (path == "[Content_Types].xml") return true;
    if (path == "_rels/.rels") return true;
    if (path == "xl/workbook.xml") return true;
    if (path == "xl/_rels/workbook.xml.rels") return true;
    if (path == "xl/sharedStrings.xml") return true;
    if (path == "xl/styles.xml") return true;
    if (path.size() > 17 && path.substr(0,17) == "xl/worksheets/she" &&
        path.find(".xml") != std::string::npos) return true;
    if (path.size() > 23 && path.substr(0,23) == "xl/worksheets/_rels/she" &&
        path.find(".xml.rels") != std::string::npos) return true;
    return false;
}

static std::string content_type_for(const std::string& path) {
    auto dot = path.rfind('.');
    if (dot == std::string::npos) return "application/octet-stream";
    std::string ext = path.substr(dot+1);
    if (ext=="png")  return "image/png";
    if (ext=="jpg" || ext=="jpeg") return "image/jpeg";
    if (ext=="gif")  return "image/gif";
    if (ext=="emf")  return "image/x-emf";
    if (ext=="wmf")  return "image/x-wmf";
    if (ext=="xml")  return "application/xml";
    if (ext=="rels") return "application/vnd.openxmlformats-package.relationships+xml";
    if (ext=="vml")  return "application/vnd.openxmlformats-officedocument.vmlDrawing";
    return "application/octet-stream";
}

std::unique_ptr<XlsxWorkbook> XlsxParser::parse(const std::vector<uint8_t>& data) {
    ZipReader zip(data);

    XlsxSharedStrings sst;
    if (zip.has("xl/sharedStrings.xml")) sst.parse(zip.text("xl/sharedStrings.xml"));

    XlsxStyles styles;
    std::string styles_xml;
    if (zip.has("xl/styles.xml")) {
        styles_xml = zip.text("xl/styles.xml");
        styles.parse(styles_xml);
    }

    if (!zip.has("xl/workbook.xml")) throw ParseError("XLSX: missing xl/workbook.xml");

    std::vector<Relationship> rels;
    if (zip.has("xl/_rels/workbook.xml.rels"))
        rels = parse_rels(zip.text("xl/_rels/workbook.xml.rels"));

    auto metas = parse_workbook_xml(zip.text("xl/workbook.xml"), rels);
    auto wb = std::make_unique<XlsxWorkbook>();
    wb->original_styles_xml_ = styles_xml;

    // Collect sheet paths to mark as known
    std::vector<std::string> sheet_paths;
    for (auto& meta : metas) {
        std::string path;
        if (!meta.path.empty() && meta.path[0] == '/')
            path = "xl" + meta.path;
        else if (meta.path.find('/') != std::string::npos)
            path = "xl/" + meta.path;
        else
            path = "xl/worksheets/" + meta.path;
        sheet_paths.push_back(path);
    }

    // Collect passthrough entries
    bool has_unknown = false;
    for (auto& p : zip.paths()) {
        bool known = is_known_entry(p);
        if (!known) {
            for (auto& sp : sheet_paths) if (p == sp) { known = true; break; }
        }
        if (!known) {
            wb->passthrough_entries_[p] = zip.get(p).data;
            has_unknown = true;
        }
    }
    if (has_unknown) {
        warn(ParseWarning::Kind::UnknownXmlElement, "[Content_Types].xml",
             "This file contains elements excellib does not parse (they will be preserved on save).");
    }

    for (size_t si = 0; si < metas.size(); ++si) {
        auto& meta = metas[si];
        std::string path = sheet_paths[si];

        if (!zip.has(path)) throw ParseError("XLSX: sheet not found in ZIP: " + path);
        wb->add_parsed_sheet(parse_sheet_xml(meta.name, zip.text(path), sst, styles));
    }
    return wb;
}

std::vector<XlsxParser::SheetMeta> XlsxParser::parse_workbook_xml(
    const std::string& xml, const std::vector<Relationship>& rels)
{
    std::vector<SheetMeta> out;
    size_t pos = 0;
    while (true) {
        auto sh = xml.find("<sheet ", pos);
        if (sh == std::string::npos) break;
        auto she = xml.find('>', sh);
        if (she == std::string::npos) break;
        std::string el = xml.substr(sh, she-sh+1);
        SheetMeta m;
        m.name = xml_unescape(xml_attr(el, "name"));
        m.rid  = xml_attr(el, "r:id");
        for (auto& r : rels) if (r.id == m.rid) { m.path = r.target; break; }
        if (!m.path.empty()) out.push_back(m);
        pos = she+1;
    }
    return out;
}

std::unique_ptr<XlsxSheet> XlsxParser::parse_sheet_xml(
    const std::string& name, const std::string& xml,
    const XlsxSharedStrings& sst, const XlsxStyles& styles)
{
    auto sheet = std::make_unique<XlsxSheet>(name);
    size_t pos = 0;
    while (true) {
        auto rs = xml.find("<row ", pos);
        if (rs == std::string::npos) break;
        auto re = xml.find("</row>", rs);
        if (re == std::string::npos) break;
        std::string row_blk = xml.substr(rs, re-rs);

        size_t cp = 0;
        while (true) {
            auto cs = row_blk.find("<c ", cp);
            if (cs == std::string::npos) break;
            auto tag_end = row_blk.find('>', cs);
            if (tag_end == std::string::npos) break;
            bool self_close = tag_end > 0 && row_blk[tag_end-1] == '/';

            std::string cell_xml;
            size_t next_cp;
            if (self_close) {
                cell_xml = row_blk.substr(cs, tag_end-cs+1);
                next_cp  = tag_end+1;
            } else {
                auto ce = row_blk.find("</c>", tag_end);
                if (ce == std::string::npos) { cp = tag_end+1; break; }
                cell_xml = row_blk.substr(cs, ce-cs+4);
                next_cp  = ce+4;
            }

            std::string tag = cell_xml.substr(0, cell_xml.find('>')+1);
            std::string ref  = xml_attr(tag, "r");
            std::string type = xml_attr(tag, "t");
            std::string s    = xml_attr(tag, "s");
            std::string v    = self_close ? "" : xml_text(cell_xml, "v");
            std::string f    = self_close ? "" : xml_text(cell_xml, "f");

            if (!ref.empty()) {
                try {
                    Cell c = parse_cell_element(ref, type, s, v, f, sst, styles);
                    sheet->put_cell(c);
                } catch (const FormatError& e) {
                    if (opts_.strict_validation) throw;
                    warn(ParseWarning::Kind::MalformedField,
                         "xl/worksheets/ row cell ref=" + ref, e.what());
                } catch (const std::exception& e) {
                    warn(ParseWarning::Kind::DataDropped,
                         "xl/worksheets/ row cell ref=" + ref, e.what());
                }
            }
            cp = next_cp;
        }
        pos = re + 6;
    }

    // Parse mergeCells
    auto mc_start = xml.find("<mergeCells");
    if (mc_start != std::string::npos) {
        auto mc_end = xml.find("</mergeCells>", mc_start);
        if (mc_end == std::string::npos) mc_end = xml.size();
        std::string blk = xml.substr(mc_start, mc_end - mc_start);
        size_t p = 0;
        while (true) {
            auto mp = blk.find("<mergeCell ", p);
            if (mp == std::string::npos) break;
            auto mpe = blk.find('>', mp);
            if (mpe == std::string::npos) break;
            std::string el = blk.substr(mp, mpe-mp+1);
            std::string ref_str = xml_attr(el, "ref");
            if (!ref_str.empty()) {
                try {
                    sheet->merges_.push_back(CellRange::from_a1(ref_str));
                } catch (const std::exception& e) {
                    warn(ParseWarning::Kind::MalformedField,
                         "xl/worksheets/ mergeCells ref=" + ref_str, e.what());
                }
            }
            p = mpe + 1;
        }
    }

    return sheet;
}

Cell XlsxParser::parse_cell_element(
    const std::string& ref, const std::string& type,
    const std::string& s_attr, const std::string& v_text,
    const std::string& f_text,
    const XlsxSharedStrings& sst, const XlsxStyles& styles)
{
    Cell c;
    c.address = parse_ref(ref);

    if (!s_attr.empty()) {
        try {
            uint32_t xf = static_cast<uint32_t>(std::stoul(s_attr));
            CellStyle cs; cs.format = styles.cell_format(xf); c.style = cs;
        } catch (...) {}
    }
    if (!f_text.empty()) c.formula = xml_unescape(f_text);

    if (type == "s") {
        c.type = CellType::String;
        // FIX: guard against empty/non-numeric v_text
        if (!v_text.empty()) {
            try { c.value = sst.get(static_cast<uint32_t>(std::stoul(v_text))); }
            catch (...) { c.value = std::string(""); }
        }
    } else if (type == "str" || type == "inlineStr") {
        c.type = CellType::String;
        c.value = xml_unescape(v_text);
    } else if (type == "b") {
        c.type = CellType::Boolean;
        c.value = (v_text == "1" || v_text == "true");
    } else if (type == "e") {
        c.type = CellType::Error;
        c.value = ErrorValue{xml_unescape(v_text)};
    } else {
        if (v_text.empty()) {
            c.type = CellType::Blank; c.value = BlankValue{};
        } else {
            c.type = CellType::Number;
            try {
                // Prefer integer if no decimal/exponent
                if (v_text.find('.') == std::string::npos &&
                    v_text.find('e') == std::string::npos &&
                    v_text.find('E') == std::string::npos) {
                    c.value = static_cast<int64_t>(std::stoll(v_text));
                } else {
                    c.value = std::stod(v_text);
                }
            } catch (...) {
                c.type = CellType::String; c.value = v_text;
            }
        }
    }
    if (!f_text.empty() && c.type != CellType::Error) c.type = CellType::Formula;
    return c;
}

// ============================================================
//  XlsxWriter — minimal ZIP builder
// ============================================================
static void ap16(std::vector<uint8_t>& v, uint16_t x) { v.push_back(x&0xFF); v.push_back((x>>8)&0xFF); }
static void ap32(std::vector<uint8_t>& v, uint32_t x) {
    v.push_back(x&0xFF); v.push_back((x>>8)&0xFF);
    v.push_back((x>>16)&0xFF); v.push_back((x>>24)&0xFF);
}

struct ZFile { std::string name; std::vector<uint8_t> data; uint32_t off{0}, crc{0}; };

static std::vector<uint8_t> build_zip(std::vector<ZFile>& files) {
    std::vector<uint8_t> out;
    for (auto& f : files) {
        f.off = uint32_t(out.size());
        f.crc = excellib::detail::crc32_compute(f.data.data(), f.data.size());
        ap32(out,0x04034B50); ap16(out,20); ap16(out,0); ap16(out,0);
        ap16(out,0); ap16(out,0);
        ap32(out,f.crc); ap32(out,uint32_t(f.data.size())); ap32(out,uint32_t(f.data.size()));
        ap16(out,uint16_t(f.name.size())); ap16(out,0);
        out.insert(out.end(),f.name.begin(),f.name.end());
        out.insert(out.end(),f.data.begin(),f.data.end());
    }
    uint32_t cd_off = uint32_t(out.size());
    for (auto& f : files) {
        ap32(out,0x02014B50); ap16(out,20); ap16(out,20); ap16(out,0);
        ap16(out,0); ap16(out,0); ap16(out,0);
        ap32(out,f.crc); ap32(out,uint32_t(f.data.size())); ap32(out,uint32_t(f.data.size()));
        ap16(out,uint16_t(f.name.size())); ap16(out,0); ap16(out,0);
        ap16(out,0); ap16(out,0); ap32(out,0); ap32(out,f.off);
        out.insert(out.end(),f.name.begin(),f.name.end());
    }
    uint32_t cd_sz = uint32_t(out.size()) - cd_off;
    ap32(out,0x06054B50); ap16(out,0); ap16(out,0);
    ap16(out,uint16_t(files.size())); ap16(out,uint16_t(files.size()));
    ap32(out,cd_sz); ap32(out,cd_off); ap16(out,0);
    return out;
}

static std::string esc(const std::string& s) {
    std::string r;
    for (char c:s) {
        switch(c){
            case '&': r+="&amp;"; break; case '<': r+="&lt;"; break;
            case '>': r+="&gt;"; break;  case '"': r+="&quot;"; break;
            default:  r+=c;
        }
    }
    return r;
}

static std::vector<uint8_t> sb(const std::string& s) {
    return {s.begin(), s.end()};
}

std::string XlsxWriter::sheet_xml(const XlsxSheet& sh, XlsxSharedStrings& sst) {
    std::ostringstream o;
    o << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
      << "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
      << "<sheetData>";
    for (uint32_t r = 0; r < sh.row_count(); ++r) {
        auto cells = sh.row(r);
        if (cells.empty()) continue;
        o << "<row r=\"" << r+1 << "\">";
        for (auto& c : cells) {
            std::string ref = CellAddress{c.address.row,c.address.col}.to_a1();
            if (c.has_formula()) {
                o << "<c r=\"" << ref << "\"><f>" << esc(*c.formula) << "</f></c>";
                continue;
            }
            std::visit([&](auto&& v){
                using T=std::decay_t<decltype(v)>;
                if constexpr(std::is_same_v<T,BlankValue>) {}
                else if constexpr(std::is_same_v<T,bool>)
                    o << "<c r=\""<<ref<<"\" t=\"b\"><v>"<<(v?1:0)<<"</v></c>";
                else if constexpr(std::is_same_v<T,int64_t>)
                    o << "<c r=\""<<ref<<"\"><v>"<<v<<"</v></c>";
                else if constexpr(std::is_same_v<T,double>)
                    o << "<c r=\""<<ref<<"\"><v>"<<v<<"</v></c>";
                else if constexpr(std::is_same_v<T,std::string>)
                    o << "<c r=\""<<ref<<"\" t=\"s\"><v>"<<sst.intern(v)<<"</v></c>";
                else if constexpr(std::is_same_v<T,ErrorValue>)
                    o << "<c r=\""<<ref<<"\" t=\"e\"><v>"<<esc(v.code)<<"</v></c>";
            }, c.value);
        }
        o << "</row>";
    }
    o << "</sheetData>";
    auto& merges = sh.merges_;
    if (!merges.empty()) {
        o << "<mergeCells count=\"" << merges.size() << "\">";
        for (auto& m : merges)
            o << "<mergeCell ref=\"" << esc(m.to_a1()) << "\"/>";
        o << "</mergeCells>";
    }
    o << "</worksheet>";
    return o.str();
}

std::vector<uint8_t> XlsxWriter::write(const XlsxWorkbook& wb, const SaveOptions&) {
    XlsxSharedStrings sst;
    std::vector<ZFile> files;
    auto names = wb.sheet_names();

    // Track sheet paths for passthrough filtering
    std::vector<std::string> sheet_paths;
    for (size_t i = 0; i < wb.sheet_count(); ++i) {
        const auto& sh = dynamic_cast<const XlsxSheet&>(wb.sheet(i));
        ZFile f;
        std::string sp = "xl/worksheets/sheet" + std::to_string(i+1) + ".xml";
        sheet_paths.push_back(sp);
        f.name = sp;
        f.data = sb(sheet_xml(sh, sst));
        files.push_back(std::move(f));
    }

    bool has_ss = sst.size() > 0;
    if (has_ss) {
        std::ostringstream o;
        o << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
          << "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
          << " count=\""<<sst.size()<<"\" uniqueCount=\""<<sst.size()<<"\">";
        for (size_t i=0;i<sst.size();++i)
            o<<"<si><t xml:space=\"preserve\">"<<esc(sst.get(i))<<"</t></si>";
        o << "</sst>";
        files.push_back({"xl/sharedStrings.xml", sb(o.str())});
    }

    // styles.xml: prefer original if available
    if (!wb.original_styles_xml_.empty()) {
        files.push_back({"xl/styles.xml", sb(wb.original_styles_xml_)});
    } else {
        files.push_back({"xl/styles.xml", sb(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
            "<fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts>"
            "<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill>"
            "<fill><patternFill patternType=\"gray125\"/></fill></fills>"
            "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>"
            "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\"/></cellStyleXfs>"
            "<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" xfId=\"0\"/></cellXfs>"
            "</styleSheet>"
        )});
    }

    // workbook.xml
    {
        std::ostringstream o;
        o << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
          << "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
          << " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
          << "<sheets>";
        for (size_t i=0;i<names.size();++i)
            o<<"<sheet name=\""<<esc(names[i])<<"\" sheetId=\""<<i+1
             <<"\" r:id=\"rId"<<i+1<<"\"/>";
        o << "</sheets></workbook>";
        files.push_back({"xl/workbook.xml", sb(o.str())});
    }

    // rels
    {
        std::ostringstream o;
        o << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
          << "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
          << "<Relationship Id=\"rId0\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>";
        for (size_t i=0;i<wb.sheet_count();++i)
            o<<"<Relationship Id=\"rId"<<i+1<<"\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet"<<i+1<<".xml\"/>";
        if (has_ss)
            o<<"<Relationship Id=\"rIdSS\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>";
        o << "</Relationships>";
        files.push_back({"xl/_rels/workbook.xml.rels", sb(o.str())});
    }

    // Passthrough entries (non-sheet xml)
    for (auto& [path, data] : wb.passthrough_entries_) {
        // Skip entries that we've already generated
        bool skip = false;
        for (auto& sp : sheet_paths) if (path == sp) { skip = true; break; }
        if (skip) continue;
        files.push_back({path, data});
    }

    // content types — generate ours and append content types for passthrough entries
    {
        std::ostringstream o;
        o << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
          << "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
          << "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
          << "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
          << "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
          << "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>";
        for (size_t i=0;i<wb.sheet_count();++i)
            o<<"<Override PartName=\"/xl/worksheets/sheet"<<i+1<<".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>";
        if (has_ss)
            o<<"<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>";
        // Append content types for passthrough entries
        for (auto& [path, data] : wb.passthrough_entries_) {
            if (path.empty() || path[0] == '_') continue; // skip rels
            std::string dot = path.rfind('.') != std::string::npos ? path.substr(path.rfind('.')+1) : "";
            if (dot == "rels" || dot == "xml") continue; // covered by Default
            std::string ct = content_type_for(path);
            o << "<Override PartName=\"/" << path << "\" ContentType=\"" << ct << "\"/>";
        }
        o << "</Types>";
        files.push_back({"[Content_Types].xml", sb(o.str())});
    }

    files.push_back({"_rels/.rels", sb(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
        "</Relationships>"
    )});

    return build_zip(files);
}

} // namespace excellib::xlsx
