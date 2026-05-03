#include "excellib/xls_parser.hpp"
#include <cstring>
#include <cmath>
#include <cstdio>
#include <algorithm>
#include <sstream>
#include <fstream>

namespace excellib::xls {

static void chk(const std::vector<uint8_t>& d, size_t o, size_t n, const char* w) {
    if (o + n > d.size())
        throw ParseError(std::string("BIFF truncated reading ") + w);
}
uint8_t  BiffRecord::u8 (size_t o) const { chk(data,o,1,"u8");  return data[o]; }
uint16_t BiffRecord::u16(size_t o) const {
    chk(data,o,2,"u16");
    return uint16_t(data[o]) | uint16_t(uint16_t(data[o+1])<<8);
}
uint32_t BiffRecord::u32(size_t o) const {
    chk(data,o,4,"u32");
    return uint32_t(data[o])|uint32_t(data[o+1])<<8|uint32_t(data[o+2])<<16|uint32_t(data[o+3])<<24;
}
int16_t  BiffRecord::i16(size_t o) const { return static_cast<int16_t>(u16(o)); }
double   BiffRecord::f64(size_t o) const {
    chk(data,o,8,"f64"); double v; std::memcpy(&v,data.data()+o,8); return v;
}

double decode_rk(uint32_t rk) {
    double v;
    if (rk & 0x02) {
        v = static_cast<double>(static_cast<int32_t>(rk) >> 2);
    } else {
        uint64_t tmp = static_cast<uint64_t>(rk & 0xFFFFFFFC) << 32;
        std::memcpy(&v, &tmp, 8);
    }
    if (rk & 0x01) v /= 100.0;
    return v;
}

std::string decode_biff8_string(const uint8_t* data, size_t avail,
                                 size_t& consumed, bool short_hdr) {
    consumed = 0;
    size_t hdr = short_hdr ? 1u : 2u;
    if (avail < hdr + 1) throw ParseError("BIFF8 string too short");
    uint16_t nc = short_hdr ? uint16_t(data[0])
                            : uint16_t(data[0]) | uint16_t(uint16_t(data[1])<<8);
    consumed += hdr;
    uint8_t flags = data[consumed++];
    bool compressed = !(flags & 0x01);
    bool has_rich   = (flags & 0x08) != 0;
    bool has_ext    = (flags & 0x04) != 0;
    uint16_t rt = 0; uint32_t ext = 0;
    if (has_rich) {
        if (consumed+2>avail) throw ParseError("BIFF8 rich-text truncated");
        rt = uint16_t(data[consumed])|uint16_t(uint16_t(data[consumed+1])<<8); consumed+=2;
    }
    if (has_ext) {
        if (consumed+4>avail) throw ParseError("BIFF8 ext-data truncated");
        ext = uint32_t(data[consumed])|uint32_t(data[consumed+1])<<8
             |uint32_t(data[consumed+2])<<16|uint32_t(data[consumed+3])<<24; consumed+=4;
    }
    size_t cb = compressed ? nc : size_t(nc)*2;
    if (consumed+cb>avail) throw ParseError("BIFF8 char data truncated");
    std::string r; r.reserve(nc);
    if (compressed) {
        for (uint16_t i=0;i<nc;++i) {
            uint8_t ch=data[consumed+i];
            if (ch<0x80) r+=char(ch);
            else { r+=char(0xC0|(ch>>6)); r+=char(0x80|(ch&0x3F)); }
        }
    } else {
        for (uint16_t i=0;i<nc;++i) {
            uint16_t cp=uint16_t(data[consumed+size_t(i)*2])|uint16_t(uint16_t(data[consumed+size_t(i)*2+1])<<8);
            if (cp<0x80) r+=char(cp);
            else if (cp<0x800) { r+=char(0xC0|(cp>>6)); r+=char(0x80|(cp&0x3F)); }
            else { r+=char(0xE0|(cp>>12)); r+=char(0x80|((cp>>6)&0x3F)); r+=char(0x80|(cp&0x3F)); }
        }
    }
    consumed += cb + size_t(rt)*4 + ext;
    return r;
}

void SharedStringTable::parse(const BiffRecord& sst, const std::vector<BiffRecord>& conts) {
    if (sst.data.size()<8) throw ParseError("SST too short");
    uint32_t total = sst.u32(4); strings_.reserve(total);
    std::vector<uint8_t> buf(sst.data);
    for (auto& c:conts) buf.insert(buf.end(),c.data.begin(),c.data.end());
    const uint8_t* p=buf.data()+8; size_t left=buf.size()-8;
    for (uint32_t i=0;i<total;++i) {
        if (!left) throw ParseError("SST exhausted");
        size_t consumed=0;
        strings_.push_back(decode_biff8_string(p,left,consumed));
        if (consumed>left) throw ParseError("SST overrun");
        p+=consumed; left-=consumed;
    }
}
const std::string& SharedStringTable::get(uint32_t i) const {
    if (i>=strings_.size()) throw RangeError("SST index out of range");
    return strings_[i];
}

FormatRegistry::FormatRegistry() {
    struct E { uint16_t id; const char* fmt; };
    static const E B[]={
        {0,"General"},{1,"0"},{2,"0.00"},{3,"#,##0"},{4,"#,##0.00"},
        {9,"0%"},{10,"0.00%"},{11,"0.00E+00"},{12,"# ?/?"},
        {13,"# ??/?"},{14,"M/D/YYYY"},{15,"D-MMM-YY"},{16,"D-MMM"},{17,"MMM-YY"},
        {18,"h:mm AM/PM"},{19,"h:mm:ss AM/PM"},{20,"h:mm"},{21,"h:mm:ss"},
        {22,"M/D/YYYY h:mm"},{45,"mm:ss"},{46,"[h]:mm:ss"},{47,"mmss.0"},{49,"@"},
    };
    for (auto& b:B) formats_[b.id]=b.fmt;
}
void FormatRegistry::add(uint16_t i, const std::string& s) { formats_[i]=s; }
FormatInfo FormatRegistry::get(uint16_t i) const {
    FormatInfo fi; fi.format_index=i;
    auto it=formats_.find(i);
    if (it!=formats_.end()) { fi.format_string=it->second; fi.is_date_format=is_date_format(i); }
    return fi;
}
bool FormatRegistry::is_date_format(uint16_t i) const {
    if ((i>=14&&i<=22)||(i>=45&&i<=47)) return true;
    auto it=formats_.find(i);
    return it!=formats_.end()&&looks_like_date(it->second);
}
bool FormatRegistry::looks_like_date(const std::string& f) {
    std::string lo=f; std::transform(lo.begin(),lo.end(),lo.begin(),::tolower);
    return lo.find('y')!=std::string::npos||
           (lo.find('d')!=std::string::npos)||
           (lo.find('m')!=std::string::npos&&lo.find('h')==std::string::npos);
}

void XFTable::add(const XFRecord& x) { records_.push_back(x); }
const XFRecord& XFTable::get(uint16_t i) const {
    if (i>=records_.size()) throw RangeError("XF index out of range");
    return records_[i];
}

void XlsSheet::fill_type(Cell& c) {
    std::visit([&](auto&& v){
        using T=std::decay_t<decltype(v)>;
        if      constexpr(std::is_same_v<T,BlankValue>)   c.type=CellType::Blank;
        else if constexpr(std::is_same_v<T,bool>)         c.type=CellType::Boolean;
        else if constexpr(std::is_same_v<T,int64_t>)      c.type=CellType::Number;
        else if constexpr(std::is_same_v<T,double>)       c.type=CellType::Number;
        else if constexpr(std::is_same_v<T,std::string>)  c.type=CellType::String;
        else if constexpr(std::is_same_v<T,ErrorValue>)   c.type=CellType::Error;
    },c.value);
}
void XlsSheet::put_cell(const Cell& c) {
    data_[c.address.row][c.address.col]=c;
    if (c.address.row+1>row_count_) row_count_=c.address.row+1;
    if (c.address.col+1>col_count_) col_count_=c.address.col+1;
}
void XlsSheet::set_dimensions(uint32_t r,uint32_t c){row_count_=r;col_count_=c;}
static Cell mk_blank(uint32_t r,uint32_t c){Cell x;x.address={r,c};x.type=CellType::Blank;x.value=BlankValue{};return x;}
Cell XlsSheet::cell(uint32_t r,uint32_t c) const {
    auto ri=data_.find(r);
    if (ri!=data_.end()){auto ci=ri->second.find(c);if(ci!=ri->second.end())return ci->second;}
    return mk_blank(r,c);
}
Cell XlsSheet::cell(const CellAddress& a) const {return cell(a.row,a.col);}
Cell XlsSheet::cell(const std::string& a) const {return cell(CellAddress::from_a1(a));}
std::optional<Cell> XlsSheet::try_cell(uint32_t r,uint32_t c) const {
    auto ri=data_.find(r); if(ri==data_.end()) return std::nullopt;
    auto ci=ri->second.find(c); if(ci==ri->second.end()||ci->second.is_blank()) return std::nullopt;
    return ci->second;
}
std::vector<Cell> XlsSheet::row(uint32_t r) const {
    std::vector<Cell> out;
    auto ri=data_.find(r); if(ri!=data_.end()) for(auto&[c,x]:ri->second) out.push_back(x);
    return out;
}
std::vector<Cell> XlsSheet::col(uint32_t col_idx) const {
    std::vector<Cell> out;
    for (uint32_t r = 0; r < row_count_; ++r) {
        auto oc = try_cell(r, col_idx);
        if (oc) out.push_back(*oc);
    }
    return out;
}
std::vector<Cell> XlsSheet::cells() const {
    std::vector<Cell> out;
    for(auto&[r,cols]:data_) for(auto&[c,x]:cols) if(!x.is_blank()) out.push_back(x);
    return out;
}
void XlsSheet::for_each_cell(std::function<void(const Cell&)> fn) const {
    for(auto&[r,cols]:data_) for(auto&[c,x]:cols) fn(x);
}
void XlsSheet::set_cell(uint32_t r,uint32_t c,const CellValue& v){Cell x;x.address={r,c};x.value=v;fill_type(x);put_cell(x);}
void XlsSheet::set_cell(const std::string& a,const CellValue& v){auto addr=CellAddress::from_a1(a);set_cell(addr.row,addr.col,v);}
void XlsSheet::set_formula(const std::string& a,const std::string& f){
    auto addr=CellAddress::from_a1(a);Cell c;c.address=addr;c.type=CellType::Formula;c.formula=f;c.value=BlankValue{};put_cell(c);
}
void XlsSheet::set_row(uint32_t row_idx, const std::vector<CellValue>& values) {
    for (uint32_t c = 0; c < static_cast<uint32_t>(values.size()); ++c)
        set_cell(row_idx, c, values[c]);
}
void XlsSheet::merge(const CellRange& range) {
    throw WriteError("XLS is read-only");
}
void XlsSheet::unmerge(const CellRange& range) {
    throw WriteError("XLS is read-only");
}
std::vector<CellRange> XlsSheet::merged_ranges() const {
    return merges_;
}

void XlsWorkbook::add_parsed_sheet(std::unique_ptr<XlsSheet> s){sheets_.push_back(std::move(s));}
Sheet& XlsWorkbook::sheet(size_t i){if(i>=sheets_.size())throw RangeError("Sheet index OOB");return *sheets_[i];}
const Sheet& XlsWorkbook::sheet(size_t i) const {if(i>=sheets_.size())throw RangeError("Sheet index OOB");return *sheets_[i];}
Sheet& XlsWorkbook::sheet(const std::string& n){for(auto&s:sheets_) if(s->name()==n) return *s;throw RangeError("Sheet not found: "+n);}
std::vector<std::string> XlsWorkbook::sheet_names() const {std::vector<std::string> r;for(auto&s:sheets_)r.push_back(s->name());return r;}
Sheet& XlsWorkbook::add_sheet(const std::string& n){for(auto&s:sheets_)if(s->name()==n)throw WriteError("Dup sheet: "+n);sheets_.push_back(std::make_unique<XlsSheet>(n));return*sheets_.back();}
void XlsWorkbook::remove_sheet(size_t i){if(i>=sheets_.size())throw RangeError("OOB");sheets_.erase(sheets_.begin()+static_cast<ptrdiff_t>(i));}
void XlsWorkbook::rename_sheet(size_t i,const std::string& n){
    if(i>=sheets_.size())throw RangeError("OOB");
    for(size_t j=0;j<sheets_.size();++j)if(j!=i&&sheets_[j]->name()==n)throw WriteError("Name exists");
    auto old=std::move(sheets_[i]);auto ns=std::make_unique<XlsSheet>(n);
    old->for_each_cell([&](const Cell& c){ns->put_cell(c);});
    ns->set_dimensions(old->row_count(),old->col_count());sheets_[i]=std::move(ns);
}
void XlsWorkbook::save(const std::string&,const SaveOptions&) const {
    throw WriteError("XLS write is not supported. Open the file and re-save as XLSX using FileFormat::XLSX.");
}
std::vector<uint8_t> XlsWorkbook::to_bytes(FileFormat,const SaveOptions&) const {
    throw WriteError("XLS write is not supported. Open the file and re-save as XLSX using FileFormat::XLSX.");
}

static constexpr uint8_t OLE2_MAGIC[8]={0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1};

std::vector<uint8_t> XlsParser::extract_workbook_stream(const std::vector<uint8_t>& raw) {
    if (raw.size()<512) throw ParseError("File too small for OLE2");
    if (std::memcmp(raw.data(),OLE2_MAGIC,8)!=0) throw ParseError("Not OLE2 (bad magic)");

    auto r16=[&](size_t o)->uint16_t{return uint16_t(raw[o])|uint16_t(uint16_t(raw[o+1])<<8);};
    auto r32=[&](size_t o)->uint32_t{return uint32_t(raw[o])|uint32_t(raw[o+1])<<8|uint32_t(raw[o+2])<<16|uint32_t(raw[o+3])<<24;};

    uint16_t ss_pow = r16(30), ms_pow = r16(32);
    uint32_t num_fat= r32(44), dir_s  = r32(48);
    uint32_t mc     = r32(56), mfs    = r32(60);
    size_t   ss     = size_t(1)<<ss_pow;
    size_t   ms     = size_t(1)<<ms_pow;
    auto     so     = [&](uint32_t sid){return 512+size_t(sid)*ss;};

    std::vector<uint32_t> fat;
    for (uint32_t i=0;i<std::min(num_fat,uint32_t(109));++i) {
        uint32_t fs=r32(76+i*4); if(fs>=0xFFFFFFFE) break;
        size_t off=so(fs); if(off+ss>raw.size()) throw ParseError("FAT OOB");
        for (size_t j=0;j<ss/4;++j){uint32_t e;std::memcpy(&e,raw.data()+off+j*4,4);fat.push_back(e);}
    }

    auto read_chain=[&](uint32_t start)->std::vector<uint8_t>{
        std::vector<uint8_t> out; uint32_t sid=start;
        while(sid<0xFFFFFFFE){
            if(sid>=fat.size()) throw ParseError("FAT chain broken");
            size_t off=so(sid); if(off+ss>raw.size()) throw ParseError("Sector OOB");
            out.insert(out.end(),raw.data()+off,raw.data()+off+ss); sid=fat[sid];
        }
        return out;
    };

    auto dir=read_chain(dir_s);
    constexpr size_t DE=128;
    if(dir.size()<DE) throw ParseError("Dir too small");

    uint32_t root_start=r32(dir.data()+116-dir.data()); /* offset 116 in first dir entry */
    {/* use memcpy for portability */
        std::memcpy(&root_start,dir.data()+116,4);
    }
    uint64_t root_sz=0; std::memcpy(&root_sz,dir.data()+120,8);

    std::vector<uint8_t> mini_stream;
    std::vector<uint32_t> minifat;
    if (root_start<0xFFFFFFFE) {
        mini_stream=read_chain(root_start); mini_stream.resize(static_cast<size_t>(root_sz));
        uint32_t mfs_sid=mfs;
        while(mfs_sid<0xFFFFFFFE){
            if(mfs_sid>=fat.size()) break;
            size_t off=so(mfs_sid); if(off+ss>raw.size()) break;
            for(size_t j=0;j<ss/4;++j){uint32_t e;std::memcpy(&e,raw.data()+off+j*4,4);minifat.push_back(e);}
            mfs_sid=fat[mfs_sid];
        }
    }
    auto read_mini=[&](uint32_t start,uint64_t sz)->std::vector<uint8_t>{
        std::vector<uint8_t> out; uint32_t sid=start;
        while(sid<0xFFFFFFFE){
            if(sid>=minifat.size()) throw ParseError("MiniFAT broken");
            size_t off=size_t(sid)*ms; if(off+ms>mini_stream.size()) throw ParseError("Mini OOB");
            out.insert(out.end(),mini_stream.data()+off,mini_stream.data()+off+ms); sid=minifat[sid];
        }
        out.resize(static_cast<size_t>(sz)); return out;
    };

    size_t ne=dir.size()/DE;
    for(size_t i=1;i<ne;++i){
        const uint8_t* e=dir.data()+i*DE;
        if(e[66]!=2) continue;
        uint16_t nl=uint16_t(e[64])|uint16_t(uint16_t(e[65])<<8);
        std::string nm; size_t chars=nl>1?(nl-2)/2:0;
        for(size_t j=0;j<chars;++j){uint16_t ch;std::memcpy(&ch,e+j*2,2);if(ch<128)nm+=char(ch);}
        if(nm=="Workbook"||nm=="Book"){
            uint32_t s=0; std::memcpy(&s,e+116,4);
            uint64_t z=0; std::memcpy(&z,e+120,8);
            if(z<mc&&!mini_stream.empty()) return read_mini(s,z);
            auto d=read_chain(s); d.resize(static_cast<size_t>(z)); return d;
        }
    }
    throw ParseError("OLE2: no Workbook stream found");
}

std::vector<BiffRecord> XlsParser::read_records(const std::vector<uint8_t>& stream) {
    std::vector<BiffRecord> recs; size_t pos=0;
    while(pos+4<=stream.size()){
        uint16_t rt,len;
        std::memcpy(&rt, stream.data()+pos,  2);
        std::memcpy(&len,stream.data()+pos+2,2);
        pos+=4;
        if(pos+len>stream.size()) throw ParseError("BIFF record overruns stream");
        BiffRecord r; r.raw_type=rt; r.type=static_cast<RecordType>(rt);
        r.data.assign(stream.data()+pos,stream.data()+pos+len); pos+=len;
        recs.push_back(std::move(r));
    }
    return recs;
}

static void require_size(const BiffRecord& r, size_t n, const char* where) {
    if (r.data.size() < n)
        throw ParseError(std::string(where) + ": record too short (" +
                         std::to_string(r.data.size()) + " < " + std::to_string(n) + ")");
}

std::unique_ptr<XlsWorkbook> XlsParser::parse(const std::vector<uint8_t>& raw) {
    auto stream=extract_workbook_stream(raw);
    auto recs  =read_records(stream);
    GlobalContext ctx; bool sst_seen=false;

    for(size_t i=0;i<recs.size();++i){
        auto& r=recs[i];
        if(r.type==RecordType::FORMAT&&r.data.size()>=3){
            uint16_t id=r.u16(0); size_t consumed=0;
            ctx.formats.add(id,decode_biff8_string(r.data.data()+2,r.data.size()-2,consumed));
        } else if(r.type==RecordType::XF&&r.data.size()>=4){
            XFRecord xf; xf.font_index=r.u16(0); xf.format_index=r.u16(2);
            if(r.data.size()>4) xf.is_style_xf=(r.data[4]&0x04)!=0;
            ctx.xf_table.add(xf);
        } else if(r.type==RecordType::SST&&!sst_seen){
            sst_seen=true;
            std::vector<BiffRecord> conts;
            for(size_t j=i+1;j<recs.size()&&recs[j].type==RecordType::CONTINUE;++j)
                conts.push_back(recs[j]);
            ctx.sst.parse(r,conts);
        } else if(r.type==RecordType::BOUNDSHEET&&r.data.size()>=8){
            if(r.data[5]==0){
                size_t consumed=0;
                std::string nm=decode_biff8_string(r.data.data()+6,r.data.size()-6,consumed,true);
                ctx.sheets.push_back({nm,r.u32(0)});
            }
        }
    }

    auto wb=std::make_unique<XlsWorkbook>();
    bool past_global=false; size_t sidx=0;
    std::unique_ptr<XlsSheet> cur;

    for(size_t i=0;i<recs.size();++i){
        auto& r=recs[i];
        if(r.type==RecordType::BOF){
            if(!past_global){past_global=true;continue;}
            std::string nm=sidx<ctx.sheets.size()?ctx.sheets[sidx].first:"Sheet"+std::to_string(sidx+1);
            cur=std::make_unique<XlsSheet>(nm); ++sidx; continue;
        }
        if(!cur) continue;
        if(r.type==RecordType::EOF_){wb->add_parsed_sheet(std::move(cur));cur.reset();continue;}

        if(r.type==RecordType::DIMENSION&&r.data.size()>=10){
            require_size(r,10,"DIMENSION");
            uint32_t rf=r.u32(0),rl=r.u32(4);
            uint16_t cf=r.u16(8),cl=r.data.size()>=12?r.u16(10):cf;
            cur->set_dimensions(rl>rf?rl-rf:0, cl>cf?uint32_t(cl-cf):0);
        } else if(r.type==RecordType::NUMBER&&r.data.size()>=14){
            Cell c;c.address={r.u16(0),r.u16(2)};c.type=CellType::Number;c.value=r.f64(6);
            c.style=xf_to_style(r.u16(4),ctx.xf_table,ctx.formats);cur->put_cell(c);
        } else if(r.type==RecordType::LABELSST&&r.data.size()>=10){
            require_size(r,10,"LABELSST");
            Cell c;c.address={r.u16(0),r.u16(2)};c.type=CellType::String;
            c.value=ctx.sst.get(r.u32(6));c.style=xf_to_style(r.u16(4),ctx.xf_table,ctx.formats);cur->put_cell(c);
        } else if(r.type==RecordType::LABEL&&r.data.size()>=6){
            Cell c;c.address={r.u16(0),r.u16(2)};c.type=CellType::String;
            size_t consumed=0;c.value=decode_biff8_string(r.data.data()+6,r.data.size()-6,consumed);cur->put_cell(c);
        } else if(r.type==RecordType::BOOLERR&&r.data.size()>=8){
            Cell c;c.address={r.u16(0),r.u16(2)};
            uint8_t val=r.data[6],is_err=r.data[7];
            if(is_err){static const char*EN[]={"#NULL!","#DIV/0!","#VALUE!","#REF!","#NAME?","#NUM!","#N/A"};
                c.type=CellType::Error;c.value=ErrorValue{val<7?EN[val]:"#ERR!"};}
            else{c.type=CellType::Boolean;c.value=bool(val!=0);}
            cur->put_cell(c);
        } else if(r.type==RecordType::RK&&r.data.size()>=10){
            require_size(r,10,"RK");
            Cell c;c.address={r.u16(0),r.u16(2)};c.type=CellType::Number;c.value=decode_rk(r.u32(6));
            c.style=xf_to_style(r.u16(4),ctx.xf_table,ctx.formats);cur->put_cell(c);
        } else if(r.type==RecordType::MULRK){
            require_size(r,4,"MULRK");
            int64_t cnt=(int64_t(r.data.size())-4)/6;
            if(cnt>0){
                uint16_t row_i=r.u16(0),col_f=r.u16(2);
                for(int64_t k=0;k<cnt;++k){
                    size_t base=4+size_t(k)*6; if(base+6>r.data.size()) break;
                    Cell c;c.address={uint32_t(row_i),uint32_t(col_f+k)};c.type=CellType::Number;
                    uint16_t xi=r.u16(base);uint32_t rv=r.u32(base+2);
                    c.value=decode_rk(rv);c.style=xf_to_style(xi,ctx.xf_table,ctx.formats);cur->put_cell(c);
                }
            }
        } else if(r.type==RecordType::FORMULA&&r.data.size()>=14){
            Cell c;c.address={r.u16(0),r.u16(2)};c.type=CellType::Formula;
            c.style=xf_to_style(r.u16(4),ctx.xf_table,ctx.formats);
            bool special=r.data[12]==0xFF&&r.data[13]==0xFF;
            if(special){
                uint8_t t=r.data[6];
                if(t==1){c.type=CellType::Boolean;c.value=bool(r.data[8]!=0);}
                else if(t==2){c.type=CellType::Error;c.value=ErrorValue{"#ERR!"};}
                else if(t==3){c.type=CellType::Blank;c.value=BlankValue{};}
                else c.value=std::string("<string>");
            } else {c.value=r.f64(6);}
            c.formula=std::nullopt;cur->put_cell(c);
        } else if(r.type==RecordType::BLANK&&r.data.size()>=6){
            Cell c;c.address={r.u16(0),r.u16(2)};c.type=CellType::Blank;c.value=BlankValue{};
            c.style=xf_to_style(r.u16(4),ctx.xf_table,ctx.formats);cur->put_cell(c);
        } else if(r.type==RecordType::MERGEDCELLS&&r.data.size()>=2){
            uint16_t cnt=r.u16(0);
            for(uint16_t mi=0;mi<cnt;++mi){
                size_t base=2+size_t(mi)*8;
                if(base+8>r.data.size()) break;
                uint16_t rf=r.u16(base),rl=r.u16(base+2);
                uint16_t cf=r.u16(base+4),cl=r.u16(base+6);
                cur->merges_.push_back({uint32_t(rf),uint32_t(cf),uint32_t(rl),uint32_t(cl)});
            }
        } else if(r.type!=RecordType::EOF_&&r.type!=RecordType::ROW&&
                  r.type!=RecordType::CONTINUE&&r.type!=RecordType::BLANK&&
                  r.type!=RecordType::MULBLANK&&r.type!=RecordType::BOF){
            warn(ParseWarning::Kind::UnsupportedRecord, "XLS",
                 "unhandled record type 0x" + [](uint16_t v){
                     char buf[8]; std::snprintf(buf,sizeof(buf),"%04X",v); return std::string(buf);
                 }(r.raw_type));
        }
    }
    return wb;
}

CellStyle XlsParser::xf_to_style(uint16_t xi,const XFTable& xf,const FormatRegistry& fmt){
    CellStyle s; try{s.format=fmt.get(xf.get(xi).format_index);}catch(...){}  return s;
}

} // namespace excellib::xls
