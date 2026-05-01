// XLS (BIFF8 形式) のパーサー実装
// .xls ファイルは OLE2 コンテナの中に BIFF8 (Binary Interchange File Format 8) が入っている
// BIFF8 はレコードの列で構成され、各レコードが種別コード+データを持つ
#include "excellib/xls_parser.hpp"
// std::memcpy: バイト列をコピーする C 標準関数
#include <cstring>
// std::sqrt, std::isnan など数学関数
#include <cmath>
// std::transform: 文字列を小文字化などに使う
#include <algorithm>
// std::ostringstream: 文字列を組み立てるストリーム
#include <sstream>
// std::ofstream: ファイルへの書き込み
#include <fstream>

namespace excellib::xls {

// ---- BIFF レコードのバウンドチェックヘルパー ----
// chk(): バイト列 d のオフセット o から n バイト読もうとしたとき、
//        バッファ範囲を超えないことを確認する
// 範囲外なら ParseError を投げる
static void chk(const std::vector<uint8_t>& d, size_t o, size_t n, const char* w) {
    if (o + n > d.size())
        throw ParseError(std::string("BIFF truncated reading ") + w);
}
// BiffRecord::u8(): オフセット o から 1バイト符号なし整数を読む
uint8_t  BiffRecord::u8 (size_t o) const { chk(data,o,1,"u8");  return data[o]; }
// BiffRecord::u16(): オフセット o から 2バイト符号なし整数をリトルエンディアンで読む
uint16_t BiffRecord::u16(size_t o) const {
    chk(data,o,2,"u16");
    // リトルエンディアン: 下位バイト (data[o]) + 上位バイト (data[o+1]<<8)
    return uint16_t(data[o]) | uint16_t(uint16_t(data[o+1])<<8);
}
// BiffRecord::u32(): オフセット o から 4バイト符号なし整数をリトルエンディアンで読む
uint32_t BiffRecord::u32(size_t o) const {
    chk(data,o,4,"u32");
    return uint32_t(data[o])|uint32_t(data[o+1])<<8|uint32_t(data[o+2])<<16|uint32_t(data[o+3])<<24;
}
// BiffRecord::i16(): オフセット o から 2バイト符号付き整数を読む (u16 の結果をキャスト)
int16_t  BiffRecord::i16(size_t o) const { return static_cast<int16_t>(u16(o)); }
// BiffRecord::f64(): オフセット o から 8バイト IEEE 754 倍精度浮動小数点数を読む
double   BiffRecord::f64(size_t o) const {
    chk(data,o,8,"f64");
    double v;
    // memcpy でバイト列を double のメモリ表現として解釈する (型パニングの安全な方法)
    std::memcpy(&v,data.data()+o,8);
    return v;
}

// decode_rk(): BIFF の RK 値 (コンパクトな数値表現) を double に変換する
// RK は整数または 100 倍した整数を 30ビットに詰め込んだ BIFF 独自の形式
double decode_rk(uint32_t rk) {
    double v;
    if (rk & 0x02) {
        // ビット1が立っている場合: 上位30ビットが整数値 (符号付き)
        // >> 2 で下位2ビットを捨てる
        v = static_cast<double>(static_cast<int32_t>(rk) >> 2);
    } else {
        // ビット1が立っていない場合: 上位30ビットが IEEE 754 倍精度の上位部分
        // 下位2ビットをクリアして64ビットの上位32ビットに配置
        uint64_t tmp = static_cast<uint64_t>(rk & 0xFFFFFFFC) << 32;
        std::memcpy(&v, &tmp, 8);    // バイト列を double として解釈
    }
    // ビット0が立っている場合: 値を100で割る (パーセント値など)
    if (rk & 0x01) v /= 100.0;
    return v;
}

// decode_biff8_string(): BIFF8 形式の文字列を UTF-8 に変換して返す
// BIFF8 文字列はヘッダー (文字数) + フラグ + 文字データで構成される
// data: 文字列データへのポインタ  avail: 残りバイト数
// consumed: 実際に消費したバイト数 (出力)  short_hdr: ヘッダーが1バイトか2バイトか
std::string decode_biff8_string(const uint8_t* data, size_t avail,
                                 size_t& consumed, bool short_hdr) {
    consumed = 0;
    // ヘッダーサイズ: short_hdr=true なら1バイト、false なら2バイト
    size_t hdr = short_hdr ? 1u : 2u;
    if (avail < hdr + 1) throw ParseError("BIFF8 string too short");
    // nc: 文字数 (バイト数ではなく文字数。UTF-16 なら2バイト/文字)
    uint16_t nc = short_hdr ? uint16_t(data[0])
                            : uint16_t(data[0]) | uint16_t(uint16_t(data[1])<<8);
    consumed += hdr;
    // flags: 圧縮フラグ・リッチテキストフラグ・拡張データフラグを持つビットマスク
    uint8_t flags = data[consumed++];
    // compressed: ビット0が0なら圧縮 (1バイト/文字)、1なら非圧縮 (2バイト/文字 = UTF-16LE)
    bool compressed = !(flags & 0x01);
    // has_rich: ビット3が立っていればリッチテキスト (フォント情報あり)
    bool has_rich   = (flags & 0x08) != 0;
    // has_ext: ビット2が立っていれば拡張データあり
    bool has_ext    = (flags & 0x04) != 0;
    uint16_t rt = 0; uint32_t ext = 0;
    if (has_rich) {
        // リッチテキストの場合: 追加で 2バイトのランカウントを読む
        if (consumed+2>avail) throw ParseError("BIFF8 rich-text truncated");
        rt = uint16_t(data[consumed])|uint16_t(uint16_t(data[consumed+1])<<8); consumed+=2;
    }
    if (has_ext) {
        // 拡張データの場合: 追加で 4バイトのサイズを読む
        if (consumed+4>avail) throw ParseError("BIFF8 ext-data truncated");
        ext = uint32_t(data[consumed])|uint32_t(data[consumed+1])<<8
             |uint32_t(data[consumed+2])<<16|uint32_t(data[consumed+3])<<24; consumed+=4;
    }
    // cb: 文字データのバイト数 (圧縮=1バイト/文字、非圧縮=2バイト/文字)
    size_t cb = compressed ? nc : size_t(nc)*2;
    if (consumed+cb>avail) throw ParseError("BIFF8 char data truncated");
    // 文字列を格納するバッファ (nc 文字分を事前確保)
    std::string r; r.reserve(nc);
    if (compressed) {
        // 1バイト圧縮: ASCII (< 0x80) はそのまま、それ以外は UTF-8 に変換
        for (uint16_t i=0;i<nc;++i) {
            uint8_t ch=data[consumed+i];
            if (ch<0x80) r+=char(ch);    // ASCII 範囲: そのまま
            else {
                // 0x80〜0xFF: 2バイト UTF-8 エンコード
                r+=char(0xC0|(ch>>6));    // 先頭バイト: 110xxxxx
                r+=char(0x80|(ch&0x3F)); // 後続バイト: 10xxxxxx
            }
        }
    } else {
        // UTF-16LE: 2バイト/文字を UTF-8 に変換
        for (uint16_t i=0;i<nc;++i) {
            // 2バイトをリトルエンディアンで結合して UTF-16 コードポイントを取得
            uint16_t cp=uint16_t(data[consumed+size_t(i)*2])|uint16_t(uint16_t(data[consumed+size_t(i)*2+1])<<8);
            if (cp<0x80) r+=char(cp);    // ASCII: 1バイト UTF-8
            else if (cp<0x800) {
                // U+0080〜U+07FF: 2バイト UTF-8
                r+=char(0xC0|(cp>>6));
                r+=char(0x80|(cp&0x3F));
            }
            else {
                // U+0800〜U+FFFF: 3バイト UTF-8
                r+=char(0xE0|(cp>>12));
                r+=char(0x80|((cp>>6)&0x3F));
                r+=char(0x80|(cp&0x3F));
            }
        }
    }
    // 文字データ + リッチテキストのランデータ (rt*4) + 拡張データ (ext) を消費済みに加算
    consumed += cb + size_t(rt)*4 + ext;
    return r;
}

// SharedStringTable::parse(): SST (共有文字列テーブル) レコードをパースする
// SST レコードには Excel ファイル内の全文字列が格納されている
// sst: SST レコード  conts: CONTINUE レコードのリスト (SST が複数レコードに分割される場合)
void SharedStringTable::parse(const BiffRecord& sst, const std::vector<BiffRecord>& conts) {
    if (sst.data.size()<8) throw ParseError("SST too short");
    // オフセット 4 から 4バイト: 文字列の総数
    uint32_t total = sst.u32(4); strings_.reserve(total);
    // SST + CONTINUE レコードのデータを1つのバッファに結合
    std::vector<uint8_t> buf(sst.data);
    for (auto& c:conts) buf.insert(buf.end(),c.data.begin(),c.data.end());
    // 最初の 8バイトはヘッダー (総数など) なのでスキップ
    const uint8_t* p=buf.data()+8; size_t left=buf.size()-8;
    // total 個の文字列をパース
    for (uint32_t i=0;i<total;++i) {
        if (!left) throw ParseError("SST exhausted");
        size_t consumed=0;
        // BIFF8 文字列を UTF-8 に変換して追加
        strings_.push_back(decode_biff8_string(p,left,consumed));
        if (consumed>left) throw ParseError("SST overrun");
        p+=consumed; left-=consumed;
    }
}
// SharedStringTable::get(): インデックスで共有文字列を取得する
const std::string& SharedStringTable::get(uint32_t i) const {
    if (i>=strings_.size()) throw RangeError("SST index out of range");
    return strings_[i];
}

// FormatRegistry のコンストラクタ: 組み込み数値フォーマットを登録する
// BIFF8 には ID 0〜49 の組み込みフォーマットが定義されている
FormatRegistry::FormatRegistry() {
    // 組み込みフォーマットの ID と書式文字列のリスト
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
// FormatRegistry::add(): カスタム数値フォーマットを追加する
void FormatRegistry::add(uint16_t i, const std::string& s) { formats_[i]=s; }
// FormatRegistry::get(): フォーマット ID に対応する FormatInfo を返す
FormatInfo FormatRegistry::get(uint16_t i) const {
    FormatInfo fi; fi.format_index=i;
    auto it=formats_.find(i);
    if (it!=formats_.end()) { fi.format_string=it->second; fi.is_date_format=is_date_format(i); }
    return fi;
}
// FormatRegistry::is_date_format(): フォーマット ID が日付形式かどうかを判定する
bool FormatRegistry::is_date_format(uint16_t i) const {
    // ID 14〜22 と 45〜47 は組み込みの日付/時刻フォーマット
    if ((i>=14&&i<=22)||(i>=45&&i<=47)) return true;
    // カスタムフォーマットの場合は書式文字列で判定
    auto it=formats_.find(i);
    return it!=formats_.end()&&looks_like_date(it->second);
}
// FormatRegistry::looks_like_date(): 書式文字列が日付形式かどうかを推定する
bool FormatRegistry::looks_like_date(const std::string& f) {
    // 小文字化して検索
    std::string lo=f; std::transform(lo.begin(),lo.end(),lo.begin(),::tolower);
    // 'y' (year) があれば日付形式
    return lo.find('y')!=std::string::npos||
           // 'd' (day) があれば日付形式
           (lo.find('d')!=std::string::npos)||
           // 'm' (month) があり 'h' (hour) がなければ日付形式
           (lo.find('m')!=std::string::npos&&lo.find('h')==std::string::npos);
}

// XFTable::add(): XF レコード (セルのスタイル情報) を追加する
void XFTable::add(const XFRecord& x) { records_.push_back(x); }
// XFTable::get(): インデックスで XF レコードを取得する
const XFRecord& XFTable::get(uint16_t i) const {
    if (i>=records_.size()) throw RangeError("XF index out of range");
    return records_[i];
}

// XlsSheet::fill_type(): セルの値の型に応じて CellType を設定する
void XlsSheet::fill_type(Cell& c) {
    // std::visit で CellValue の実際の型に応じて CellType を設定する
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
// XlsSheet::put_cell(): セルをシートのデータ構造に格納し、行数・列数を更新する
void XlsSheet::put_cell(const Cell& c) {
    // data_[行][列] = セル という 2段階のマップに格納
    data_[c.address.row][c.address.col]=c;
    // 行数・列数を最大値に更新 (row/col は 0-indexed なので +1 が個数)
    if (c.address.row+1>row_count_) row_count_=c.address.row+1;
    if (c.address.col+1>col_count_) col_count_=c.address.col+1;
}
// XlsSheet::set_dimensions(): 行数・列数を直接セットする (DIMENSION レコードから)
void XlsSheet::set_dimensions(uint32_t r,uint32_t c){row_count_=r;col_count_=c;}
// mk_blank(): 指定位置の空セルを作成するローカルヘルパー関数
static Cell mk_blank(uint32_t r,uint32_t c){Cell x;x.address={r,c};x.type=CellType::Blank;x.value=BlankValue{};return x;}
// XlsSheet::cell(): 行・列インデックスでセルを取得する (空・範囲外は空セルを返す)
Cell XlsSheet::cell(uint32_t r,uint32_t c) const {
    // 行マップで行 r を探す
    auto ri=data_.find(r);
    if (ri!=data_.end()){
        // 列マップで列 c を探す
        auto ci=ri->second.find(c);
        if(ci!=ri->second.end()) return ci->second;    // 見つかればそのセルを返す
    }
    // 見つからなければ空セルを返す
    return mk_blank(r,c);
}
// CellAddress / A1 記法でのアクセスも委譲
Cell XlsSheet::cell(const CellAddress& a) const {return cell(a.row,a.col);}
Cell XlsSheet::cell(const std::string& a) const {return cell(CellAddress::from_a1(a));}
// XlsSheet::try_cell(): 空・範囲外のセルは nullopt を返す安全版
std::optional<Cell> XlsSheet::try_cell(uint32_t r,uint32_t c) const {
    auto ri=data_.find(r); if(ri==data_.end()) return std::nullopt;
    auto ci=ri->second.find(c); if(ci==ri->second.end()||ci->second.is_blank()) return std::nullopt;
    return ci->second;
}
// XlsSheet::row(): 指定行の全セルをベクタで返す
std::vector<Cell> XlsSheet::row(uint32_t r) const {
    std::vector<Cell> out;
    auto ri=data_.find(r); if(ri!=data_.end()) for(auto&[c,x]:ri->second) out.push_back(x);
    return out;
}
// XlsSheet::cells(): 全非空セルをベクタで返す
std::vector<Cell> XlsSheet::cells() const {
    std::vector<Cell> out;
    // 全行・全列を走査し、空でないセルだけを追加
    for(auto&[r,cols]:data_) for(auto&[c,x]:cols) if(!x.is_blank()) out.push_back(x);
    return out;
}
// XlsSheet::for_each_cell(): 全非空セルにコールバック fn を呼ぶ
void XlsSheet::for_each_cell(std::function<void(const Cell&)> fn) const {
    for(auto&[r,cols]:data_) for(auto&[c,x]:cols) fn(x);
}
// XlsSheet::set_cell(): 値をセルに書き込む (行・列インデックス版)
void XlsSheet::set_cell(uint32_t r,uint32_t c,const CellValue& v){Cell x;x.address={r,c};x.value=v;fill_type(x);put_cell(x);}
// XlsSheet::set_cell(): A1 記法でセルに値を書き込む
void XlsSheet::set_cell(const std::string& a,const CellValue& v){auto addr=CellAddress::from_a1(a);set_cell(addr.row,addr.col,v);}
// XlsSheet::set_formula(): A1 記法でセルに数式を書き込む
void XlsSheet::set_formula(const std::string& a,const std::string& f){
    auto addr=CellAddress::from_a1(a);
    Cell c; c.address=addr; c.type=CellType::Formula; c.formula=f; c.value=BlankValue{};
    put_cell(c);
}

// XlsWorkbook のメンバ関数実装
// add_parsed_sheet(): パース済みシートをワークブックに追加する
void XlsWorkbook::add_parsed_sheet(std::unique_ptr<XlsSheet> s){sheets_.push_back(std::move(s));}
// sheet(size_t): インデックスでシートを取得する
Sheet& XlsWorkbook::sheet(size_t i){if(i>=sheets_.size())throw RangeError("Sheet index OOB");return *sheets_[i];}
const Sheet& XlsWorkbook::sheet(size_t i) const {if(i>=sheets_.size())throw RangeError("Sheet index OOB");return *sheets_[i];}
// sheet(string): 名前でシートを取得する
Sheet& XlsWorkbook::sheet(const std::string& n){for(auto&s:sheets_) if(s->name()==n) return *s;throw RangeError("Sheet not found: "+n);}
// sheet_names(): 全シート名をベクタで返す
std::vector<std::string> XlsWorkbook::sheet_names() const {std::vector<std::string> r;for(auto&s:sheets_)r.push_back(s->name());return r;}
// add_sheet(): 新しいシートを追加する (重複名はエラー)
Sheet& XlsWorkbook::add_sheet(const std::string& n){for(auto&s:sheets_)if(s->name()==n)throw WriteError("Dup sheet: "+n);sheets_.push_back(std::make_unique<XlsSheet>(n));return*sheets_.back();}
// remove_sheet(): インデックスでシートを削除する
void XlsWorkbook::remove_sheet(size_t i){if(i>=sheets_.size())throw RangeError("OOB");sheets_.erase(sheets_.begin()+static_cast<ptrdiff_t>(i));}
// rename_sheet(): シートをリネームする (重複名チェックあり)
void XlsWorkbook::rename_sheet(size_t i,const std::string& n){
    if(i>=sheets_.size())throw RangeError("OOB");
    for(size_t j=0;j<sheets_.size();++j)if(j!=i&&sheets_[j]->name()==n)throw WriteError("Name exists");
    // 旧シートの全セルを新シートにコピーして入れ替える
    auto old=std::move(sheets_[i]);auto ns=std::make_unique<XlsSheet>(n);
    old->for_each_cell([&](const Cell& c){ns->put_cell(c);});
    ns->set_dimensions(old->row_count(),old->col_count());sheets_[i]=std::move(ns);
}
// XlsWorkbook::save(): ファイルに保存する (XLS として保存するか XLSX として保存するか)
void XlsWorkbook::save(const std::string& p,const SaveOptions& o) const {
    // Auto フォーマットなら XLS として、明示されていればその形式で保存
    auto b=to_bytes(o.format==FileFormat::Auto?FileFormat::XLS:o.format,o);
    std::ofstream f(p,std::ios::binary);if(!f)throw IOError("Cannot write: "+p);
    f.write(reinterpret_cast<const char*>(b.data()),static_cast<std::streamsize>(b.size()));
}
// XlsWorkbook::to_bytes(): バイト列として取得する (XLS 書き込みは未実装)
std::vector<uint8_t> XlsWorkbook::to_bytes(FileFormat,const SaveOptions&) const {
    // XLS フォーマットの書き込みは複雑なため未実装。XLSX として保存することを推奨
    throw WriteError("XLS write not implemented. Use save() with FileFormat::XLSX.");
}

// OLE2 ファイルのマジックバイト (先頭 8 バイト)
static constexpr uint8_t OLE2_MAGIC[8]={0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1};

// XlsParser::extract_workbook_stream(): OLE2 コンテナから "Workbook" ストリームを取り出す
// OLE2 は Windows の「複合ドキュメント形式」で、ファイルシステムのような構造を持つ
std::vector<uint8_t> XlsParser::extract_workbook_stream(const std::vector<uint8_t>& raw) {
    if (raw.size()<512) throw ParseError("File too small for OLE2");
    if (std::memcmp(raw.data(),OLE2_MAGIC,8)!=0) throw ParseError("Not OLE2 (bad magic)");

    // リトルエンディアンの 16ビット・32ビット値を読むローカルラムダ
    auto r16=[&](size_t o)->uint16_t{return uint16_t(raw[o])|uint16_t(uint16_t(raw[o+1])<<8);};
    auto r32=[&](size_t o)->uint32_t{return uint32_t(raw[o])|uint32_t(raw[o+1])<<8|uint32_t(raw[o+2])<<16|uint32_t(raw[o+3])<<24;};

    // OLE2 ヘッダーから各種パラメータを読む
    uint16_t ss_pow = r16(30);    // セクターサイズの指数 (通常 9 → 512バイト)
    uint16_t ms_pow = r16(32);    // ミニセクターサイズの指数 (通常 6 → 64バイト)
    uint32_t num_fat= r32(44);    // FAT セクターの数
    uint32_t dir_s  = r32(48);    // ディレクトリの先頭セクター番号
    uint32_t mc     = r32(56);    // ミニストリームカットオフ (このサイズ以下はミニ FAT を使う)
    uint32_t mfs    = r32(60);    // ミニ FAT の先頭セクター番号
    size_t   ss     = size_t(1)<<ss_pow;    // セクターサイズ (バイト)
    size_t   ms     = size_t(1)<<ms_pow;    // ミニセクターサイズ (バイト)
    // so(): セクター番号 sid をオフセットに変換 (先頭 512 バイトがヘッダー)
    auto     so     = [&](uint32_t sid){return 512+size_t(sid)*ss;};

    // FAT (File Allocation Table): セクターの連結情報を格納するテーブル
    std::vector<uint32_t> fat;
    // ヘッダーのオフセット 76 から FAT セクター番号を読む (最大 109 個)
    for (uint32_t i=0;i<std::min(num_fat,uint32_t(109));++i) {
        uint32_t fs=r32(76+i*4); if(fs>=0xFFFFFFFE) break;    // 0xFFFFFFFE は終端マーカー
        size_t off=so(fs); if(off+ss>raw.size()) throw ParseError("FAT OOB");
        // FAT セクター内の全エントリを読み込む
        for (size_t j=0;j<ss/4;++j){uint32_t e;std::memcpy(&e,raw.data()+off+j*4,4);fat.push_back(e);}
    }

    // read_chain(): FAT チェーンを辿ってセクター群をバイト列として読み出す
    auto read_chain=[&](uint32_t start)->std::vector<uint8_t>{
        std::vector<uint8_t> out; uint32_t sid=start;
        while(sid<0xFFFFFFFE){    // 0xFFFFFFFE は終端
            if(sid>=fat.size()) throw ParseError("FAT chain broken");
            size_t off=so(sid); if(off+ss>raw.size()) throw ParseError("Sector OOB");
            out.insert(out.end(),raw.data()+off,raw.data()+off+ss);
            sid=fat[sid];    // 次のセクターへ
        }
        return out;
    };

    // ディレクトリエントリを読み込む
    auto dir=read_chain(dir_s);
    constexpr size_t DE=128;    // ディレクトリエントリのサイズ (128 バイト)
    if(dir.size()<DE) throw ParseError("Dir too small");

    // ルートエントリ (インデックス 0) からミニストリームの先頭セクターとサイズを取得
    uint32_t root_start=r32(dir.data()+116-dir.data()); /* offset 116 in first dir entry */
    {/* use memcpy for portability */
        std::memcpy(&root_start,dir.data()+116,4);    // オフセット 116: ストリームの先頭セクター
    }
    uint64_t root_sz=0; std::memcpy(&root_sz,dir.data()+120,8);    // オフセット 120: ストリームのサイズ

    // ミニストリームとミニ FAT を読み込む
    std::vector<uint8_t> mini_stream;
    std::vector<uint32_t> minifat;
    if (root_start<0xFFFFFFFE) {
        // ルートエントリのストリームがミニストリームの実体
        mini_stream=read_chain(root_start); mini_stream.resize(static_cast<size_t>(root_sz));
        // ミニ FAT を読み込む
        uint32_t mfs_sid=mfs;
        while(mfs_sid<0xFFFFFFFE){
            if(mfs_sid>=fat.size()) break;
            size_t off=so(mfs_sid); if(off+ss>raw.size()) break;
            for(size_t j=0;j<ss/4;++j){uint32_t e;std::memcpy(&e,raw.data()+off+j*4,4);minifat.push_back(e);}
            mfs_sid=fat[mfs_sid];
        }
    }
    // read_mini(): ミニ FAT チェーンを辿ってミニストリームからデータを読み出す
    auto read_mini=[&](uint32_t start,uint64_t sz)->std::vector<uint8_t>{
        std::vector<uint8_t> out; uint32_t sid=start;
        while(sid<0xFFFFFFFE){
            if(sid>=minifat.size()) throw ParseError("MiniFAT broken");
            size_t off=size_t(sid)*ms; if(off+ms>mini_stream.size()) throw ParseError("Mini OOB");
            out.insert(out.end(),mini_stream.data()+off,mini_stream.data()+off+ms);
            sid=minifat[sid];
        }
        out.resize(static_cast<size_t>(sz)); return out;
    };

    // ディレクトリエントリを順に走査して "Workbook" または "Book" ストリームを探す
    size_t ne=dir.size()/DE;
    for(size_t i=1;i<ne;++i){    // 0 はルートエントリなのでスキップ
        const uint8_t* e=dir.data()+i*DE;
        if(e[66]!=2) continue;    // オブジェクトタイプ 2 = ストリーム (ファイル相当)
        // エントリ名の文字数 (バイト数)
        uint16_t nl=uint16_t(e[64])|uint16_t(uint16_t(e[65])<<8);
        // UTF-16LE の名前を ASCII として読む (Workbook は ASCII 範囲)
        std::string nm; size_t chars=nl>1?(nl-2)/2:0;
        for(size_t j=0;j<chars;++j){uint16_t ch;std::memcpy(&ch,e+j*2,2);if(ch<128)nm+=char(ch);}
        // "Workbook" (Excel 97+) または "Book" (Excel 95) を探す
        if(nm=="Workbook"||nm=="Book"){
            uint32_t s=0; std::memcpy(&s,e+116,4);    // ストリームの先頭セクター
            uint64_t z=0; std::memcpy(&z,e+120,8);    // ストリームのバイトサイズ
            // サイズが mc (ミニストリームカットオフ) 未満ならミニストリームから読む
            if(z<mc&&!mini_stream.empty()) return read_mini(s,z);
            // そうでなければ通常 FAT チェーンから読む
            auto d=read_chain(s); d.resize(static_cast<size_t>(z)); return d;
        }
    }
    throw ParseError("OLE2: no Workbook stream found");
}

// XlsParser::read_records(): BIFF ストリームを BIFF レコードのリストに変換する
std::vector<BiffRecord> XlsParser::read_records(const std::vector<uint8_t>& stream) {
    std::vector<BiffRecord> recs; size_t pos=0;
    // 4バイト (レコード種別 2バイト + データ長 2バイト) が読める間ループ
    while(pos+4<=stream.size()){
        uint16_t rt,len;
        // レコード種別コードをリトルエンディアンで読む
        std::memcpy(&rt, stream.data()+pos,  2);
        // データ長をリトルエンディアンで読む
        std::memcpy(&len,stream.data()+pos+2,2);
        pos+=4;
        // データ長分のバイトが残っていることを確認
        if(pos+len>stream.size()) throw ParseError("BIFF record overruns stream");
        BiffRecord r; r.raw_type=rt;
        // 数値を RecordType 列挙型に変換 (未知の種別は UNKNOWN になる)
        r.type=static_cast<RecordType>(rt);
        // データ部分を BiffRecord にコピー
        r.data.assign(stream.data()+pos,stream.data()+pos+len); pos+=len;
        recs.push_back(std::move(r));
    }
    return recs;
}

// XlsParser::parse(): XLS ファイルのバイト列をパースして XlsWorkbook を返す
std::unique_ptr<XlsWorkbook> XlsParser::parse(const std::vector<uint8_t>& raw) {
    // OLE2 コンテナから Workbook ストリームを取り出す
    auto stream=extract_workbook_stream(raw);
    // ストリームを BIFF レコードのリストに変換
    auto recs  =read_records(stream);
    GlobalContext ctx; bool sst_seen=false;

    // 第1パス: グローバルコンテキスト (共有文字列・スタイル・シート情報) を収集する
    for(size_t i=0;i<recs.size();++i){
        auto& r=recs[i];
        if(r.type==RecordType::FORMAT&&r.data.size()>=3){
            // FORMAT レコード: カスタム数値フォーマットを登録する
            uint16_t id=r.u16(0); size_t consumed=0;
            ctx.formats.add(id,decode_biff8_string(r.data.data()+2,r.data.size()-2,consumed));
        } else if(r.type==RecordType::XF&&r.data.size()>=4){
            // XF レコード: セルのスタイル (フォント・フォーマット) 情報
            XFRecord xf; xf.font_index=r.u16(0); xf.format_index=r.u16(2);
            if(r.data.size()>4) xf.is_style_xf=(r.data[4]&0x04)!=0;
            ctx.xf_table.add(xf);
        } else if(r.type==RecordType::SST&&!sst_seen){
            // SST レコード: 共有文字列テーブル (1ファイルに1回だけ現れる)
            sst_seen=true;
            // SST は大きいため後続の CONTINUE レコードに分割されることがある
            std::vector<BiffRecord> conts;
            for(size_t j=i+1;j<recs.size()&&recs[j].type==RecordType::CONTINUE;++j)
                conts.push_back(recs[j]);
            ctx.sst.parse(r,conts);
        } else if(r.type==RecordType::BOUNDSHEET&&r.data.size()>=8){
            // BOUNDSHEET レコード: シートの名前とストリーム内のオフセットを持つ
            if(r.data[5]==0){    // シート種別 0 = ワークシート (それ以外はマクロシートなど)
                size_t consumed=0;
                // オフセット 6 からシート名を読む (short_hdr=true で 1バイトヘッダー)
                std::string nm=decode_biff8_string(r.data.data()+6,r.data.size()-6,consumed,true);
                ctx.sheets.push_back({nm,r.u32(0)});    // シート名とストリームオフセット
            }
        }
    }

    auto wb=std::make_unique<XlsWorkbook>();
    bool past_global=false; size_t sidx=0;    // 現在処理中のシートのインデックス
    std::unique_ptr<XlsSheet> cur;    // 現在構築中のシート

    // 第2パス: セルデータを読み込んでシートを構築する
    for(size_t i=0;i<recs.size();++i){
        auto& r=recs[i];
        if(r.type==RecordType::BOF){
            // BOF (Beginning Of File): ストリームの開始マーカー
            if(!past_global){past_global=true;continue;}    // 最初の BOF はグローバルストリームの開始
            // 2つ目以降の BOF はシートの開始
            std::string nm=sidx<ctx.sheets.size()?ctx.sheets[sidx].first:"Sheet"+std::to_string(sidx+1);
            cur=std::make_unique<XlsSheet>(nm); ++sidx; continue;
        }
        if(!cur) continue;    // BOF が来る前はシートがないのでスキップ
        if(r.type==RecordType::EOF_){
            // EOF (End Of File): シートの終了マーカー
            wb->add_parsed_sheet(std::move(cur)); cur.reset(); continue;
        }

        if(r.type==RecordType::DIMENSION&&r.data.size()>=10){
            // DIMENSION レコード: シートの使用範囲を指定する
            uint32_t rf=r.u32(0),rl=r.u32(4);    // 開始行・終了行
            uint16_t cf=r.u16(8),cl=r.data.size()>=12?r.u16(10):cf;    // 開始列・終了列
            cur->set_dimensions(rl>rf?rl-rf:0, cl>cf?uint32_t(cl-cf):0);
        } else if(r.type==RecordType::NUMBER&&r.data.size()>=14){
            // NUMBER レコード: IEEE 754 倍精度浮動小数点数のセル値
            Cell c;c.address={r.u16(0),r.u16(2)};c.type=CellType::Number;c.value=r.f64(6);
            c.style=xf_to_style(r.u16(4),ctx.xf_table,ctx.formats);cur->put_cell(c);
        } else if(r.type==RecordType::LABELSST&&r.data.size()>=10){
            // LABELSST レコード: 共有文字列テーブルへのインデックスを持つ文字列セル
            Cell c;c.address={r.u16(0),r.u16(2)};c.type=CellType::String;
            c.value=ctx.sst.get(r.u32(6));    // SST から文字列を取得
            c.style=xf_to_style(r.u16(4),ctx.xf_table,ctx.formats);cur->put_cell(c);
        } else if(r.type==RecordType::LABEL&&r.data.size()>=6){
            // LABEL レコード: 文字列を直接持つセル (旧形式)
            Cell c;c.address={r.u16(0),r.u16(2)};c.type=CellType::String;
            size_t consumed=0;c.value=decode_biff8_string(r.data.data()+6,r.data.size()-6,consumed);cur->put_cell(c);
        } else if(r.type==RecordType::BOOLERR&&r.data.size()>=8){
            // BOOLERR レコード: 真偽値またはエラー値
            Cell c;c.address={r.u16(0),r.u16(2)};
            uint8_t val=r.data[6];    // 値 (BOOL の場合: 0=FALSE, 1=TRUE)
            uint8_t is_err=r.data[7]; // 0=BOOL, 1=ERROR
            if(is_err){
                // エラー値: val の値に対応するエラーコードを設定
                static const char*EN[]={"#NULL!","#DIV/0!","#VALUE!","#REF!","#NAME?","#NUM!","#N/A"};
                c.type=CellType::Error;c.value=ErrorValue{val<7?EN[val]:"#ERR!"};}
            else{c.type=CellType::Boolean;c.value=bool(val!=0);}
            cur->put_cell(c);
        } else if(r.type==RecordType::RK&&r.data.size()>=10){
            // RK レコード: コンパクトな数値表現 (RK 値) を持つセル
            Cell c;c.address={r.u16(0),r.u16(2)};c.type=CellType::Number;
            c.value=decode_rk(r.u32(6));    // RK 値を double に変換
            c.style=xf_to_style(r.u16(4),ctx.xf_table,ctx.formats);cur->put_cell(c);
        } else if(r.type==RecordType::MULRK&&r.data.size()>=10){
            // MULRK レコード: 複数の RK 値をまとめて持つレコード (連続した列に数値がある場合)
            uint16_t row_i=r.u16(0);    // 行インデックス
            uint16_t col_f=r.u16(2);    // 開始列インデックス
            size_t cnt=(r.data.size()-4)/6;    // エントリ数 (各エントリは 6バイト: 2バイト XF + 4バイト RK)
            for(size_t k=0;k<cnt;++k){
                size_t base=4+k*6; if(base+6>r.data.size()) break;
                Cell c;c.address={uint32_t(row_i),uint32_t(col_f+k)};c.type=CellType::Number;
                uint16_t xi=r.u16(base);    // XF インデックス
                uint32_t rv=r.u32(base+2);  // RK 値
                c.value=decode_rk(rv);c.style=xf_to_style(xi,ctx.xf_table,ctx.formats);cur->put_cell(c);
            }
        } else if(r.type==RecordType::FORMULA&&r.data.size()>=14){
            // FORMULA レコード: 数式を持つセル
            Cell c;c.address={r.u16(0),r.u16(2)};c.type=CellType::Formula;
            c.style=xf_to_style(r.u16(4),ctx.xf_table,ctx.formats);
            // オフセット 12-13 が 0xFF,0xFF なら「特殊な数式結果」
            bool special=r.data[12]==0xFF&&r.data[13]==0xFF;
            if(special){
                uint8_t t=r.data[6];    // 特殊値の種別
                if(t==1){c.type=CellType::Boolean;c.value=bool(r.data[8]!=0);}    // 真偽値
                else if(t==2){c.type=CellType::Error;c.value=ErrorValue{"#ERR!"};}  // エラー
                else if(t==3){c.type=CellType::Blank;c.value=BlankValue{};}          // 空
                else c.value=std::string("<string>");    // 文字列 (実際の値は STRING レコードで来る)
            } else {c.value=r.f64(6);}    // 通常の数値結果
            c.formula="<formula>";cur->put_cell(c);    // 実際の数式文字列は簡易表現で代替
        } else if(r.type==RecordType::BLANK&&r.data.size()>=6){
            // BLANK レコード: スタイル情報を持つ空セル
            Cell c;c.address={r.u16(0),r.u16(2)};c.type=CellType::Blank;c.value=BlankValue{};
            c.style=xf_to_style(r.u16(4),ctx.xf_table,ctx.formats);cur->put_cell(c);
        }
    }
    return wb;
}

// XlsParser::xf_to_style(): XF テーブルとフォーマットレジストリから CellStyle を生成する
CellStyle XlsParser::xf_to_style(uint16_t xi,const XFTable& xf,const FormatRegistry& fmt){
    CellStyle s;
    // XF インデックスからフォーマット ID を取得し、FormatInfo を生成
    try{s.format=fmt.get(xf.get(xi).format_index);}catch(...){}
    return s;
}

} // namespace excellib::xls
