// XLSX パーサーおよびライターの実装ファイル
// XLSX は ZIP アーカイブ内に XML ファイルを格納した形式 (OOXML 規格)
// 外部ライブラリ不使用で ZIP 読み込み・DEFLATE 展開・XML 解析を自前実装している
#include "excellib/xlsx_parser.hpp"
// deflate.hpp: 自前 DEFLATE 解凍 + CRC32 実装 (zlib 不使用)
#include "../common/deflate.hpp"
// std::memcmp: バイト列比較 (ZIP シグネチャの確認に使う)
#include <cstring>
// std::transform: 文字列を小文字に変換するために使う
#include <algorithm>
// std::ostringstream: 文字列を組み立てるストリーム
#include <sstream>
// std::ofstream: ファイル書き込みストリーム
#include <fstream>
// std::runtime_error などの例外クラス
#include <stdexcept>

namespace excellib::xlsx {

// ============================================================
//  XML ヘルパー関数 (外部依存なし)
// ============================================================

// xml_attr(): XML 要素文字列から属性値を取り出す
// el:   "<c r=\"B3\" t=\"s\" s=\"2\"/>" のような要素文字列
// attr: 取り出したい属性名 (例: "r", "t", "s")
// 戻り値: 属性値の文字列。属性が存在しなければ空文字列
static std::string xml_attr(const std::string& el, const std::string& attr) {
    // 検索する文字列を作成: 例 attr="  →  r="
    std::string needle = attr + "=\"";
    // needle を el 内で検索
    auto pos = el.find(needle);
    if (pos == std::string::npos) return {};   // 見つからなければ空を返す
    // needle の末尾 (属性値の開始位置) にスキップ
    pos += needle.size();
    // 次の " を探して属性値の終端を見つける
    auto end = el.find('"', pos);
    // end が見つからなければ空。見つかれば pos から end の手前まで切り出す
    return end == std::string::npos ? std::string{} : el.substr(pos, end-pos);
}

// xml_text(): XML 文字列から特定タグのテキストコンテンツを取り出す
// xml: XML 文字列全体, tag: タグ名 (例: "v", "f")
// 戻り値: <tag>...</tag> の間のテキスト。タグが見つからなければ空文字列
static std::string xml_text(const std::string& xml, const std::string& tag) {
    // 開始タグと終了タグの文字列を組み立てる
    std::string open  = "<" + tag + ">";   // 例: "<v>"
    std::string close = "</" + tag + ">"; // 例: "</v>"
    // 開始タグを検索
    auto pos = xml.find(open);
    if (pos == std::string::npos) return {};   // 開始タグがなければ空
    // 開始タグの直後から検索開始
    pos += open.size();
    // 終了タグを検索
    auto end = xml.find(close, pos);
    // 終了タグがなければ空。あれば開始〜終了の間を切り出す
    return end == std::string::npos ? std::string{} : xml.substr(pos, end-pos);
}

// xml_unescape(): XML のエスケープシーケンスを元の文字に戻す
// XML では & < > ' " の 5 文字を直接書けないため &amp; &lt; などに置換されている
// 例: "&amp;lt;" → "&lt;"
static std::string xml_unescape(const std::string& s) {
    std::string r;
    r.reserve(s.size());   // 変換後は元と同じかそれ以下のサイズなので予約
    for (size_t i = 0; i < s.size(); ) {
        if (s[i] != '&') { r += s[i++]; continue; }   // & 以外はそのままコピー
        // & から始まるシーケンスを判定
        if (s.substr(i,4) == "&lt;")   { r += '<';  i += 4; }   // &lt; → <
        else if (s.substr(i,4) == "&gt;")   { r += '>';  i += 4; }   // &gt; → >
        else if (s.substr(i,5) == "&amp;")  { r += '&';  i += 5; }   // &amp; → &
        else if (s.substr(i,6) == "&apos;") { r += '\''; i += 6; }   // &apos; → '
        else if (s.substr(i,6) == "&quot;") { r += '"';  i += 6; }   // &quot; → "
        else { r += s[i++]; }   // 不明なシーケンスはそのままコピー
    }
    return r;
}

// ============================================================
//  ZipReader — EOCD 探索修正済み・stoul 安全版
// ============================================================

// rd32(): 4 バイトを Little Endian で読み込んで uint32_t を返すヘルパー
// p[0] は最下位バイト、p[3] は最上位バイト
static uint32_t rd32(const uint8_t* p) {
    return uint32_t(p[0])|(uint32_t(p[1])<<8)|(uint32_t(p[2])<<16)|(uint32_t(p[3])<<24);
}
// rd16(): 2 バイトを Little Endian で読み込んで uint16_t を返すヘルパー
static uint16_t rd16(const uint8_t* p) {
    return uint16_t(p[0])|(uint16_t(p[1])<<8);
}

// ZipReader::inflate(): raw DEFLATE ストリームを展開する
// excellib::detail::deflate_decompress() (自前実装) に委譲する
// comp: 圧縮データ, comp_sz: 圧縮サイズ, uncomp_sz: 展開後の目安サイズ
std::vector<uint8_t> ZipReader::inflate(const uint8_t* comp, size_t comp_sz, size_t uncomp_sz) {
    return excellib::detail::deflate_decompress(comp, comp_sz, uncomp_sz);
}

// ZipReader::parse(): ZIP アーカイブのバイト列を解析してエントリマップを作る
// ZIP 構造:
//   [ローカルファイルヘッダー + データ] × n
//   [セントラルディレクトリ (CD)]
//   [EOCD (End Of Central Directory) レコード]
// パース手順: EOCD → CD → 各ローカルファイルヘッダーの順で解析する
void ZipReader::parse(const std::vector<uint8_t>& data) {
    // EOCD は最小 22 バイトなので、それより小さければ ZIP でない
    if (data.size() < 22) throw ParseError("ZIP file too small");

    // EOCD 探索: ファイル末尾から前方へスキャンして署名 "PK\x05\x06" を探す
    // コメント付き ZIP もあるため末尾から探す。署名は 0x50,0x4B,0x05,0x06
    size_t eocd = std::string::npos;
    // 注意: signed の ptrdiff_t を使わないと i==0 の境界チェックが失敗する
    for (ptrdiff_t i = static_cast<ptrdiff_t>(data.size()) - 22; i >= 0; --i) {
        auto j = static_cast<size_t>(i);   // signed → unsigned に変換してインデックスに使う
        // EOCD シグネチャ: 0x50 0x4B 0x05 0x06 ("PK\x05\x06")
        if (data[j]==0x50 && data[j+1]==0x4B && data[j+2]==0x05 && data[j+3]==0x06) {
            eocd = j; break;   // 見つかったら位置を記録してループ脱出
        }
    }
    if (eocd == std::string::npos) throw ParseError("ZIP: EOCD record not found");

    // EOCD オフセット +16: セントラルディレクトリの開始オフセット (4 バイト)
    uint32_t cd_off   = rd32(data.data() + eocd + 16);
    // EOCD オフセット +10: セントラルディレクトリのエントリ数 (2 バイト)
    uint16_t cd_count = rd16(data.data() + eocd + 10);

    // CD オフセットがファイルサイズを超えていれば壊れた ZIP
    if (cd_off > data.size()) throw ParseError("ZIP: central directory offset beyond EOF");

    // セントラルディレクトリのパース: 各エントリの情報を読み取る
    size_t pos = cd_off;
    for (uint16_t i = 0; i < cd_count; ++i) {
        // CD エントリの最小サイズは 46 バイト
        if (pos + 46 > data.size()) throw ParseError("ZIP: central directory truncated");
        // CD エントリのシグネチャ確認: 0x02014B50 ("PK\x01\x02")
        if (rd32(data.data()+pos) != 0x02014B50) throw ParseError("ZIP: bad CD signature");

        // CD エントリから各フィールドを読み取る (ZIP 仕様のオフセット通り)
        uint16_t method    = rd16(data.data()+pos+10);   // 圧縮方式: 0=無圧縮, 8=DEFLATE
        uint32_t comp_sz   = rd32(data.data()+pos+20);   // 圧縮後サイズ
        uint32_t uncomp_sz = rd32(data.data()+pos+24);   // 展開後サイズ
        uint16_t fn_len    = rd16(data.data()+pos+28);   // ファイル名の長さ
        uint16_t ex_len    = rd16(data.data()+pos+30);   // 拡張フィールドの長さ
        uint16_t cm_len    = rd16(data.data()+pos+32);   // コメントの長さ
        uint32_t lhdr_off  = rd32(data.data()+pos+42);   // ローカルヘッダーへのオフセット

        // ファイル名がデータ範囲外でないか確認
        if (pos + 46 + fn_len > data.size()) throw ParseError("ZIP: filename beyond EOF");
        // ファイル名を文字列として取得 (UTF-8 のはずだが ZIP 規格上は任意)
        std::string fname(reinterpret_cast<const char*>(data.data()+pos+46), fn_len);
        // 次のエントリへ移動: 46 (固定) + fn_len + ex_len (拡張) + cm_len (コメント)
        pos += 46 + fn_len + ex_len + cm_len;

        // ローカルファイルヘッダーを参照して実際のデータ開始位置を計算する
        // ローカルヘッダーは +26 に「ファイル名長」, +28 に「拡張フィールド長」がある
        if (lhdr_off + 30 > data.size()) throw ParseError("ZIP: local header OOB");
        uint16_t lfn = rd16(data.data()+lhdr_off+26);   // ローカルヘッダーのファイル名長
        uint16_t lex = rd16(data.data()+lhdr_off+28);   // ローカルヘッダーの拡張フィールド長
        // データ開始位置: ローカルヘッダー先頭 + 30 (固定部) + lfn + lex
        size_t   doff = lhdr_off + 30 + lfn + lex;

        // データがファイル範囲内か確認
        if (doff + comp_sz > data.size()) throw ParseError("ZIP: file data OOB: " + fname);

        // ZipEntry を作成してデータを格納する
        ZipEntry entry;
        entry.path = fname;
        if (method == 0) {
            // 圧縮方式 0 = 無圧縮 (Stored): そのままコピー
            entry.data.assign(data.data()+doff, data.data()+doff+comp_sz);
        } else if (method == 8) {
            // 圧縮方式 8 = DEFLATE: 自前の inflate() で展開する
            entry.data = inflate(data.data()+doff, comp_sz, uncomp_sz);
        } else {
            // 未対応の圧縮方式 (bzip2 など) はエラー
            throw ParseError("ZIP: unsupported compression method " + std::to_string(method));
        }
        // ハッシュマップに格納 (std::move でコピーを避けて所有権を移動)
        entries_[fname] = std::move(entry);
    }
}

// コンストラクタ: バイト列を受け取って即座に parse() を呼ぶ
ZipReader::ZipReader(const std::vector<uint8_t>& data) { parse(data); }
// has(): 指定パスのエントリがハッシュマップに存在するか
bool ZipReader::has(const std::string& p) const { return entries_.count(p)>0; }
// get(): 指定パスのエントリを取得する。存在しなければ IOError
const ZipEntry& ZipReader::get(const std::string& p) const {
    auto it = entries_.find(p);
    if (it == entries_.end()) throw IOError("ZIP entry not found: " + p);
    return it->second;
}
// text(): データを const char* にキャストして std::string に変換して返す
std::string ZipReader::text(const std::string& p) const {
    auto& e = get(p);
    return {reinterpret_cast<const char*>(e.data.data()), e.data.size()};
}
// paths(): 全エントリのパスリストを返す (範囲 for で使いやすい)
std::vector<std::string> ZipReader::paths() const {
    std::vector<std::string> r;
    // 構造化束縛 [k,v] でキーと値を分解しながら反復
    for (auto& [k,v]:entries_) r.push_back(k);
    return r;
}

// ============================================================
//  リレーションシップ (.rels ファイル)
// ============================================================
// parse_rels(): .rels ファイルの XML を解析して Relationship のリストを返す
// .rels ファイルの形式:
//   <Relationships>
//     <Relationship Id="rId1" Type="...worksheet..." Target="worksheets/sheet1.xml"/>
//   </Relationships>
std::vector<Relationship> parse_rels(const std::string& xml) {
    std::vector<Relationship> out;
    size_t pos = 0;
    while (true) {
        // "<Relationship " タグの開始を探す
        auto r = xml.find("<Relationship ", pos);
        if (r == std::string::npos) break;   // もうなければ終了
        // タグの終端 ">" を探す
        auto re = xml.find('>', r);
        if (re == std::string::npos) break;
        // タグ文字列を切り出して属性を取得
        std::string elem = xml.substr(r, re-r+1);
        // Id, Type, Target 属性を取得して Relationship に格納
        out.push_back({xml_attr(elem,"Id"), xml_attr(elem,"Type"), xml_attr(elem,"Target")});
        pos = re+1;   // 次のタグを探す位置を更新
    }
    return out;
}

// ============================================================
//  XlsxSharedStrings (xl/sharedStrings.xml)
// ============================================================
// XlsxSharedStrings::parse(): sharedStrings.xml の XML をパースする
// 形式:
//   <sst>
//     <si><t>文字列1</t></si>
//     <si><r><t>パート1</t></r><r><t>パート2</t></r></si>  ← リッチテキスト
//   </sst>
// リッチテキスト (<r> タグ) の場合は各 <r> 内の <t> を連結して 1 文字列にする
void XlsxSharedStrings::parse(const std::string& xml) {
    size_t pos = 0;
    while (true) {
        // "<si>" (string item) タグの開始を探す
        auto si = xml.find("<si>", pos);
        if (si == std::string::npos) break;   // もうなければ終了
        // "</si>" の終端を探す
        auto sie = xml.find("</si>", si);
        if (sie == std::string::npos) break;
        // <si>...</si> のブロック全体を切り出す
        std::string block = xml.substr(si, sie-si+5);
        // 1 つの <si> が持つテキストを結合するバッファ
        std::string text_val;
        size_t tp = 0;
        while (true) {
            // "<t" タグの開始を探す (<t> または <t xml:space="preserve"> の両方に対応)
            auto ts = block.find("<t", tp);
            if (ts == std::string::npos) break;
            // "<t" タグの ">" を探す
            auto te = block.find('>', ts);
            if (te == std::string::npos) break;
            // "</t>" の終端を探す
            auto tc = block.find("</t>", te);
            if (tc == std::string::npos) break;
            // <t> と </t> の間のテキストをアンエスケープして結合
            text_val += xml_unescape(block.substr(te+1, tc-te-1));
            tp = tc + 4;   // "</t>" の次から検索を続ける
        }
        strings_.push_back(text_val);   // 結合した文字列をリストに追加
        pos = sie + 5;   // 次の <si> を探す
    }
}

// get(): インデックスに対応する文字列を返す
const std::string& XlsxSharedStrings::get(uint32_t idx) const {
    if (idx >= strings_.size())
        throw RangeError("SharedStrings index " + std::to_string(idx)
                         + " out of range (size=" + std::to_string(strings_.size()) + ")");
    return strings_[idx];
}

// intern(): 文字列を共有文字列テーブルに登録してインデックスを返す
// 同じ文字列が既に登録されていれば既存のインデックスを返す (重複排除)
uint32_t XlsxSharedStrings::intern(const std::string& s) {
    // index_map_ を使って O(1) で既存チェック
    auto it = index_map_.find(s);
    if (it != index_map_.end()) return it->second;   // 既存のインデックスを返す
    // 新規追加: strings_ の現在のサイズが新しいインデックスになる
    auto idx = static_cast<uint32_t>(strings_.size());
    strings_.push_back(s);
    index_map_[s] = idx;   // 逆引きマップにも登録
    return idx;
}

// ============================================================
//  XlsxStyles (xl/styles.xml)
// ============================================================
// XlsxStyles::parse(): styles.xml を解析して num_fmts_ と cell_xfs_ を埋める
// styles.xml の主要セクション:
//   <numFmts>  : カスタム数値フォーマット定義 (ID 164以降)
//   <cellXfs>  : セル書式レコード (XF) のリスト
void XlsxStyles::parse(const std::string& xml) {
    // ---- numFmts セクションのパース ----
    // <numFmts>...</numFmts> の範囲を見つける
    auto nf_start = xml.find("<numFmts");
    auto nf_end   = xml.find("</numFmts>");
    if (nf_start != std::string::npos && nf_end != std::string::npos) {
        std::string blk = xml.substr(nf_start, nf_end-nf_start);   // <numFmts>...</numFmts> 部分
        size_t pos = 0;
        while (true) {
            // <numFmt ... /> タグを探す
            auto p = blk.find("<numFmt ", pos);
            if (p == std::string::npos) break;
            auto pe = blk.find('>', p);
            if (pe == std::string::npos) break;
            std::string el = blk.substr(p, pe-p+1);
            std::string id_s  = xml_attr(el, "numFmtId");    // 数値フォーマット ID
            std::string fmt_s = xml_attr(el, "formatCode");  // フォーマット文字列
            // 修正: 空文字列や数値でない文字列に対する stoi を安全にガード
            if (!id_s.empty()) {
                try { num_fmts_[static_cast<uint16_t>(std::stoi(id_s))] = fmt_s; }
                catch (...) {}   // 不正な ID は無視
            }
            pos = pe+1;
        }
    }
    // ---- cellXfs セクションのパース ----
    // <cellXfs>...</cellXfs> の範囲を見つける
    auto xf_start = xml.find("<cellXfs");
    auto xf_end   = xml.find("</cellXfs>");
    if (xf_start != std::string::npos && xf_end != std::string::npos) {
        std::string blk = xml.substr(xf_start, xf_end-xf_start);   // <cellXfs>...</cellXfs> 部分
        size_t pos = 0;
        while (true) {
            // <xf ... > タグを探す (セル書式レコード)
            auto p = blk.find("<xf ", pos);
            if (p == std::string::npos) break;
            auto pe = blk.find('>', p);
            if (pe == std::string::npos) break;
            std::string el = blk.substr(p, pe-p+1);
            XlsxXF xf;   // 新しい XF レコード
            std::string nfid = xml_attr(el, "numFmtId");   // 数値フォーマット ID
            if (!nfid.empty()) {
                try { xf.num_fmt_id = static_cast<uint16_t>(std::stoi(nfid)); } catch (...) {}
            }
            // applyNumberFormat="1" または "true" なら数値フォーマット適用フラグを立てる
            std::string apnf = xml_attr(el, "applyNumberFormat");
            xf.apply_number_format = (apnf == "1" || apnf == "true");
            cell_xfs_.push_back(xf);   // リストに追加
            pos = pe+1;
        }
    }
}

// cell_format(): XF インデックスに対応する FormatInfo を返す
// FormatInfo は format_index, format_string, is_date_format の 3 フィールドを持つ
FormatInfo XlsxStyles::cell_format(uint32_t xf_idx) const {
    FormatInfo fi;
    if (xf_idx >= cell_xfs_.size()) return fi;   // 範囲外なら空の FormatInfo を返す
    auto& xf = cell_xfs_[xf_idx];
    fi.format_index   = xf.num_fmt_id;                   // 数値フォーマット ID をコピー
    auto it = num_fmts_.find(xf.num_fmt_id);
    if (it != num_fmts_.end()) fi.format_string = it->second;   // カスタムフォーマット文字列
    fi.is_date_format = is_date_format(xf.num_fmt_id, fi.format_string);   // 日付判定
    return fi;
}

// is_date_format(): ID とフォーマット文字列から日付書式か判定する
// 組み込み日付フォーマット ID の範囲: 14〜22 および 45〜47
bool XlsxStyles::is_date_format(uint16_t id, const std::string& fmt) {
    if ((id >= 14 && id <= 22) || (id >= 45 && id <= 47)) return true;   // 組み込み日付範囲
    return looks_like_date(fmt);   // フォーマット文字列から推測
}
// looks_like_date(): フォーマット文字列に 'y' (年) または 'd' (日) が含まれるか調べる
// 注意: 'd' は数字の桁数指定にも使われるため、'h' (時) がない場合だけ日付と判断する
bool XlsxStyles::looks_like_date(const std::string& fmt) {
    std::string lo = fmt;
    // フォーマット文字列を小文字に変換 (Y/y, D/d 両方に対応するため)
    std::transform(lo.begin(), lo.end(), lo.begin(), ::tolower);
    return lo.find('y') != std::string::npos          // 'y' があれば年表示 → 日付
        || (lo.find('d') != std::string::npos && lo.find('h') == std::string::npos);
        // 'd' があり 'h' がなければ日付 ('h' があれば数値の桁書式の可能性)
}

// ============================================================
//  XlsxSheet の実装
// ============================================================
// put_cell(): パーサーが解析済みセルをシートデータに直接格納する
// data_ は 行番号 → (列番号 → Cell) の二重マップ
void XlsxSheet::put_cell(const Cell& c) {
    data_[c.address.row][c.address.col] = c;   // 二重マップに格納
    // 最大行数・列数を更新する (0-indexed なので +1)
    if (c.address.row + 1 > row_count_) row_count_ = c.address.row + 1;
    if (c.address.col + 1 > col_count_) col_count_ = c.address.col + 1;
}
// set_dimensions(): 行数・列数を直接設定する
void XlsxSheet::set_dimensions(uint32_t r, uint32_t c) { row_count_=r; col_count_=c; }

// blank_cell(): 指定アドレスの空セルを生成するヘルパー
// cell() が空セルを返す必要があるときに使う
static Cell blank_cell(uint32_t r, uint32_t c) {
    Cell cell; cell.address={r,c}; cell.type=CellType::Blank; cell.value=BlankValue{}; return cell;
}

// cell(): 行・列インデックスでセルを取得する
// データが存在すれば返し、存在しなければ blank_cell() を返す
Cell XlsxSheet::cell(uint32_t r, uint32_t c) const {
    // まず行を検索
    auto ri = data_.find(r);
    if (ri != data_.end()) {
        // 行が存在すれば列を検索
        auto ci = ri->second.find(c);
        if (ci != ri->second.end()) return ci->second;   // セルが存在すれば返す
    }
    return blank_cell(r,c);   // データがなければ空セルを返す
}
// cell(): CellAddress 版。row と col に分解して上の cell() に委譲
Cell XlsxSheet::cell(const CellAddress& a) const { return cell(a.row,a.col); }
// cell(): A1 記法版。CellAddress に変換して上の cell() に委譲
Cell XlsxSheet::cell(const std::string& a1) const { return cell(CellAddress::from_a1(a1)); }

// try_cell(): セルが存在して空でない場合のみ返す安全版
// 空セルや存在しないセルの場合は std::nullopt を返す
std::optional<Cell> XlsxSheet::try_cell(uint32_t r, uint32_t c) const {
    auto ri = data_.find(r);
    if (ri == data_.end()) return std::nullopt;   // 行がなければ nullopt
    auto ci = ri->second.find(c);
    // 列がない、またはセルが空の場合は nullopt
    if (ci == ri->second.end() || ci->second.is_blank()) return std::nullopt;
    return ci->second;   // セルが存在して非空なら値を返す
}
// row(): 指定行の全セルを列番号順で返す
std::vector<Cell> XlsxSheet::row(uint32_t r) const {
    std::vector<Cell> out;
    auto ri = data_.find(r);
    // 行が存在すれば、その行の全セルを out に追加する
    if (ri != data_.end()) for (auto& [c,cell]:ri->second) out.push_back(cell);
    return out;
}
// cells(): シート全体の非空セルを返す
std::vector<Cell> XlsxSheet::cells() const {
    std::vector<Cell> out;
    // 全行・全列を走査して空でないセルのみ収集
    for (auto& [r,cols]:data_) for (auto& [c,cell]:cols) if (!cell.is_blank()) out.push_back(cell);
    return out;
}
// for_each_cell(): 全セルを fn に渡して走査する
void XlsxSheet::for_each_cell(std::function<void(const Cell&)> fn) const {
    for (auto& [r,cols]:data_) for (auto& [c,cell]:cols) fn(cell);
}

// fill_type(): セル値の variant 型 (CellValue) から CellType を推定して設定する
// CellValue は std::variant<BlankValue, bool, int64_t, double, string, ErrorValue>
static void fill_type(Cell& c) {
    // std::visit でラムダを variant の実際の型に合わせて呼び出す
    std::visit([&](auto&& v){
        using T=std::decay_t<decltype(v)>;   // 実際の型を T に束縛
        if      constexpr(std::is_same_v<T,BlankValue>)   c.type=CellType::Blank;
        else if constexpr(std::is_same_v<T,bool>)         c.type=CellType::Boolean;
        else if constexpr(std::is_same_v<T,int64_t>)      c.type=CellType::Number;
        else if constexpr(std::is_same_v<T,double>)       c.type=CellType::Number;
        else if constexpr(std::is_same_v<T,std::string>)  c.type=CellType::String;
        else if constexpr(std::is_same_v<T,ErrorValue>)   c.type=CellType::Error;
    }, c.value);
}
// set_cell(): 行・列インデックスでセルに値を書き込む
void XlsxSheet::set_cell(uint32_t r, uint32_t c, const CellValue& v) {
    Cell cell; cell.address={r,c}; cell.value=v; fill_type(cell); put_cell(cell);
}
// set_cell(): A1 記法でセルに値を書き込む
void XlsxSheet::set_cell(const std::string& a1, const CellValue& v) {
    auto a=CellAddress::from_a1(a1); set_cell(a.row,a.col,v);
}
// set_formula(): セルに数式を設定する
void XlsxSheet::set_formula(const std::string& a1, const std::string& f) {
    auto a=CellAddress::from_a1(a1);
    // formula フィールドに数式文字列を設定し、value は空 (計算前) にする
    Cell c; c.address=a; c.type=CellType::Formula; c.formula=f; c.value=BlankValue{}; put_cell(c);
}

// ============================================================
//  XlsxWorkbook の実装
// ============================================================
// add_parsed_sheet(): パーサーが解析済みシートを追加する内部 API
// std::move で unique_ptr の所有権を sheets_ に移動する
void XlsxWorkbook::add_parsed_sheet(std::unique_ptr<XlsxSheet> s) { sheets_.push_back(std::move(s)); }
// sheet(): インデックスでシートを取得 (非 const 版)
Sheet& XlsxWorkbook::sheet(size_t i) {
    if(i>=sheets_.size()) throw RangeError("Sheet index "+std::to_string(i)+" out of range");
    return *sheets_[i];   // unique_ptr をデリファレンスして参照を返す
}
// sheet(): インデックスでシートを取得 (const 版)
const Sheet& XlsxWorkbook::sheet(size_t i) const {
    if(i>=sheets_.size()) throw RangeError("Sheet index "+std::to_string(i)+" out of range");
    return *sheets_[i];
}
// sheet(): シート名でシートを取得
Sheet& XlsxWorkbook::sheet(const std::string& name) {
    for(auto& s:sheets_) if(s->name()==name) return *s;   // 名前が一致するシートを探す
    throw RangeError("Sheet not found: "+name);
}
// sheet_names(): 全シート名をリストで返す
std::vector<std::string> XlsxWorkbook::sheet_names() const {
    std::vector<std::string> n; for(auto& s:sheets_) n.push_back(s->name()); return n;
}
// add_sheet(): 新しいシートを末尾に追加する
Sheet& XlsxWorkbook::add_sheet(const std::string& name) {
    // 同名シートが存在すれば WriteError を投げる
    for(auto& s:sheets_) if(s->name()==name) throw WriteError("Sheet exists: "+name);
    sheets_.push_back(std::make_unique<XlsxSheet>(name));   // 新しい XlsxSheet を生成
    return *sheets_.back();   // 末尾要素の参照を返す
}
// remove_sheet(): インデックス指定でシートを削除する
void XlsxWorkbook::remove_sheet(size_t i) {
    if(i>=sheets_.size()) throw RangeError("Sheet index out of range");
    // vector の erase: イテレータで指定する。ptrdiff_t にキャストして符号なし整数の問題を避ける
    sheets_.erase(sheets_.begin()+static_cast<ptrdiff_t>(i));
}
// rename_sheet(): シート名を変更する
void XlsxWorkbook::rename_sheet(size_t i, const std::string& name) {
    if(i>=sheets_.size()) throw RangeError("Sheet index out of range");
    // 同名シートが別に存在すれば WriteError
    for(size_t j=0;j<sheets_.size();++j)
        if(j!=i && sheets_[j]->name()==name) throw WriteError("Sheet name exists: "+name);
    // 古いシートから新しいシートへデータをコピーする
    auto old = std::move(sheets_[i]);             // 古いシートの所有権を取得
    auto ns  = std::make_unique<XlsxSheet>(name); // 新しい名前でシートを作成
    // for_each_cell で全セルをコピー
    old->for_each_cell([&](const Cell& c){ ns->put_cell(c); });
    // 行数・列数もコピー
    ns->set_dimensions(old->row_count(), old->col_count());
    sheets_[i] = std::move(ns);   // 新しいシートで置き換え
}
// save(): ファイルパスに XLSX として保存する
void XlsxWorkbook::save(const std::string& path, const SaveOptions& opts) const {
    auto bytes = to_bytes(FileFormat::XLSX, opts);   // バイト列を生成
    std::ofstream f(path, std::ios::binary);          // バイナリモードで開く
    if(!f) throw IOError("Cannot write: "+path);
    // バイト列をファイルに書き込む
    f.write(reinterpret_cast<const char*>(bytes.data()),
            static_cast<std::streamsize>(bytes.size()));
}
// to_bytes(): XLSX バイト列 (ZIP アーカイブ) を生成して返す
std::vector<uint8_t> XlsxWorkbook::to_bytes(FileFormat, const SaveOptions&) const {
    XlsxWriter w; return w.write(*this, {});   // XlsxWriter に処理を委譲
}

// ============================================================
//  XlsxParser の実装
// ============================================================
// parse_ref(): "B3" のような A1 記法を CellAddress (0-indexed) に変換する
// CellAddress::from_a1() に委譲するだけのラッパー
CellAddress XlsxParser::parse_ref(const std::string& ref) {
    return CellAddress::from_a1(ref);
}

// parse(): XLSX バイト列全体をパースして XlsxWorkbook を返す
// 処理の流れ:
//   1. ZipReader で ZIP を展開して各 XML ファイルを取り出す
//   2. sharedStrings.xml をパース → 共有文字列テーブルを構築
//   3. styles.xml をパース → 書式テーブルを構築
//   4. workbook.xml + .rels をパース → シートメタ情報 (名前とパス) を収集
//   5. 各シートの XML をパース → XlsxSheet を生成して XlsxWorkbook に追加
std::unique_ptr<XlsxWorkbook> XlsxParser::parse(const std::vector<uint8_t>& data) {
    ZipReader zip(data);   // ZIP を展開して全エントリをマップに格納

    // 共有文字列テーブルのパース (存在しないファイルはスキップ)
    XlsxSharedStrings sst;
    if (zip.has("xl/sharedStrings.xml")) sst.parse(zip.text("xl/sharedStrings.xml"));

    // スタイルテーブルのパース
    XlsxStyles styles;
    if (zip.has("xl/styles.xml")) styles.parse(zip.text("xl/styles.xml"));

    // workbook.xml は必須ファイル (なければ XLSX として壊れている)
    if (!zip.has("xl/workbook.xml")) throw ParseError("XLSX: missing xl/workbook.xml");

    // リレーションシップのパース (workbook.xml → 各シート XML のパスを解決するため)
    std::vector<Relationship> rels;
    if (zip.has("xl/_rels/workbook.xml.rels"))
        rels = parse_rels(zip.text("xl/_rels/workbook.xml.rels"));

    // workbook.xml からシートメタ情報 (名前, rId, パス) を収集
    auto metas = parse_workbook_xml(zip.text("xl/workbook.xml"), rels);
    auto wb = std::make_unique<XlsxWorkbook>();   // 空の XlsxWorkbook を作成

    // 各シートをパースして XlsxWorkbook に追加
    for (auto& meta : metas) {
        // パスの形式を正規化: OOXML のパスは絶対/相対/ベア名前の 3 種類がある
        std::string path;
        if (!meta.path.empty() && meta.path[0] == '/') {
            path = "xl" + meta.path;                   // 絶対パス: "/worksheets/sheet1.xml" → "xl/worksheets/sheet1.xml"
        } else if (meta.path.find('/') != std::string::npos) {
            path = "xl/" + meta.path;                  // 相対パス: "worksheets/sheet1.xml" → "xl/worksheets/sheet1.xml"
        } else {
            path = "xl/worksheets/" + meta.path;       // ベア名前: "sheet1.xml" → "xl/worksheets/sheet1.xml"
        }

        // ZIP 内に実際にファイルが存在するか確認
        if (!zip.has(path)) throw ParseError("XLSX: sheet not found in ZIP: " + path);
        // シート XML をパースして XlsxSheet を生成し、ワークブックに追加
        wb->add_parsed_sheet(parse_sheet_xml(meta.name, zip.text(path), sst, styles));
    }
    return wb;
}

// parse_workbook_xml(): workbook.xml を解析してシートのメタ情報リストを返す
// workbook.xml の <sheets> セクション例:
//   <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
std::vector<XlsxParser::SheetMeta> XlsxParser::parse_workbook_xml(
    const std::string& xml, const std::vector<Relationship>& rels)
{
    std::vector<SheetMeta> out;
    size_t pos = 0;
    while (true) {
        // "<sheet " タグを探す
        auto sh = xml.find("<sheet ", pos);
        if (sh == std::string::npos) break;
        auto she = xml.find('>', sh);
        if (she == std::string::npos) break;
        std::string el = xml.substr(sh, she-sh+1);   // <sheet ... /> タグ全体
        SheetMeta m;
        m.name = xml_unescape(xml_attr(el, "name"));   // シート名 (XML エスケープを戻す)
        m.rid  = xml_attr(el, "r:id");                 // リレーション ID (例: "rId1")
        // rels からリレーション ID に対応するターゲットパスを探す
        for (auto& r : rels) if (r.id == m.rid) { m.path = r.target; break; }
        if (!m.path.empty()) out.push_back(m);   // パスが解決できたシートだけ追加
        pos = she+1;
    }
    return out;
}

// parse_sheet_xml(): シートの XML をパースして XlsxSheet を返す
// 処理の流れ:
//   1. <row> タグを見つけて行ブロックを切り出す
//   2. 行ブロック内の <c> タグ (セル) を処理する
//   3. 各セルの属性 (r, t, s) とコンテンツ (v, f) を取得して parse_cell_element() に渡す
std::unique_ptr<XlsxSheet> XlsxParser::parse_sheet_xml(
    const std::string& name, const std::string& xml,
    const XlsxSharedStrings& sst, const XlsxStyles& styles)
{
    auto sheet = std::make_unique<XlsxSheet>(name);   // 名前付きシートを作成
    size_t pos = 0;
    while (true) {
        // "<row " タグで行の開始を探す
        auto rs = xml.find("<row ", pos);
        if (rs == std::string::npos) break;
        // "</row>" で行の終了を探す
        auto re = xml.find("</row>", rs);
        if (re == std::string::npos) break;
        // 行ブロック (<row ...>...</row>) を切り出す
        std::string row_blk = xml.substr(rs, re-rs);

        // 行ブロック内のセル (<c> タグ) を処理する
        size_t cp = 0;
        while (true) {
            // "<c " タグでセルの開始を探す
            auto cs = row_blk.find("<c ", cp);
            if (cs == std::string::npos) break;
            // ">" でタグの終端を探す
            auto tag_end = row_blk.find('>', cs);
            if (tag_end == std::string::npos) break;
            // 自己終了タグ (<c ... />) かどうかを確認 (">" の直前が "/" なら自己終了)
            bool self_close = tag_end > 0 && row_blk[tag_end-1] == '/';

            std::string cell_xml;   // セルの XML 文字列
            size_t next_cp;          // 次の検索開始位置
            if (self_close) {
                // 自己終了タグ: <c r="A1" t="n"/>  → <v> も <f> もない
                cell_xml = row_blk.substr(cs, tag_end-cs+1);
                next_cp  = tag_end+1;
            } else {
                // 通常タグ: </c> を探して終端を決める
                auto ce = row_blk.find("</c>", tag_end);
                if (ce == std::string::npos) { cp = tag_end+1; break; }   // </c> がなければスキップ
                cell_xml = row_blk.substr(cs, ce-cs+4);   // <c...>...</c> 全体
                next_cp  = ce+4;
            }

            // 開始タグ部分 (<c ... >) を切り出して属性を取得する
            std::string tag = cell_xml.substr(0, cell_xml.find('>')+1);
            std::string ref  = xml_attr(tag, "r");    // セル参照 (例: "B3")
            std::string type = xml_attr(tag, "t");    // セルタイプ ("s"=共有文字列, "b"=bool, "e"=エラー, ""=数値)
            std::string s    = xml_attr(tag, "s");    // スタイルインデックス (XF ID)
            // 自己終了タグのときは <v> や <f> が存在しないので空文字列
            std::string v    = self_close ? "" : xml_text(cell_xml, "v");   // <v> の値
            std::string f    = self_close ? "" : xml_text(cell_xml, "f");   // <f> の数式

            // ref が空でなければセルとして処理する
            if (!ref.empty()) {
                try {
                    // セルデータを解析して Cell を生成し、シートに格納する
                    Cell c = parse_cell_element(ref, type, s, v, f, sst, styles);
                    sheet->put_cell(c);
                } catch (const FormatError&) {
                    // strict_validation が false (デフォルト) なら壊れたセルはスキップ
                    if (opts_.strict_validation) throw;
                    // 非 strict モード: パース失敗セルを無視して続行
                }
            }
            cp = next_cp;   // 次のセルへ
        }
        pos = re + 6;   // "</row>" の後から次の行を探す
    }
    return sheet;
}

// parse_cell_element(): <c> 要素のデータから Cell オブジェクトを生成する
// type 属性の値でセルの種類を判別する:
//   "s"         → 共有文字列インデックス (v_text をインデックスとして sst から取得)
//   "str"       → インライン文字列 (数式の結果が文字列の場合)
//   "inlineStr" → インライン文字列 (型省略を避けるための表記)
//   "b"         → bool (v_text が "1" なら true)
//   "e"         → エラー値 (#DIV/0! など)
//   ""          → 数値 (小数点や指数がなければ int64_t, あれば double)
Cell XlsxParser::parse_cell_element(
    const std::string& ref, const std::string& type,
    const std::string& s_attr, const std::string& v_text,
    const std::string& f_text,
    const XlsxSharedStrings& sst, const XlsxStyles& styles)
{
    Cell c;
    c.address = parse_ref(ref);   // "B3" → {row=2, col=1} に変換

    // スタイル (XF インデックス) を設定する
    if (!s_attr.empty()) {
        try {
            uint32_t xf = static_cast<uint32_t>(std::stoul(s_attr));
            CellStyle cs; cs.format = styles.cell_format(xf); c.style = cs;
        } catch (...) {}   // 不正な s 属性は無視
    }
    // 数式があれば formula フィールドに格納 (XML エスケープを戻す)
    if (!f_text.empty()) c.formula = xml_unescape(f_text);

    if (type == "s") {
        // 共有文字列型: v_text はテーブルのインデックス
        c.type = CellType::String;
        if (!v_text.empty()) {
            try { c.value = sst.get(static_cast<uint32_t>(std::stoul(v_text))); }
            catch (...) { c.value = std::string(""); }   // 不正なインデックスは空文字列
        }
    } else if (type == "str" || type == "inlineStr") {
        // インライン文字列型: v_text が文字列そのもの
        c.type = CellType::String;
        c.value = xml_unescape(v_text);
    } else if (type == "b") {
        // bool 型: "1" または "true" なら true
        c.type = CellType::Boolean;
        c.value = (v_text == "1" || v_text == "true");
    } else if (type == "e") {
        // エラー型: v_text がエラーコード ("#DIV/0!" など)
        c.type = CellType::Error;
        c.value = ErrorValue{xml_unescape(v_text)};
    } else {
        // 数値型 (デフォルト): type 属性が空の場合
        if (v_text.empty()) {
            // 値がなければ空セル
            c.type = CellType::Blank; c.value = BlankValue{};
        } else {
            c.type = CellType::Number;
            try {
                // 小数点 '.' や指数 'e'/'E' がなければ整数として解釈する
                // 例: "42" → int64_t{42}、"3.14" → double{3.14}
                if (v_text.find('.') == std::string::npos &&
                    v_text.find('e') == std::string::npos &&
                    v_text.find('E') == std::string::npos) {
                    c.value = static_cast<int64_t>(std::stoll(v_text));   // 整数
                } else {
                    c.value = std::stod(v_text);   // 浮動小数点数
                }
            } catch (...) {
                // 数値変換失敗 (不正な文字列) → 文字列として格納
                c.type = CellType::String; c.value = v_text;
            }
        }
    }
    // 数式があり、かつエラーでなければ型を Formula に上書きする
    if (!f_text.empty() && c.type != CellType::Error) c.type = CellType::Formula;
    return c;
}

// ============================================================
//  XlsxWriter — 最小限の ZIP ビルダー
// ============================================================

// ap16(): uint16_t を Little Endian でバイト列に追記するヘルパー
// 例: ap16(v, 0x0102) → v に {0x02, 0x01} を追加
static void ap16(std::vector<uint8_t>& v, uint16_t x) { v.push_back(x&0xFF); v.push_back((x>>8)&0xFF); }
// ap32(): uint32_t を Little Endian でバイト列に追記するヘルパー
static void ap32(std::vector<uint8_t>& v, uint32_t x) {
    v.push_back(x&0xFF); v.push_back((x>>8)&0xFF);
    v.push_back((x>>16)&0xFF); v.push_back((x>>24)&0xFF);
}

// ZFile: ZIP に格納する 1 ファイルを表す構造体
// name: ファイル名 (ZIP 内パス), data: ファイルデータ, off: ローカルヘッダーオフセット, crc: CRC32
struct ZFile { std::string name; std::vector<uint8_t> data; uint32_t off{0}, crc{0}; };

// build_zip(): ZFile のリストから ZIP アーカイブのバイト列を生成する
// ZIP 構造:
//   [ローカルファイルヘッダー + データ] × n  (無圧縮: method=0)
//   [セントラルディレクトリ]
//   [EOCD]
static std::vector<uint8_t> build_zip(std::vector<ZFile>& files) {
    std::vector<uint8_t> out;
    // ---- ローカルファイルヘッダー + データを追加 ----
    for (auto& f : files) {
        f.off = uint32_t(out.size());   // このファイルのローカルヘッダー開始位置を記録
        // CRC32 を計算 (ZIP はデータの整合性チェックに使う)
        f.crc = excellib::detail::crc32_compute(f.data.data(), f.data.size());
        // ローカルファイルヘッダー (各フィールドは ZIP 仕様通り)
        ap32(out,0x04034B50);                         // シグネチャ "PK\x03\x04"
        ap16(out,20);                                 // バージョン必要: 2.0 (= 20)
        ap16(out,0);                                  // 汎用フラグ: 0
        ap16(out,0);                                  // 圧縮方式: 0 = 無圧縮 (Stored)
        ap16(out,0); ap16(out,0);                     // 最終更新時刻・日付: 0 (未設定)
        ap32(out,f.crc);                              // CRC32
        ap32(out,uint32_t(f.data.size()));            // 圧縮サイズ (無圧縮なので = 元サイズ)
        ap32(out,uint32_t(f.data.size()));            // 展開後サイズ
        ap16(out,uint16_t(f.name.size()));            // ファイル名長
        ap16(out,0);                                  // 拡張フィールド長: 0
        // ファイル名を追記
        out.insert(out.end(),f.name.begin(),f.name.end());
        // ファイルデータを追記
        out.insert(out.end(),f.data.begin(),f.data.end());
    }
    // ---- セントラルディレクトリを追加 ----
    uint32_t cd_off = uint32_t(out.size());   // CD の開始位置を記録
    for (auto& f : files) {
        // CD エントリ (各フィールドは ZIP 仕様通り)
        ap32(out,0x02014B50);                         // シグネチャ "PK\x01\x02"
        ap16(out,20); ap16(out,20);                   // バージョン作成/必要: 2.0
        ap16(out,0); ap16(out,0); ap16(out,0);        // フラグ/圧縮/時刻: 0
        ap32(out,f.crc);                              // CRC32
        ap32(out,uint32_t(f.data.size()));            // 圧縮サイズ
        ap32(out,uint32_t(f.data.size()));            // 展開後サイズ
        ap16(out,uint16_t(f.name.size()));            // ファイル名長
        ap16(out,0); ap16(out,0);                     // 拡張/コメント長: 0
        ap16(out,0); ap16(out,0);                     // ディスク開始番号/内部属性: 0
        ap32(out,0);                                  // 外部属性: 0
        ap32(out,f.off);                              // ローカルヘッダーオフセット
        out.insert(out.end(),f.name.begin(),f.name.end());   // ファイル名
    }
    // ---- EOCD (End Of Central Directory) を追加 ----
    uint32_t cd_sz = uint32_t(out.size()) - cd_off;   // CD のサイズを計算
    ap32(out,0x06054B50);                             // EOCD シグネチャ "PK\x05\x06"
    ap16(out,0); ap16(out,0);                         // ディスク番号/CD ディスク番号: 0
    ap16(out,uint16_t(files.size()));                 // このディスクの CD エントリ数
    ap16(out,uint16_t(files.size()));                 // CD エントリ総数
    ap32(out,cd_sz);                                  // CD サイズ
    ap32(out,cd_off);                                 // CD オフセット
    ap16(out,0);                                      // コメント長: 0
    return out;
}

// esc(): XML の特殊文字をエスケープするヘルパー (&→&amp; など)
static std::string esc(const std::string& s) {
    std::string r;
    for (char c:s) {
        switch(c){
            case '&': r+="&amp;"; break;   // & → &amp;
            case '<': r+="&lt;";  break;   // < → &lt;
            case '>': r+="&gt;";  break;   // > → &gt;
            case '"': r+="&quot;"; break;  // " → &quot;
            default:  r+=c;                // それ以外はそのまま
        }
    }
    return r;
}

// sb(): std::string をバイト列 (std::vector<uint8_t>) に変換するヘルパー
// "string bytes" の略
static std::vector<uint8_t> sb(const std::string& s) {
    return {s.begin(), s.end()};
}

// sheet_xml(): XlsxSheet の内容を XML 文字列に変換する
// 文字列セルは sst.intern() で共有文字列テーブルに登録してインデックスを使う
std::string XlsxWriter::sheet_xml(const XlsxSheet& sh, XlsxSharedStrings& sst) {
    std::ostringstream o;
    // XML ヘッダーとルート要素
    o << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
      << "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
      << "<sheetData>";
    // 全行を走査 (0-indexed)
    for (uint32_t r = 0; r < sh.row_count(); ++r) {
        auto cells = sh.row(r);
        if (cells.empty()) continue;   // 空行はスキップ
        o << "<row r=\"" << r+1 << "\">";   // <row r="1"> (1-indexed)
        for (auto& c : cells) {
            // セルアドレスを A1 記法に変換
            std::string ref = CellAddress{c.address.row,c.address.col}.to_a1();
            if (c.has_formula()) {
                // 数式セル: <f> タグに数式を格納
                o << "<c r=\"" << ref << "\"><f>" << esc(*c.formula) << "</f></c>";
                continue;
            }
            // 値の型に応じて異なる XML を出力 (std::visit で variant を分岐)
            std::visit([&](auto&& v){
                using T=std::decay_t<decltype(v)>;
                if constexpr(std::is_same_v<T,BlankValue>) {}   // 空値は出力しない
                else if constexpr(std::is_same_v<T,bool>)
                    // bool: t="b" 属性を付けて 1/0 を出力
                    o << "<c r=\""<<ref<<"\" t=\"b\"><v>"<<(v?1:0)<<"</v></c>";
                else if constexpr(std::is_same_v<T,int64_t>)
                    // 整数: 型属性なし (デフォルトは数値)
                    o << "<c r=\""<<ref<<"\"><v>"<<v<<"</v></c>";
                else if constexpr(std::is_same_v<T,double>)
                    // 浮動小数点: 型属性なし
                    o << "<c r=\""<<ref<<"\"><v>"<<v<<"</v></c>";
                else if constexpr(std::is_same_v<T,std::string>)
                    // 文字列: t="s" 属性を付けてインデックスを出力
                    o << "<c r=\""<<ref<<"\" t=\"s\"><v>"<<sst.intern(v)<<"</v></c>";
                else if constexpr(std::is_same_v<T,ErrorValue>)
                    // エラー: t="e" 属性を付けてエラーコードを出力
                    o << "<c r=\""<<ref<<"\" t=\"e\"><v>"<<esc(v.code)<<"</v></c>";
            }, c.value);
        }
        o << "</row>";
    }
    o << "</sheetData></worksheet>";
    return o.str();
}

// write(): XlsxWorkbook を XLSX バイト列 (ZIP) として書き出す
// 生成する XML ファイル一覧:
//   xl/worksheets/sheet{i+1}.xml  → 各シートの XML
//   xl/sharedStrings.xml          → 共有文字列テーブル (文字列がある場合のみ)
//   xl/styles.xml                 → スタイル定義 (最小限)
//   xl/workbook.xml               → ワークブック (シートリスト)
//   xl/_rels/workbook.xml.rels    → リレーションシップ
//   [Content_Types].xml           → コンテンツタイプ定義
//   _rels/.rels                   → ルートリレーションシップ
std::vector<uint8_t> XlsxWriter::write(const XlsxWorkbook& wb, const SaveOptions&) {
    XlsxSharedStrings sst;   // 書き込み用共有文字列テーブル (空から始める)
    std::vector<ZFile> files;  // ZIP に格納するファイルのリスト
    auto names = wb.sheet_names();

    // ---- 各シートの XML を生成 ----
    for (size_t i = 0; i < wb.sheet_count(); ++i) {
        // XlsxSheet 型にダウンキャストして sheet_xml() を呼ぶ
        const auto& sh = dynamic_cast<const XlsxSheet&>(wb.sheet(i));
        ZFile f;
        f.name = "xl/worksheets/sheet" + std::to_string(i+1) + ".xml";
        f.data = sb(sheet_xml(sh, sst));   // sheet_xml() で XML を生成してバイト列に変換
        files.push_back(std::move(f));
    }

    // ---- 共有文字列テーブルの XML を生成 (文字列セルがある場合のみ) ----
    bool has_ss = sst.size() > 0;   // 文字列セルが 1 つでもあれば true
    if (has_ss) {
        std::ostringstream o;
        // <sst> ルート要素 (count: 総参照数, uniqueCount: ユニーク数)
        o << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
          << "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
          << " count=\""<<sst.size()<<"\" uniqueCount=\""<<sst.size()<<"\">";
        // 各文字列を <si><t>...</t></si> 形式で出力
        for (size_t i=0;i<sst.size();++i)
            o<<"<si><t xml:space=\"preserve\">"<<esc(sst.get(i))<<"</t></si>";
        o << "</sst>";
        files.push_back({"xl/sharedStrings.xml", sb(o.str())});
    }

    // ---- 最小限のスタイル XML を生成 ----
    // セルスタイルを使わない場合でも styles.xml は必須ファイル
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

    // ---- workbook.xml を生成 ----
    {
        std::ostringstream o;
        o << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
          << "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
          // r: 名前空間 = リレーションシップ ID の参照に使う
          << " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
          << "<sheets>";
        for (size_t i=0;i<names.size();++i)
            // <sheet name="..." sheetId="..." r:id="rId..."/>
            o<<"<sheet name=\""<<esc(names[i])<<"\" sheetId=\""<<i+1
             <<"\" r:id=\"rId"<<i+1<<"\"/>";
        o << "</sheets></workbook>";
        files.push_back({"xl/workbook.xml", sb(o.str())});
    }

    // ---- xl/_rels/workbook.xml.rels を生成 ----
    {
        std::ostringstream o;
        o << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
          << "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
          // スタイルシートへのリレーション (rId0)
          << "<Relationship Id=\"rId0\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>";
        // 各シートへのリレーション (rId1, rId2, ...)
        for (size_t i=0;i<wb.sheet_count();++i)
            o<<"<Relationship Id=\"rId"<<i+1<<"\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet"<<i+1<<".xml\"/>";
        // 共有文字列テーブルへのリレーション (文字列があるときのみ)
        if (has_ss)
            o<<"<Relationship Id=\"rIdSS\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>";
        o << "</Relationships>";
        files.push_back({"xl/_rels/workbook.xml.rels", sb(o.str())});
    }

    // ---- [Content_Types].xml を生成 ----
    // ZIP 内の各ファイルの MIME タイプを定義するマニフェストファイル
    {
        std::ostringstream o;
        o << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
          << "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
          // デフォルト: .rels ファイルと .xml ファイルのタイプ
          << "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
          << "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
          // ワークブック本体の MIME タイプ
          << "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
          // スタイルの MIME タイプ
          << "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>";
        // 各シートの MIME タイプ
        for (size_t i=0;i<wb.sheet_count();++i)
            o<<"<Override PartName=\"/xl/worksheets/sheet"<<i+1<<".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>";
        // 共有文字列テーブルの MIME タイプ
        if (has_ss)
            o<<"<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>";
        o << "</Types>";
        files.push_back({"[Content_Types].xml", sb(o.str())});
    }

    // ---- _rels/.rels を生成 ----
    // パッケージ全体のルートリレーションシップ (xl/workbook.xml を指す)
    files.push_back({"_rels/.rels", sb(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
        "</Relationships>"
    )});

    // 全ファイルを ZIP にまとめて返す
    return build_zip(files);
}

} // namespace excellib::xlsx
