// xlsx_page_setup.cpp
// XLSX ファイル (ZIP + XML) にページ設定 (印刷設定) を直接書き込む実装
// Excel COM を使わずに XLSX の内部 XML を直接編集することで、
// Windows 以外の環境でもページ設定をファイルに適用できる
#include "excellib/excel_printer.hpp"
// xlsx_parser.hpp: ZipReader (ZIP 読み書き) と XlsxSharedStrings などの定義
#include "excellib/xlsx_parser.hpp"
// deflate.hpp: CRC32 計算と DEFLATE 展開 (ZIP の読み書きに使う)
#include "../common/deflate.hpp"
// std::vector, std::string などの標準コンテナ
#include <vector>
#include <string>
// std::ostringstream: 文字列を組み立てるストリーム
#include <sstream>
// std::ofstream / std::ifstream: ファイルの読み書き
#include <fstream>
// std::memcmp: バイト列比較
#include <cstring>
// std::transform: 文字列変換 (大文字/小文字など)
#include <algorithm>
// std::runtime_error などの例外クラス
#include <stdexcept>

namespace excellib {

// ============================================================
//  ZIP 読み書きヘルパー (ZipReader の内部と同じアルゴリズム)
// ============================================================
// read_le32(): 4 バイトを Little Endian として uint32_t で読む
static uint32_t read_le32(const uint8_t* p) {
    return uint32_t(p[0])|(uint32_t(p[1])<<8)|(uint32_t(p[2])<<16)|(uint32_t(p[3])<<24);
}
// read_le16(): 2 バイトを Little Endian として uint16_t で読む
static uint16_t read_le16(const uint8_t* p) {
    return uint16_t(p[0])|(uint16_t(p[1])<<8);
}
// append_le16(): uint16_t を Little Endian でバイト列に追記する
static void append_le16(std::vector<uint8_t>& v, uint16_t x) {
    v.push_back(x&0xFF); v.push_back((x>>8)&0xFF);
}
// append_le32(): uint32_t を Little Endian でバイト列に追記する
static void append_le32(std::vector<uint8_t>& v, uint32_t x) {
    v.push_back(x&0xFF); v.push_back((x>>8)&0xFF);
    v.push_back((x>>16)&0xFF); v.push_back((x>>24)&0xFF);
}

// ZipEntry2: 展開済みの ZIP エントリを表す構造体
struct ZipEntry2 {
    std::string           name;           // ファイル名 (ZIP 内パス)
    std::vector<uint8_t>  data;           // 展開済みデータ
    uint32_t local_offset{0};            // ローカルヘッダーの位置 (書き込み時に使う)
    uint32_t crc{0};                      // CRC32 チェックサム (書き込み時に計算)
};

// read_file(): バイナリファイルを読み込んでバイト列で返す
static std::vector<uint8_t> read_file(const std::string& path) {
    std::ifstream f(path, std::ios::binary);
    if (!f) throw IOError("Cannot read: " + path);
    // イテレータを使ってファイル全体を一度に読み込む
    return {std::istreambuf_iterator<char>(f), {}};
}

// write_file(): バイト列をバイナリファイルに書き込む
static void write_file(const std::string& path, const std::vector<uint8_t>& data) {
    std::ofstream f(path, std::ios::binary);
    if (!f) throw IOError("Cannot write: " + path);
    f.write(reinterpret_cast<const char*>(data.data()),
            static_cast<std::streamsize>(data.size()));
}

// zip_read_all(): ZIP ファイル全体を展開して全エントリのリストを返す
// xlsx_parser.cpp の ZipReader::parse() と同じアルゴリズムだが、
// こちらは変更してから zip_write_all() で書き直すためのリスト形式で保持する
static std::vector<ZipEntry2> zip_read_all(const std::vector<uint8_t>& data) {
    std::vector<ZipEntry2> entries;
    if (data.size() < 22) throw ParseError("ZIP too small");

    // EOCD (End of Central Directory) シグネチャ "PK\x05\x06" を後ろから探す
    size_t eocd = std::string::npos;
    for (size_t i = data.size()-22; i > 0; --i)
        if (data[i]==0x50&&data[i+1]==0x4B&&data[i+2]==0x05&&data[i+3]==0x06)
            { eocd=i; break; }
    if (eocd==std::string::npos) throw ParseError("ZIP EOCD not found");

    // EOCD から CD (セントラルディレクトリ) の開始位置とエントリ数を取得する
    uint32_t cd_off = read_le32(data.data()+eocd+16);
    uint16_t count  = read_le16(data.data()+eocd+10);
    size_t pos = cd_off;   // 現在の解析位置

    // 各 CD エントリを読む
    for (uint16_t i=0; i<count; ++i) {
        if (pos+46>data.size()) throw ParseError("CD truncated");
        // CD エントリから各フィールドを読む
        uint16_t method   = read_le16(data.data()+pos+10);   // 圧縮方式
        uint32_t comp_sz  = read_le32(data.data()+pos+20);   // 圧縮サイズ
        uint32_t uncomp_sz= read_le32(data.data()+pos+24);   // 展開後サイズ
        uint16_t fn_len   = read_le16(data.data()+pos+28);   // ファイル名長
        uint16_t ex_len   = read_le16(data.data()+pos+30);   // 拡張フィールド長
        uint16_t cm_len   = read_le16(data.data()+pos+32);   // コメント長
        uint32_t lhdr_off = read_le32(data.data()+pos+42);   // ローカルヘッダーオフセット
        // ファイル名を取得する
        std::string name(reinterpret_cast<const char*>(data.data()+pos+46), fn_len);
        pos += 46+fn_len+ex_len+cm_len;   // 次のエントリへ移動

        // ローカルヘッダーからデータ開始位置を計算する
        uint16_t lfn = read_le16(data.data()+lhdr_off+26);   // ローカルヘッダーのファイル名長
        uint16_t lex = read_le16(data.data()+lhdr_off+28);   // ローカルヘッダーの拡張フィールド長
        size_t doff  = lhdr_off+30+lfn+lex;   // データ開始位置

        ZipEntry2 e; e.name=name;
        if (method==0) {
            // 無圧縮: そのままコピー
            e.data.assign(data.data()+doff, data.data()+doff+comp_sz);
        } else if (method==8) {
            // DEFLATE: 自前の展開関数で展開する
            e.data = excellib::detail::deflate_decompress(data.data()+doff, comp_sz, uncomp_sz);
        }
        // method が 0 でも 8 でもない場合はデータなしでエントリを追加する
        entries.push_back(std::move(e));
    }
    return entries;
}

// zip_write_all(): ZipEntry2 のリストから ZIP アーカイブのバイト列を生成する
// 注意: ここでは圧縮せず Stored (method=0) で全ファイルを格納する
static std::vector<uint8_t> zip_write_all(std::vector<ZipEntry2>& entries) {
    std::vector<uint8_t> out;
    // ---- ローカルファイルヘッダー + データを追記 ----
    for (auto& e : entries) {
        e.local_offset = uint32_t(out.size());   // このエントリの先頭位置を記録
        // CRC32 を計算する (ZIP の整合性チェック)
        e.crc = excellib::detail::crc32_compute(e.data.data(), e.data.size());
        // ローカルファイルヘッダー (ZIP 仕様通りのフィールド順)
        append_le32(out,0x04034B50);                        // シグネチャ "PK\x03\x04"
        append_le16(out,20); append_le16(out,0);            // バージョン 2.0, フラグ 0
        append_le16(out,0); append_le16(out,0); append_le16(out,0); // 圧縮 0, 時刻 0, 日付 0
        append_le32(out,e.crc);                             // CRC32
        append_le32(out,uint32_t(e.data.size()));           // 圧縮サイズ
        append_le32(out,uint32_t(e.data.size()));           // 展開後サイズ
        append_le16(out,uint16_t(e.name.size())); append_le16(out,0); // ファイル名長, 拡張長 0
        out.insert(out.end(),e.name.begin(),e.name.end()); // ファイル名
        out.insert(out.end(),e.data.begin(),e.data.end()); // ファイルデータ
    }
    // ---- セントラルディレクトリを追記 ----
    uint32_t cd_off = uint32_t(out.size());   // CD の開始位置
    for (auto& e : entries) {
        append_le32(out,0x02014B50); append_le16(out,20); append_le16(out,20); // シグネチャ, バージョン
        append_le16(out,0); append_le16(out,0); append_le16(out,0); append_le16(out,0); // フラグ, 圧縮, 時刻, 日付
        append_le32(out,e.crc);                             // CRC32
        append_le32(out,uint32_t(e.data.size()));           // 圧縮サイズ
        append_le32(out,uint32_t(e.data.size()));           // 展開後サイズ
        append_le16(out,uint16_t(e.name.size()));           // ファイル名長
        append_le16(out,0); append_le16(out,0); append_le16(out,0); // 拡張長, コメント長, ディスク番号
        append_le16(out,0); append_le32(out,0);             // 内部属性, 外部属性
        append_le32(out,e.local_offset);                    // ローカルヘッダーへのオフセット
        out.insert(out.end(),e.name.begin(),e.name.end()); // ファイル名
    }
    // ---- EOCD を追記 ----
    uint32_t cd_sz = uint32_t(out.size())-cd_off;   // CD サイズ
    append_le32(out,0x06054B50); append_le16(out,0); append_le16(out,0); // シグネチャ, ディスク番号
    append_le16(out,uint16_t(entries.size()));       // このディスクの CD エントリ数
    append_le16(out,uint16_t(entries.size()));       // CD エントリ総数
    append_le32(out,cd_sz); append_le32(out,cd_off); append_le16(out,0); // CD サイズ, オフセット, コメント長
    return out;
}

// ============================================================
//  XML ヘルパー
// ============================================================
// xml_attr(): XML 要素文字列から属性値を取り出す (xlsx_parser.cpp の同名関数と同じロジック)
static std::string xml_attr(const std::string& el, const std::string& attr) {
    auto pos = el.find(attr+"=\"");
    if (pos==std::string::npos) return {};
    pos += attr.size()+2;
    auto end = el.find('"',pos);
    return end==std::string::npos ? "" : el.substr(pos,end-pos);
}

// to_str(): double を見やすい文字列に変換する (末尾のゼロを除去)
// "%.4f" で小数点以下 4 桁にフォーマットして末尾 0 を削除する
// 例: 0.7000 → "0.7", 0.7500 → "0.75"
static std::string to_str(double v) {
    char buf[32]; std::snprintf(buf,sizeof(buf),"%.4f",v);
    std::string s(buf);
    // 小数点がある場合、末尾のゼロを削除する
    auto dot = s.find('.');
    if (dot!=std::string::npos) {
        // 末尾から最後の非ゼロ文字を探す
        size_t last = s.find_last_not_of('0');
        if (last==dot) ++last;   // "0." の場合は "0.0" にする
        s = s.substr(0,last+1);
    }
    return s;
}

// ============================================================
//  PageSetup 構造体から XLSX の XML 要素を生成する
// ============================================================

// build_page_setup_xml(): <pageSetup> 要素を生成する
// XLSX の sheet.xml に挿入するページ設定 XML
static std::string build_page_setup_xml(const PageSetup& ps) {
    std::ostringstream o;
    o << "<pageSetup";
    // paperSize: 用紙サイズ (例: 9 = A4, 1 = Letter)
    o << " paperSize=\"" << uint16_t(ps.paper_size) << "\"";

    // orientation: 印刷方向 ("landscape" または "portrait")
    if (ps.orientation == Orientation::Landscape)
        o << " orientation=\"landscape\"";
    else
        o << " orientation=\"portrait\"";

    // FitTo 設定 (ページ数に自動縮小する場合)
    if (ps.fit_to != FitTo::None) {
        o << " fitToPage=\"1\"";   // fitToPage="1" で FitTo モードを有効化
        // fitToWidth: 横方向のページ数 (0 = 自動)
        if (ps.fit_to==FitTo::Width || ps.fit_to==FitTo::WidthAndHeight)
            o << " fitToWidth=\"" << ps.fit_to_pages_wide << "\"";
        else
            o << " fitToWidth=\"0\"";   // 0 = 自動 (制限なし)
        // fitToHeight: 縦方向のページ数 (0 = 自動)
        if (ps.fit_to==FitTo::Height || ps.fit_to==FitTo::WidthAndHeight)
            o << " fitToHeight=\"" << ps.fit_to_pages_tall << "\"";
        else
            o << " fitToHeight=\"0\"";
    } else {
        // FitTo なし: scale で倍率を指定する
        o << " scale=\"" << ps.scale_percent << "\"";
        o << " fitToPage=\"0\"";   // fitToPage="0" で通常印刷モード
    }

    // firstPageNumber: 先頭ページ番号 (0 はデフォルトなので設定しない)
    if (ps.first_page_number > 0)
        o << " firstPageNumber=\"" << ps.first_page_number
          << "\" useFirstPageNumber=\"1\"";   // useFirstPageNumber="1" も必要

    // blackAndWhite: 白黒印刷
    if (ps.black_and_white)    o << " blackAndWhite=\"1\"";
    // draft: 下書き品質印刷
    if (ps.draft_quality)      o << " draft=\"1\"";
    // pageOrder: ページ順序 (overThenDown = 横から先に印刷)
    if (ps.page_order_over_then_down) o << " pageOrder=\"overThenDown\"";
    // copies: 印刷部数
    if (ps.copies > 1)         o << " copies=\"" << ps.copies << "\"";

    o << "/>";
    return o.str();
}

// build_page_margins_xml(): <pageMargins> 要素を生成する
// XLSX では余白はインチ単位で指定する (VBA の InchesToPoints とは異なり変換不要)
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

// build_print_options_xml(): <printOptions> 要素を生成する
// 印刷オプション (グリッド線・行列見出し) を設定する
// どちらも無効な場合は空文字列を返す (要素不要)
static std::string build_print_options_xml(const PageSetup& ps) {
    if (!ps.print_gridlines && !ps.print_row_col_headings) return "";
    std::ostringstream o;
    o << "<printOptions";
    if (ps.print_gridlines)
        // gridLines="1": グリッド線を印刷する, gridLinesSet="1": グリッド線設定を有効化する
        o << " gridLines=\"1\" gridLinesSet=\"1\"";
    if (ps.print_row_col_headings)
        // headings="1": 行番号・列文字を印刷する
        o << " headings=\"1\"";
    o << "/>";
    return o.str();
}

// build_header_footer_xml(): <headerFooter> 要素を生成する
// ヘッダー・フッターの設定 (全て空の場合は空文字列を返す)
static std::string build_header_footer_xml(const HeaderFooter& hf) {
    // 全て空なら要素不要
    if (hf.odd_header.empty() && hf.odd_footer.empty() &&
        hf.even_header.empty() && hf.even_footer.empty()) return "";
    std::ostringstream o;
    o << "<headerFooter";
    // 奇数・偶数ページで異なるヘッダー/フッターを使う場合
    if (hf.different_odd_even) o << " differentOddEven=\"1\"";
    // 先頭ページだけ異なる場合
    if (hf.different_first)    o << " differentFirst=\"1\"";
    o << ">";
    // 奇数ページのヘッダー (例: "&C&14Report" = 中央に 14pt フォントで "Report")
    if (!hf.odd_header.empty())  o << "<oddHeader>"  << hf.odd_header  << "</oddHeader>";
    // 奇数ページのフッター (例: "&LPage &P / &N" = 左端にページ番号/総ページ数)
    if (!hf.odd_footer.empty())  o << "<oddFooter>"  << hf.odd_footer  << "</oddFooter>";
    // 偶数ページのヘッダー
    if (!hf.even_header.empty()) o << "<evenHeader>" << hf.even_header << "</evenHeader>";
    // 偶数ページのフッター
    if (!hf.even_footer.empty()) o << "<evenFooter>" << hf.even_footer << "</evenFooter>";
    o << "</headerFooter>";
    return o.str();
}

// ============================================================
//  シート XML にページ設定を注入する
// ============================================================
// inject_page_setup(): sheet.xml の XML 文字列に pageSetup などの要素を挿入する
// 既存の設定要素を削除してから </sheetData> の直後に新しい設定を挿入する
static std::string inject_page_setup(const std::string& sheet_xml, const PageSetup& ps) {
    std::string xml = sheet_xml;

    // 既存の pageSetup / pageMargins / printOptions / headerFooter 要素を削除するラムダ
    auto remove_element = [&](const std::string& tag) {
        // 自己終了タグ (<tag .../>) とペアタグ (<tag>...</tag>) の両方を削除する
        while (true) {
            // タグの開始 "<tag" を探す
            auto s = xml.find("<" + tag);
            if (s == std::string::npos) break;
            // 自己終了タグの終端 "/>" を探す
            auto e = xml.find("/>", s);
            // 通常タグの終端 "</tag>" を探す
            auto e2 = xml.find("</" + tag + ">", s);
            size_t end_pos;
            // 自己終了タグが通常タグより前に見つかれば自己終了タグとして扱う
            if (e != std::string::npos && (e2 == std::string::npos || e < e2))
                end_pos = e + 2;   // "/>" の 2 文字後まで削除
            else if (e2 != std::string::npos)
                end_pos = e2 + tag.size() + 3;   // "</tag>" の後まで削除
            else break;
            xml.erase(s, end_pos - s);   // 要素全体を削除する
        }
    };

    // 既存の設定要素を全て削除する (後で新しい設定を挿入するため)
    remove_element("pageSetup");
    remove_element("pageMargins");
    remove_element("printOptions");
    remove_element("headerFooter");

    // 挿入位置を </sheetData> タグの直後に決める
    std::string insert_point = "</sheetData>";
    auto ins = xml.find(insert_point);
    if (ins == std::string::npos) {
        // </sheetData> が見つからない場合は自己終了タグ <sheetData/> を試みる
        insert_point = "<sheetData/>";
        ins = xml.find(insert_point);
        if (ins != std::string::npos) ins += insert_point.size() - 2;   // タグの終端 ">" の前
    } else {
        ins += insert_point.size();   // </sheetData> の末尾の次の位置
    }

    // 挿入する XML 要素を順番に組み立てる
    std::string additions;
    additions += build_print_options_xml(ps);       // <printOptions>
    additions += build_page_margins_xml(ps.margins); // <pageMargins>
    additions += build_page_setup_xml(ps);           // <pageSetup>
    additions += build_header_footer_xml(ps.header_footer); // <headerFooter>

    // 組み立てた要素を挿入する
    if (ins != std::string::npos)
        xml.insert(ins, additions);
    else
        xml += additions;   // フォールバック: XML の末尾に追記する

    return xml;
}

// ============================================================
//  シート名から ZIP 内のパスを解決する
// ============================================================
// resolve_sheet_path(): シート名に対応する ZIP 内の XML ファイルパスを返す
// workbook.xml → rId → .rels → ターゲットパスの順に解決する
static std::string resolve_sheet_path(
    const std::vector<ZipEntry2>& entries,
    const std::string& sheet_name)
{
    // workbook.xml を見つける
    std::string wb_xml;
    for (auto& e : entries)
        if (e.name == "xl/workbook.xml")
            { wb_xml = std::string(e.data.begin(), e.data.end()); break; }
    if (wb_xml.empty()) throw ParseError("workbook.xml not found");

    // workbook.xml.rels を見つける
    std::string rels_xml;
    for (auto& e : entries)
        if (e.name == "xl/_rels/workbook.xml.rels")
            { rels_xml = std::string(e.data.begin(), e.data.end()); break; }

    // workbook.xml からシート名に対応する rId を探す
    std::string target_rid;
    size_t pos = 0;
    while (true) {
        auto sh = wb_xml.find("<sheet ", pos);
        if (sh == std::string::npos) break;
        auto she = wb_xml.find('>', sh);
        std::string elem = wb_xml.substr(sh, she-sh+1);
        std::string name = xml_attr(elem, "name");   // シート名属性を取得
        // シート名が一致する、またはシート名が空 (最初のシートを使う) 場合
        if (name == sheet_name || sheet_name.empty()) {
            target_rid = xml_attr(elem, "r:id");   // リレーション ID を取得
            break;
        }
        pos = she;
    }
    // シート名が見つからなかった場合はエラー
    if (target_rid.empty() && !sheet_name.empty())
        throw RangeError("Sheet not found: " + sheet_name);
    // シート名が空 (デフォルト = アクティブシート) の場合は最初のシートを使う
    if (target_rid.empty()) {
        auto sh = wb_xml.find("<sheet ");
        if (sh == std::string::npos) throw ParseError("No sheets in workbook");
        auto she = wb_xml.find('>', sh);
        target_rid = xml_attr(wb_xml.substr(sh, she-sh+1), "r:id");
    }

    // .rels ファイルから rId に対応するターゲットパスを解決する
    pos = 0;
    while (true) {
        auto r = rels_xml.find("<Relationship ", pos);
        if (r == std::string::npos) break;
        auto re = rels_xml.find('>', r);
        std::string elem = rels_xml.substr(r, re-r+1);
        if (xml_attr(elem, "Id") == target_rid) {
            std::string target = xml_attr(elem, "Target");
            // パスが "/" で始まれば絶対パス → "xl" を前置する
            if (!target.empty() && target[0] == '/')
                return "xl" + target;
            // 相対パス → "xl/" を前置する
            return "xl/" + target;
        }
        pos = re;
    }
    throw ParseError("Could not resolve sheet path for rId=" + target_rid);
}

// ============================================================
//  workbook.xml に definedNames (印刷範囲・タイトル行) を設定する
// ============================================================
// inject_defined_names(): workbook.xml に _xlnm.Print_Area と _xlnm.Print_Titles の
// definedName 要素を挿入する
// 印刷範囲は sheet.xml の pageSetup ではなく workbook.xml の definedNames に定義する
static void inject_defined_names(
    std::string& wb_xml,
    const std::string& sheet_name,
    const PageSetup& ps)
{
    // 印刷範囲もタイトルも指定されていなければ何もしない
    if (!ps.print_area && !ps.repeat_titles) return;

    // 既存の <definedNames>...</definedNames> ブロックを削除する
    auto dns = wb_xml.find("<definedNames");
    if (dns != std::string::npos) {
        auto dne = wb_xml.find("</definedNames>", dns);
        if (dne != std::string::npos)
            wb_xml.erase(dns, dne - dns + 15);   // "</definedNames>" の長さ 15 を加算
    }

    std::ostringstream dn;
    dn << "<definedNames>";

    // 印刷範囲 (_xlnm.Print_Area) の definedName を追加する
    if (ps.print_area) {
        // A1 記法の座標を $A$1 形式 (絶対参照) に変換するラムダ
        // 例: "A1" → "$A$1", "AA10" → "$AA$10"
        auto make_absolute = [](const std::string& a1) -> std::string {
            std::string r;
            size_t i=0;
            // アルファベット部分 (列) の各文字の前に '$' を挿入する
            while (i<a1.size() && std::isalpha((unsigned char)a1[i]))
                { r += '$'; r += a1[i++]; }
            r += '$';   // 数字 (行) の前に '$' を挿入する
            while (i<a1.size()) r += a1[i++];   // 数字部分をそのままコピー
            return r;
        };
        // 開始・終了セルの座標を絶対参照形式に変換する
        std::string abs_start = make_absolute(
            CellAddress{ps.print_area->first_row, ps.print_area->first_col}.to_a1());
        std::string abs_end = make_absolute(
            CellAddress{ps.print_area->last_row, ps.print_area->last_col}.to_a1());
        // 書式: 'シート名'!$A$1:$Z$50
        dn << "<definedName name=\"_xlnm.Print_Area\">"
           << "'" << sheet_name << "'!" << abs_start << ":" << abs_end
           << "</definedName>";
    }

    // タイトル行・列の繰り返し (_xlnm.Print_Titles) の definedName を追加する
    if (ps.repeat_titles) {
        auto& rt = *ps.repeat_titles;
        std::string val;
        // タイトル行: 書式 'シート名'!$1:$3 (1〜3 行目を繰り返す)
        if (rt.row_start && rt.row_end) {
            val += "'" + sheet_name + "'!$" +
                   std::to_string(*rt.row_start + 1) + ":$" +
                   std::to_string(*rt.row_end + 1);
        }
        // タイトル列: 書式 'シート名'!$A:$C (A〜C 列を繰り返す)
        if (rt.col_start && rt.col_end) {
            if (!val.empty()) val += ",";   // 行と列の両方がある場合はカンマで区切る
            // 列番号 (0-indexed) から列文字だけを取り出すラムダ
            // CellAddress::to_a1() は "A1" を返すので、アルファベット部分だけを取り出す
            auto col_letters = [](const std::string& a1) {
                std::string r;
                for (char c : a1) if (std::isalpha((unsigned char)c)) r+=c;
                return r;
            };
            // row=0 で to_a1() を呼んで列文字部分だけを取り出す
            std::string cs = CellAddress{0, *rt.col_start}.to_a1();
            std::string ce = CellAddress{0, *rt.col_end}.to_a1();
            val += "'" + sheet_name + "'!$" + col_letters(cs)
                + ":$" + col_letters(ce);
        }
        if (!val.empty())
            dn << "<definedName name=\"_xlnm.Print_Titles\">"
               << val << "</definedName>";
    }

    dn << "</definedNames>";

    // </workbook> タグの直前に definedNames ブロックを挿入する
    auto end = wb_xml.rfind("</workbook>");
    if (end != std::string::npos)
        wb_xml.insert(end, dn.str());
}

// ============================================================
//  ExcelPrinter::write_page_setup_to_xlsx (内部実装)
// ============================================================
// write_page_setup_to_xlsx(): XLSX ファイルを読み込み、ページ設定 XML を書き換えて保存する
// 処理の流れ:
//   1. XLSX (= ZIP) ファイルを読み込んで全エントリを展開する
//   2. 対象シートのパスを解決する (workbook.xml + .rels を使う)
//   3. シートの XML に <pageSetup> <pageMargins> などを注入する
//   4. workbook.xml に <definedNames> (印刷範囲・タイトル) を注入する
//   5. 変更した内容を ZIP に再パッケージして上書き保存する
void ExcelPrinter::write_page_setup_to_xlsx(
    const std::string& xlsx_path,
    const std::string& sheet_name,
    const PageSetup&   setup)
{
    auto raw     = read_file(xlsx_path);           // XLSX ファイルをバイト列で読み込む
    auto entries = zip_read_all(raw);              // ZIP を展開してエントリリストを取得する

    // 対象シートの ZIP 内パスを解決する
    std::string sheet_path = resolve_sheet_path(entries, sheet_name);

    // 各エントリを修正する
    for (auto& e : entries) {
        if (e.name == sheet_path) {
            // 対象シートの XML を修正する
            std::string xml(e.data.begin(), e.data.end());   // バイト列 → 文字列
            xml = inject_page_setup(xml, setup);              // ページ設定を注入
            e.data = std::vector<uint8_t>(xml.begin(), xml.end());   // 文字列 → バイト列
        }
        // workbook.xml に definedNames (印刷範囲・タイトル) を注入する
        if (e.name == "xl/workbook.xml") {
            std::string xml(e.data.begin(), e.data.end());
            // sheet_name が空の場合は workbook.xml から最初のシート名を取得する
            std::string sname = sheet_name;
            if (sname.empty()) {
                auto sh = xml.find("<sheet ");
                if (sh != std::string::npos) {
                    auto she = xml.find('>', sh);
                    sname = xml_attr(xml.substr(sh, she-sh+1), "name");
                }
            }
            inject_defined_names(xml, sname, setup);   // definedNames を注入
            e.data = std::vector<uint8_t>(xml.begin(), xml.end());
        }
    }

    // 変更済みエントリを ZIP に再パッケージしてファイルに上書き保存する
    write_file(xlsx_path, zip_write_all(entries));
}

// ============================================================
//  apply_page_setup (公開 API)
// ============================================================
// apply_page_setup(): XLSX ファイルにページ設定を適用する (Excel 不要)
// xlsx_path: 対象 XLSX ファイルのパス
// sheet_name: 設定を適用するシート名 (空なら最初のシート)
// setup: 適用するページ設定
void ExcelPrinter::apply_page_setup(
    const std::string& xlsx_path,
    const std::string& sheet_name,
    const PageSetup&   setup)
{
    log("Applying page setup to: " + xlsx_path);   // プログレスメッセージ
    write_page_setup_to_xlsx(xlsx_path, sheet_name, setup);   // 内部実装に委譲
    log("Page setup applied.");
}

} // namespace excellib
