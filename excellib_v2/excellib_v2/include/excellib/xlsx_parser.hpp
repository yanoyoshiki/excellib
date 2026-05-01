#pragma once
// workbook.hpp: CellValue / Sheet / Workbook などの基本型を提供するヘッダー
#include "excellib/workbook.hpp"
// std::vector: 動的配列
#include <vector>
// std::unordered_map: ハッシュマップ (O(1) 検索)
#include <unordered_map>
// std::map: 順序付きマップ (行・列番号をキーにして Cell を格納)
#include <map>
// std::string: 文字列型
#include <string>
// std::unique_ptr / std::shared_ptr: スマートポインタ
#include <memory>
// std::function: 関数オブジェクトを保持する型
#include <functional>

// excellib::xlsx は XLSX 形式固有のクラスをまとめる名前空間
namespace excellib::xlsx {

// ============================================================
//  ZIP エントリ
// ============================================================
// XLSX ファイルは ZIP アーカイブ。ZIP 内の 1 ファイルを表す構造体
struct ZipEntry {
    std::string              path;   // ZIP 内のパス (例: "xl/worksheets/sheet1.xml")
    std::vector<uint8_t>     data;   // 展開済みのファイルデータ (バイト列)
};

// ============================================================
//  最小限の ZIP リーダー (外部依存なし)
// ============================================================
// XLSX ファイル (= ZIP) を読み込んで内部ファイルを取り出すクラス
// zlib などの外部ライブラリを使わず、自前の DEFLATE 解凍で動作する
class ZipReader {
public:
    // コンストラクタ: バイト列を受け取って即座に ZIP をパースする
    explicit ZipReader(const std::vector<uint8_t>& data);

    // has(): 指定パスのエントリが存在するか確認する
    bool              has(const std::string& path) const;
    // get(): 指定パスのエントリを取得する (なければ IOError)
    const ZipEntry&   get(const std::string& path) const;
    // paths(): ZIP 内の全ファイルパスを返す
    std::vector<std::string> paths() const;

    // text(): 指定パスのファイルを展開して文字列で返す
    std::string text(const std::string& path) const;

private:
    // entries_: パス → ZipEntry のハッシュマップ (O(1) 検索)
    std::unordered_map<std::string, ZipEntry> entries_;

    // parse(): ZIP バイト列を解析して entries_ を埋める内部メソッド
    void parse(const std::vector<uint8_t>& data);
    // inflate(): raw DEFLATE (btype=1/2) を展開する内部静的メソッド
    // compressed: 圧縮データ, comp_size: 圧縮バイト数, uncomp_size: 展開後の目安サイズ
    static std::vector<uint8_t> inflate(const uint8_t* compressed, size_t comp_size,
                                         size_t uncomp_size);
};

// ============================================================
//  リレーションシップ (.rels ファイル) のパーサー
// ============================================================
// OOXML では各ファイルが .rels ファイルで「どのファイルとつながっているか」を管理する
// 例: workbook.xml.rels にはシートごとの rId → ファイルパスのマッピングが入る
struct Relationship {
    std::string id;      // リレーション ID (例: "rId1")
    std::string type;    // リレーションシップの種類 URL (長い文字列)
    std::string target;  // 参照先ファイルパス (例: "worksheets/sheet1.xml")
};

// parse_rels(): .rels ファイルの XML 文字列を解析して Relationship のリストを返す
std::vector<Relationship> parse_rels(const std::string& xml);

// ============================================================
//  共有文字列テーブル (xl/sharedStrings.xml)
// ============================================================
// XLSX では文字列セルの値を「共有文字列テーブル」に一元管理する
// 各セルは文字列そのものではなく、テーブルへのインデックスを持つ
class XlsxSharedStrings {
public:
    // parse(): sharedStrings.xml の XML 文字列をパースして strings_ に格納する
    void parse(const std::string& xml);
    // get(): インデックスに対応する文字列を返す (範囲外なら RangeError)
    const std::string& get(uint32_t index) const;
    // size(): 文字列テーブルのサイズを返す
    size_t size() const { return strings_.size(); }

    // intern(): 書き込み用。文字列を登録して既存ならそのインデックスを返す (重複排除)
    uint32_t intern(const std::string& s);
    // to_xml(): 共有文字列テーブルを XML 文字列に変換する (書き込み用)
    std::string to_xml() const;

private:
    // strings_: 登録済み文字列のリスト (インデックス = 位置)
    std::vector<std::string>            strings_;
    // index_map_: 文字列 → インデックスの逆引きマップ (書き込み時に重複排除に使う)
    std::unordered_map<std::string, uint32_t> index_map_;
};

// ============================================================
//  数値フォーマットとスタイル (xl/styles.xml)
// ============================================================
// XlsxXF: セルの書式設定レコード (XF = eXtended Format)
// 各セルは XF インデックスを持ち、対応する書式 (日付/通貨 etc.) が決まる
struct XlsxXF {
    uint16_t num_fmt_id{0};           // 数値フォーマット ID (14=日付, 0=標準 など)
    uint16_t font_id{0};              // フォント ID
    uint16_t fill_id{0};              // 塗りつぶし ID
    uint16_t border_id{0};            // 罫線 ID
    bool     apply_number_format{false}; // 数値フォーマットを適用するか
};

// XlsxStyles: styles.xml 全体を管理するクラス
class XlsxStyles {
public:
    // parse(): styles.xml の XML 文字列をパースして num_fmts_ と cell_xfs_ を埋める
    void parse(const std::string& xml);

    // cell_format(): XF インデックスに対応する FormatInfo を返す
    // FormatInfo には format_string と is_date_format が入る
    FormatInfo cell_format(uint32_t xf_index) const;
    // xf_count(): XF テーブルのサイズを返す
    size_t     xf_count() const { return cell_xfs_.size(); }
    // to_xml(): styles.xml 形式の XML 文字列を生成する (書き込み用)
    std::string to_xml() const;

private:
    // num_fmts_: カスタム数値フォーマット ID → フォーマット文字列のマップ
    // 例: 166 → "yyyy/mm/dd"
    std::unordered_map<uint16_t, std::string> num_fmts_;
    // cell_xfs_: セル書式レコードのリスト (インデックス = XF インデックス)
    std::vector<XlsxXF>                        cell_xfs_;
    // is_date_format(): 組み込み ID またはフォーマット文字列から日付書式か判定する
    static bool is_date_format(uint16_t id, const std::string& fmt);
    // looks_like_date(): フォーマット文字列に y/d などの日付記号が含まれるか調べる
    static bool looks_like_date(const std::string& fmt);
};

// ============================================================
//  XLSX シートの実装クラス
// ============================================================
// Sheet 抽象クラスを継承した XLSX 専用のシート実装
class XlsxSheet : public Sheet {
public:
    // コンストラクタ: シート名を受け取り name_ に格納する
    // std::move() でコピーを避けて所有権を移動する
    explicit XlsxSheet(std::string name) : name_(std::move(name)) {}

    // ---- Sheet 仮想関数のオーバーライド ----
    std::string name() const override { return name_; }         // シート名を返す
    uint32_t    row_count() const override { return row_count_; } // 行数を返す
    uint32_t    col_count() const override { return col_count_; } // 列数を返す

    // cell(): 行・列インデックスでセルを取得 (空セルは Blank を返す)
    Cell cell(uint32_t row, uint32_t col) const override;
    // cell(): CellAddress 構造体でセルを取得
    Cell cell(const CellAddress& addr) const override;
    // cell(): "A1" 記法の文字列でセルを取得
    Cell cell(const std::string& a1) const override;
    // try_cell(): 空セルなら nullopt を返す安全版
    std::optional<Cell> try_cell(uint32_t row, uint32_t col) const override;

    // row(): 指定行の全セルを vector で返す
    std::vector<Cell> row(uint32_t row_idx) const override;
    // cells(): シート内の全非空セルを vector で返す
    std::vector<Cell> cells() const override;

    // set_cell(): セルに値を書き込む (行・列インデックス版)
    void set_cell(uint32_t row, uint32_t col, const CellValue& value) override;
    // set_cell(): "A1" 記法でセルに値を書き込む
    void set_cell(const std::string& a1, const CellValue& value) override;
    // set_formula(): セルに数式を設定する ("=SUM(A1:A10)" など)
    void set_formula(const std::string& a1, const std::string& formula) override;

    // for_each_cell(): 全セルを fn に渡して走査するイテレータ
    void for_each_cell(std::function<void(const Cell&)> fn) const override;

    // put_cell(): パーサーが解析済みセルを直接格納するための内部 API
    void put_cell(const Cell& c);
    // set_dimensions(): 行数・列数を直接設定する (DIMENSION レコードから)
    void set_dimensions(uint32_t rows, uint32_t cols);

private:
    std::string name_;           // シート名
    uint32_t    row_count_{0};   // 現在の最大行数 (0-indexed + 1)
    uint32_t    col_count_{0};   // 現在の最大列数 (0-indexed + 1)
    // data_: 行番号 → (列番号 → Cell) の二重マップ
    // std::map は挿入順でなくキー順に並ぶため、行・列の順序が保たれる
    std::map<uint32_t, std::map<uint32_t, Cell>> data_;
};

// ============================================================
//  XLSX ワークブックの実装クラス
// ============================================================
// Workbook 抽象クラスを継承した XLSX 専用のワークブック実装
class XlsxWorkbook : public Workbook {
public:
    // ---- Workbook 仮想関数のオーバーライド ----
    FileFormat   format() const override { return FileFormat::XLSX; }  // フォーマット種別
    std::string  format_name() const override { return "XLSX"; }        // フォーマット名
    size_t       sheet_count() const override { return sheets_.size(); } // シート数

    // sheet(): インデックスでシートを取得 (const/非 const 両方)
    Sheet&       sheet(size_t index) override;
    const Sheet& sheet(size_t index) const override;
    // sheet(): シート名でシートを取得
    Sheet&       sheet(const std::string& name) override;
    // sheet_names(): 全シート名のリストを返す
    std::vector<std::string> sheet_names() const override;

    // add_sheet(): 新しいシートを末尾に追加して参照を返す
    Sheet&       add_sheet(const std::string& name) override;
    // remove_sheet(): インデックス指定でシートを削除する
    void         remove_sheet(size_t index) override;
    // rename_sheet(): シート名を変更する (重複名は WriteError)
    void         rename_sheet(size_t index, const std::string& name) override;

    // save(): ファイルパスに XLSX として保存する
    void         save(const std::string& path, const SaveOptions& opts = {}) const override;
    // to_bytes(): XLSX バイト列 (= ZIP バイナリ) を生成して返す
    std::vector<uint8_t> to_bytes(FileFormat fmt, const SaveOptions& opts = {}) const override;

    // add_parsed_sheet(): パーサーが解析済みシートを追加するための内部 API
    void add_parsed_sheet(std::unique_ptr<XlsxSheet> sheet);

private:
    // sheets_: XlsxSheet の unique_ptr を保持するリスト
    // unique_ptr で所有権を明確にし、ワークブック破棄時に自動でシートも解放される
    std::vector<std::unique_ptr<XlsxSheet>> sheets_;
    // sst_: 共有文字列テーブル (書き込み時に再利用)
    XlsxSharedStrings                        sst_;
    // styles_: スタイルテーブル (書き込み時に再利用)
    XlsxStyles                               styles_;
};

// ============================================================
//  XLSX パーサー
// ============================================================
// XLSX ファイル (ZIP + XML) を解析して XlsxWorkbook を生成するクラス
class XlsxParser {
public:
    // コンストラクタ: オープンオプション (strict_validation など) を受け取る
    explicit XlsxParser(const OpenOptions& opts) : opts_(opts) {}

    // parse(): XLSX バイト列を解析して XlsxWorkbook を返す
    std::unique_ptr<XlsxWorkbook> parse(const std::vector<uint8_t>& data);

private:
    OpenOptions opts_;   // オープンオプション (パース中に参照)

    // SheetMeta: workbook.xml から収集したシートのメタ情報
    // name=シート名, rid=リレーション ID, path=ZIP 内のファイルパス
    struct SheetMeta { std::string name, rid, path; };

    // parse_workbook_xml(): workbook.xml を解析して SheetMeta のリストを返す
    // rels: workbook.xml.rels の解析結果 (rId → パスの対応)
    std::vector<SheetMeta> parse_workbook_xml(
        const std::string& xml,
        const std::vector<Relationship>& rels
    );

    // parse_sheet_xml(): シートの XML を解析して XlsxSheet を返す
    // name: シート名, xml: シート XML 文字列, sst: 共有文字列テーブル, styles: スタイル
    std::unique_ptr<XlsxSheet> parse_sheet_xml(
        const std::string& name,
        const std::string& xml,
        const XlsxSharedStrings& sst,
        const XlsxStyles& styles
    );

    // parse_cell_element(): XML の <c> 要素を解析して Cell を返す
    // ref: セル参照 (例: "A1"), type_attr: t 属性, s_attr: s 属性 (スタイル ID)
    // v_text: <v> タグの内容, f_text: <f> タグの内容 (数式)
    Cell parse_cell_element(
        const std::string& ref,
        const std::string& type_attr,
        const std::string& s_attr,
        const std::string& v_text,
        const std::string& f_text,
        const XlsxSharedStrings& sst,
        const XlsxStyles& styles
    );

    // parse_ref(): "B3" のような A1 記法を CellAddress (0-indexed) に変換する
    static CellAddress parse_ref(const std::string& ref);
};

// ============================================================
//  XLSX ライター
// ============================================================
// XlsxWorkbook を XLSX バイト列 (ZIP + XML) に変換するクラス
class XlsxWriter {
public:
    // write(): ワークブックを XLSX バイト列として書き出す
    std::vector<uint8_t> write(const XlsxWorkbook& wb, const SaveOptions& opts);

private:
    // ZipBuilder: ZIP ファイル組み立て用の内部構造体 (実装は .cpp 内)
    struct ZipBuilder;

    // sheet_xml(): シートの XML 文字列を生成する。文字列セルは sst に登録する
    std::string sheet_xml(const XlsxSheet& sheet, XlsxSharedStrings& sst);
    // workbook_xml(): workbook.xml の XML 文字列を生成する
    std::string workbook_xml(const std::vector<std::string>& sheet_names);
    // workbook_rels_xml(): workbook.xml.rels の XML 文字列を生成する
    std::string workbook_rels_xml(size_t sheet_count);
    // content_types_xml(): [Content_Types].xml の XML 文字列を生成する
    std::string content_types_xml(size_t sheet_count, bool has_shared_strings);
    // root_rels_xml(): _rels/.rels の XML 文字列を生成する
    std::string root_rels_xml();

    // cell_ref(): 0-indexed の行・列番号を "B3" のような A1 記法に変換する
    static std::string cell_ref(uint32_t row, uint32_t col);
    // escape_xml(): XML の特殊文字 (&, <, >, " など) をエスケープする
    static std::string escape_xml(const std::string& s);
};

} // namespace excellib::xlsx
