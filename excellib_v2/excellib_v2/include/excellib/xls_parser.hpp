#pragma once
// workbook.hpp: CellValue / Sheet / Workbook などの基本型を提供するヘッダー
#include "excellib/workbook.hpp"
// std::vector: 動的配列
#include <vector>
// std::unordered_map: ハッシュマップ
#include <unordered_map>
// std::map: 順序付きマップ (行・列番号をキーにして Cell を格納)
#include <map>
// uint8_t / uint16_t / uint32_t など固定幅整数型
#include <cstdint>
// std::string: 文字列型
#include <string>
// std::unique_ptr: スマートポインタ
#include <memory>

// excellib::xls は XLS (BIFF8) 形式固有のクラスをまとめる名前空間
namespace excellib::xls {

// ---- BIFF8 レコードタイプコード一覧 ----
// BIFF8 (Binary Interchange File Format 8) は Excel 97-2003 の内部バイナリ形式
// ファイルは「レコード」の連続で構成され、各レコードは 2 バイトの型コードで始まる
// 以下は主要なレコード型を列挙型 (enum class) で定義したもの
enum class RecordType : uint16_t {
    BOF         = 0x0809,   // Beginning Of File: シートまたはワークブックストリームの開始
    EOF_        = 0x000A,   // End Of File: ストリームの終端
    BOUNDSHEET  = 0x0085,   // シート情報 (シート名, オフセット, 種別)
    SST         = 0x00FC,   // Shared String Table: 共有文字列テーブル
    CONTINUE    = 0x003C,   // 前のレコードが大きすぎて分割された続きデータ
    LABELSST    = 0x00FD,   // 共有文字列テーブルを参照する文字列セル
    LABEL       = 0x0204,   // 文字列セル (SST 非使用、直接文字列を格納)
    NUMBER      = 0x0203,   // 浮動小数点数セル (8バイト IEEE 754)
    BOOLERR     = 0x0205,   // 真偽値またはエラー値セル
    BLANK       = 0x0201,   // 空セル (書式情報のみ)
    MULBLANK    = 0x00BE,   // 複数の空セル (連続範囲をまとめて格納)
    MULRK       = 0x00BD,   // 複数の RK 値セル (連続範囲をまとめて格納)
    RK          = 0x027E,   // RK 値セル (整数または小数を圧縮した独自形式)
    FORMULA     = 0x0006,   // 数式セル (計算結果 + 数式トークン列)
    STRING_     = 0x0207,   // 直前の FORMULA の文字列結果 (別レコードで送られる)
    XF          = 0x00E0,   // eXtended Format: セルの書式設定 (フォント, 数値形式 etc.)
    FORMAT      = 0x041E,   // カスタム数値フォーマット文字列の定義
    DIMENSION   = 0x0200,   // シートのデータ範囲 (使用行・列数)
    ROW         = 0x0208,   // 行の高さや書式情報
    MERGEDCELLS = 0x00E5,   // セル結合の範囲情報
    UNKNOWN     = 0xFFFF,   // 未知のレコード型 (パーサーが認識しないもの)
};

// BiffRecord: BIFF8 の 1 レコードを表す構造体
// BIFF8 のレコード形式: [type: 2byte][size: 2byte][data: size bytes]
struct BiffRecord {
    RecordType           type{RecordType::UNKNOWN}; // 解釈済みのレコード型 (enum)
    uint16_t             raw_type{0};                // 生の 2 バイト型コード
    std::vector<uint8_t> data;                       // レコード本体データ

    // size(): データ部分のバイト数を返す
    size_t size() const { return data.size(); }

    // ---- 境界チェック付きフィールド読み出しメソッド ----
    // offset 位置から指定サイズの値を Little Endian で読み出す
    // 範囲外アクセスは ParseError を投げるので安全
    uint8_t  u8 (size_t offset) const;   // 1バイト符号なし整数
    uint16_t u16(size_t offset) const;   // 2バイト符号なし整数 (Little Endian)
    uint32_t u32(size_t offset) const;   // 4バイト符号なし整数 (Little Endian)
    int16_t  i16(size_t offset) const;   // 2バイト符号付き整数 (Little Endian)
    double   f64(size_t offset) const;   // 8バイト IEEE 754 倍精度浮動小数点数
};

// decode_rk(): BIFF8 の RK 値を double に変換する
// RK 値は 4 バイトの独自圧縮形式。最下位 2 ビットがフラグ:
//   bit0 = 1 → 100 で割る (1/100 スケール)
//   bit1 = 1 → 整数として解釈 (bit1=0 なら IEEE 754 上位 30bit)
double decode_rk(uint32_t rk);

// decode_biff8_string(): BIFF8 のバイト列から Unicode 文字列をデコードして UTF-8 で返す
// BIFF8 文字列ヘッダー形式: [長さ: 2byte][フラグ: 1byte][...(オプション)...][文字データ]
//   フラグ bit0 = 1 → UTF-16LE (各文字 2 バイト)
//   フラグ bit0 = 0 → 圧縮 ASCII (各文字 1 バイト、Latin-1 相当)
//   フラグ bit2 = 1 → 拡張データあり (East Asian フォントなど)
//   フラグ bit3 = 1 → リッチテキストあり (書式情報の追加ヘッダー)
// data: 文字列データへのポインタ, available: 読める最大バイト数
// bytes_consumed: 読んだバイト数が書き出される (呼び出し元が位置を進めるために使う)
// short_header: true なら長さを 1 バイトで読む (一部の内部文字列で使用)
std::string decode_biff8_string(const uint8_t* data, size_t available,
                                 size_t& bytes_consumed,
                                 bool short_header = false);

// SharedStringTable: SST レコードで定義された共有文字列テーブル
// BIFF8 では文字列セルは文字列を直接持たず、このテーブルへの参照番号だけを持つ
class SharedStringTable {
public:
    // parse(): SST レコードと続く CONTINUE レコードを合わせてパースする
    // SST/CONTINUE は文字列の途中で分割されることがあるため合わせて処理する
    void parse(const BiffRecord& sst_record,
               const std::vector<BiffRecord>& continues);
    // get(): インデックスに対応する文字列を返す (範囲外なら RangeError)
    const std::string& get(uint32_t index) const;
    // size(): テーブルに登録された文字列の数を返す
    size_t size() const { return strings_.size(); }
private:
    // strings_: 登録済み文字列のリスト (インデックス = 位置)
    std::vector<std::string> strings_;
};

// FormatRegistry: 数値フォーマット (日付・通貨・パーセントなど) を管理するクラス
// BIFF8 では組み込みフォーマット (ID 0〜49) とカスタムフォーマットがある
class FormatRegistry {
public:
    // コンストラクタ: 組み込みフォーマット (ID 0〜49) を formats_ に登録する
    FormatRegistry();
    // add(): FORMAT レコードで定義されたカスタムフォーマットを登録する
    void       add(uint16_t index, const std::string& fmt_string);
    // get(): フォーマット ID に対応する FormatInfo を返す (未知 ID は空文字列)
    FormatInfo get(uint16_t index) const;
    // is_date_format(): フォーマット ID が日付書式か判定する
    bool       is_date_format(uint16_t index) const;
private:
    // formats_: ID → フォーマット文字列のマップ
    std::unordered_map<uint16_t, std::string> formats_;
    // looks_like_date(): フォーマット文字列に日付記号が含まれるか調べる
    static bool looks_like_date(const std::string& fmt);
};

// XFRecord: 1 つの XF (eXtended Format) レコードの内容
// セルの見た目 (フォント, 数値書式, 罫線 etc.) を定義する
struct XFRecord {
    uint16_t font_index  {0};    // フォントテーブルへのインデックス
    uint16_t format_index{0};    // 数値フォーマット ID (FormatRegistry と対応)
    bool     is_style_xf {false}; // true なら「スタイル XF」(セル XF ではない)
};

// XFTable: XF レコードのリストを管理するクラス
class XFTable {
public:
    // add(): 新しい XFRecord を末尾に追加する
    void             add(const XFRecord& xf);
    // get(): インデックスに対応する XFRecord を返す (範囲外なら RangeError)
    const XFRecord&  get(uint16_t index) const;
    // size(): テーブルのサイズを返す
    size_t           size() const { return records_.size(); }
private:
    // records_: XFRecord のリスト (インデックス = 位置)
    std::vector<XFRecord> records_;
};

// ---- シート ----
// Sheet 抽象クラスを継承した XLS 専用のシート実装
class XlsSheet : public Sheet {
public:
    // コンストラクタ: シート名を受け取り name_ に格納する
    explicit XlsSheet(std::string name) : name_(std::move(name)) {}

    // ---- Sheet 仮想関数のオーバーライド ----
    std::string name()      const override { return name_; }         // シート名
    uint32_t    row_count() const override { return row_count_; }    // 行数
    uint32_t    col_count() const override { return col_count_; }    // 列数

    // cell(): 行・列インデックスでセルを取得 (空セルは Blank を返す)
    Cell                cell(uint32_t row, uint32_t col)  const override;
    // cell(): CellAddress 構造体でセルを取得
    Cell                cell(const CellAddress& addr)      const override;
    // cell(): "A1" 記法の文字列でセルを取得
    Cell                cell(const std::string& a1)        const override;
    // try_cell(): 空セルなら nullopt を返す安全版
    std::optional<Cell> try_cell(uint32_t row, uint32_t col) const override;
    // row(): 指定行の全セルを vector で返す
    std::vector<Cell>   row(uint32_t row_idx)              const override;
    // cells(): シート内の全非空セルを vector で返す
    std::vector<Cell>   cells()                            const override;
    // for_each_cell(): 全セルを fn に渡して走査するイテレータ
    void for_each_cell(std::function<void(const Cell&)> fn) const override;

    // set_cell(): セルに値を書き込む (行・列インデックス版)
    void set_cell   (uint32_t row, uint32_t col, const CellValue& v) override;
    // set_cell(): "A1" 記法でセルに値を書き込む
    void set_cell   (const std::string& a1,      const CellValue& v) override;
    // set_formula(): セルに数式を設定する
    void set_formula(const std::string& a1, const std::string& formula) override;

    // put_cell(): パーサーが解析済みセルを直接格納するための内部 API
    void put_cell       (const Cell& c);
    // set_dimensions(): DIMENSION レコードから行数・列数を設定する
    void set_dimensions (uint32_t rows, uint32_t cols);

private:
    std::string name_;           // シート名
    uint32_t    row_count_{0};   // 最大行数 (0-indexed + 1)
    uint32_t    col_count_{0};   // 最大列数 (0-indexed + 1)
    // data_: 行番号 → (列番号 → Cell) の二重マップ
    std::map<uint32_t, std::map<uint32_t, Cell>> data_;

    // fill_type(): セル値の型 (CellValue の variant 型) から CellType を設定する
    static void fill_type(Cell& c);
};

// ---- ワークブック ----
// Workbook 抽象クラスを継承した XLS 専用のワークブック実装
class XlsWorkbook : public Workbook {
public:
    // ---- Workbook 仮想関数のオーバーライド ----
    FileFormat   format()      const override { return FileFormat::XLS; }  // フォーマット種別
    std::string  format_name() const override { return "XLS"; }             // フォーマット名
    size_t       sheet_count() const override { return sheets_.size(); }    // シート数

    // sheet(): インデックスでシートを取得 (const/非 const 両方)
    Sheet&       sheet(size_t index)            override;
    const Sheet& sheet(size_t index)      const override;
    // sheet(): シート名でシートを取得
    Sheet&       sheet(const std::string& name) override;
    // sheet_names(): 全シート名のリストを返す
    std::vector<std::string> sheet_names() const override;

    // add_sheet(): 新しいシートを末尾に追加して参照を返す
    Sheet& add_sheet   (const std::string& name)              override;
    // remove_sheet(): インデックス指定でシートを削除する
    void   remove_sheet(size_t index)                         override;
    // rename_sheet(): シート名を変更する (重複名は WriteError)
    void   rename_sheet(size_t index, const std::string& name)override;

    // save(): ファイルパスに保存する (XLS の書き込みは XLSX 形式にフォールバック)
    void                 save    (const std::string& path,
                                  const SaveOptions& opts = {}) const override;
    // to_bytes(): バイト列を生成して返す
    std::vector<uint8_t> to_bytes(FileFormat fmt,
                                  const SaveOptions& opts = {}) const override;

    // add_parsed_sheet(): パーサーが解析済みシートを追加するための内部 API
    void add_parsed_sheet(std::unique_ptr<XlsSheet> s);
private:
    // sheets_: XlsSheet の unique_ptr を保持するリスト
    std::vector<std::unique_ptr<XlsSheet>> sheets_;
};

// ---- パーサー ----
// XLS (OLE2 + BIFF8) ファイルを解析して XlsWorkbook を生成するクラス
class XlsParser {
public:
    // コンストラクタ: オープンオプションを受け取る
    explicit XlsParser(const OpenOptions& opts) : opts_(opts) {}
    // parse(): XLS バイト列を解析して XlsWorkbook を返す
    std::unique_ptr<XlsWorkbook> parse(const std::vector<uint8_t>& data);

private:
    OpenOptions opts_;   // オープンオプション (strict_validation など)

    // extract_workbook_stream(): OLE2 コンテナから BIFF8 ストリームを取り出す
    // OLE2 は FAT/MiniFAT チェーンでデータブロックを管理する複合ドキュメント形式
    std::vector<uint8_t> extract_workbook_stream(const std::vector<uint8_t>& raw);
    // read_records(): BIFF8 ストリームを BiffRecord のリストに変換する
    std::vector<BiffRecord> read_records(const std::vector<uint8_t>& stream);

    // GlobalContext: ワークブックレベルのグローバル情報をまとめた構造体
    // 2 パス解析で使う: 1 パス目でこれを収集し、2 パス目のシートパースで参照する
    struct GlobalContext {
        SharedStringTable sst;     // 共有文字列テーブル (LABELSST セルが参照)
        XFTable           xf_table; // XF テーブル (セルのフォーマット情報)
        FormatRegistry    formats;  // 数値フォーマットレジストリ
        // sheets: (シート名, BOF オフセット) のペアのリスト
        // BOF オフセット = BIFF8 ストリーム内でシートデータが始まる位置
        std::vector<std::pair<std::string, uint32_t>> sheets;
    };

    // parse_globals(): レコードリストからグローバルコンテキストを収集する (1 パス目)
    GlobalContext parse_globals(const std::vector<BiffRecord>& records);

    // xf_to_style(): XF インデックスから CellStyle を生成する
    CellStyle xf_to_style(uint16_t xf_index, const XFTable& xf,
                           const FormatRegistry& fmt);

    // ---- セルレコードハンドラー (2 パス目で呼ばれる) ----
    // handle_number(): NUMBER レコード (8バイト浮動小数点数セル) を Cell に変換
    Cell handle_number  (const BiffRecord& r, const GlobalContext& ctx);
    // handle_labelsst(): LABELSST レコード (共有文字列参照セル) を Cell に変換
    Cell handle_labelsst(const BiffRecord& r, const GlobalContext& ctx);
    // handle_label(): LABEL レコード (直接文字列格納セル) を Cell に変換
    Cell handle_label   (const BiffRecord& r);
    // handle_boolerr(): BOOLERR レコード (真偽値またはエラー値セル) を Cell に変換
    Cell handle_boolerr (const BiffRecord& r);
    // handle_rk(): RK レコード (圧縮数値セル) を Cell に変換
    Cell handle_rk      (const BiffRecord& r, const GlobalContext& ctx);
    // handle_mulrk(): MULRK レコード (複数 RK 値が連続するセル群) を Cell のリストに変換
    std::vector<Cell> handle_mulrk(const BiffRecord& r, const GlobalContext& ctx);
    // handle_formula(): FORMULA レコード (数式セル) を Cell に変換
    Cell handle_formula (const BiffRecord& r, const GlobalContext& ctx);
};

} // namespace excellib::xls
