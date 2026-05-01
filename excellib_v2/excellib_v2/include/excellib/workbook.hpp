#pragma once
/**
 * excellib/workbook.hpp
 * Core types: CellValue, Cell, Sheet, Workbook, WorkbookFactory
 */

// std::string: 文字列型
#include <string>
// std::vector: 動的配列型
#include <vector>
// std::unique_ptr / std::shared_ptr: スマートポインタ型
#include <memory>
// std::variant: 複数の型のいずれかを保持できる型
#include <variant>
// std::optional: 値があるかもしれない型 (nullopt で「値なし」を表現)
#include <optional>
// std::runtime_error など標準例外クラス
#include <stdexcept>
// uint32_t / int64_t など固定幅整数型
#include <cstdint>
// std::function: 関数オブジェクト・ラムダを保持できる型
#include <functional>

// excellib ライブラリ全体を囲む名前空間
namespace excellib {

// ============================================================
//  Error hierarchy (例外クラスの階層)
// ============================================================
// ExcelError: すべての excellib 例外の基底クラス
// using ... = ... で基底クラスのコンストラクタを継承
struct ExcelError  : std::runtime_error { using std::runtime_error::runtime_error; };
// ParseError: ファイルのパース(解析)に失敗したときに投げる
struct ParseError  : ExcelError         { using ExcelError::ExcelError; };
// FormatError: ファイルフォーマットが不正・未対応のときに投げる
struct FormatError : ExcelError         { using ExcelError::ExcelError; };
// IOError: ファイルの読み書きに失敗したときに投げる
struct IOError     : ExcelError         { using ExcelError::ExcelError; };
// RangeError: シート・行・列のインデックスが範囲外のときに投げる
struct RangeError  : ExcelError         { using ExcelError::ExcelError; };
// WriteError: 書き込み操作に問題があるとき (重複シート名など) に投げる
struct WriteError  : ExcelError         { using ExcelError::ExcelError; };

// ============================================================
//  CellValue — strict typed variant, no silent coercion
//  (厳密な型付きバリアント。暗黙の型変換は行わない)
// ============================================================
// BlankValue: 空セル(何も入っていないセル)を表すタグ型
struct BlankValue {};
// ErrorValue: Excel のエラー値 (#REF! / #DIV/0! など) を表す型
struct ErrorValue { std::string code; };   // "#REF!", "#DIV/0!", etc.

// CellValue: セルの値を表す variant 型
// std::variant は「いずれかの型」を保持できる型。union の型安全版
// Excel のセルは空・真偽・整数・小数・文字列・エラーのいずれかの値を持つ
using CellValue = std::variant<
    BlankValue,    // 空セル
    bool,          // TRUE / FALSE
    int64_t,       // 整数 (64ビット符号付き整数)
    double,        // 浮動小数点数
    std::string,   // 文字列
    ErrorValue     // エラー値 (#REF! など)
>;

// ---- Type predicates (型判定関数) ----
// std::holds_alternative<T>(v) で v が型 T を保持しているか確認する
// inline はこの関数をヘッダーに直接定義できるようにするキーワード
inline bool is_blank  (const CellValue& v) { return std::holds_alternative<BlankValue>(v); }
inline bool is_bool   (const CellValue& v) { return std::holds_alternative<bool>(v); }
inline bool is_int    (const CellValue& v) { return std::holds_alternative<int64_t>(v); }
inline bool is_double (const CellValue& v) { return std::holds_alternative<double>(v); }
inline bool is_string (const CellValue& v) { return std::holds_alternative<std::string>(v); }
inline bool is_error  (const CellValue& v) { return std::holds_alternative<ErrorValue>(v); }
// is_numeric: int64_t または double のどちらかなら true
inline bool is_numeric(const CellValue& v) { return is_int(v) || is_double(v); }

// ---- Strict getters (型ミスマッチ時は std::bad_variant_access を投げる) ----
// std::get<T>(v) で v から型 T の値を取り出す
// 型が合わない場合は std::bad_variant_access 例外が発生する
inline bool        get_bool  (const CellValue& v) { return std::get<bool>(v); }
inline int64_t     get_int   (const CellValue& v) { return std::get<int64_t>(v); }
inline double      get_double(const CellValue& v) { return std::get<double>(v); }
inline std::string get_string(const CellValue& v) { return std::get<std::string>(v); }
inline ErrorValue  get_error (const CellValue& v) { return std::get<ErrorValue>(v); }

// ---- Numeric coercion (明示的な数値変換) ----
// as_double() は int64_t でも double でも double として返す
// 数値以外が渡された場合は FormatError を投げる
inline double as_double(const CellValue& v) {
    // 整数なら double にキャストして返す
    if (is_int(v))    return static_cast<double>(get_int(v));
    // double ならそのまま返す
    if (is_double(v)) return get_double(v);
    // 数値でない型が来たら例外を投げる
    throw FormatError("Cell value is not numeric");
}

// ---- Convert any non-error value to string (任意の値を文字列に変換) ----
// 宣言のみ (実装は common.cpp)
std::string to_string(const CellValue& v);

// ============================================================
//  CellAddress (0-based, row & col)
//  (行・列は 0-indexed。Excel の "A1" は row=0, col=0)
// ============================================================
struct CellAddress {
    // row: 行番号 (0始まり)  col: 列番号 (0始まり)
    uint32_t row{0};    // {0} はデフォルト値の初期化子
    uint32_t col{0};

    // == 演算子のオーバーロード: row と col が両方同じなら等しいと判定
    bool operator==(const CellAddress& o) const { return row==o.row && col==o.col; }

    /// Parse "B3" → {2, 1}.  Throws FormatError on invalid input.
    /// "B3" のような A1 記法の文字列をパースして行・列インデックスに変換する
    /// 不正な入力 ("A0", "" など) は FormatError を投げる
    static CellAddress from_a1(const std::string& a1);

    /// Convert back to "B3" form.
    /// 行・列インデックスを "B3" のような A1 記法の文字列に変換する
    std::string to_a1() const;
};

// ============================================================
//  Cell metadata (セルのメタデータ)
// ============================================================
// CellType: セルのデータ種別を表す列挙型
enum class CellType { Blank, Number, String, Boolean, Error, Formula };

// FormatInfo: 数値フォーマット情報 (日付・パーセントなどの書式)
struct FormatInfo {
    // format_index: XLSX/XLS の数値フォーマットID (例: 14 = 日付)
    uint16_t    format_index{0};
    // format_string: フォーマット文字列 (例: "YYYY-MM-DD", "0.00%")
    std::string format_string;      // e.g. "YYYY-MM-DD", "0.00%"
    // is_date_format: 日付フォーマットかどうか
    bool        is_date_format{false};
    // is_time_format: 時刻フォーマットかどうか
    bool        is_time_format{false};
};

// CellStyle: セルのスタイル情報 (書式を持つかどうか)
struct CellStyle {
    // format: 書式情報 (optional なので「書式なし」= std::nullopt も表現できる)
    std::optional<FormatInfo> format;
};

// Cell: セル1つ分のすべての情報をまとめた構造体
struct Cell {
    // address: このセルのアドレス (行・列インデックス)
    CellAddress                address;
    // type: セルのデータ種別 (デフォルトは空)
    CellType                   type{CellType::Blank};
    // value: セルの実際の値 (variant 型)
    CellValue                  value;
    // formula: 数式文字列 (optional: 数式がない場合は nullopt)
    std::optional<std::string> formula;   // raw formula string if present
    // style: スタイル情報 (書式など)
    CellStyle                  style;
    // merged: このセルが結合セルかどうか
    bool                       merged{false};

    // is_blank(): セルが空かどうかを返すメンバ関数
    bool is_blank()    const { return type == CellType::Blank; }
    // has_formula(): セルに数式があるかどうかを返すメンバ関数
    bool has_formula() const { return formula.has_value(); }
};

// ============================================================
//  Sheet interface (シートの抽象インターフェース)
// ============================================================
// class Sheet は純粋仮想関数を持つ抽象クラス (インターフェース)
// XlsxSheet や XlsSheet が具体的な実装を提供する
class Sheet {
public:
    // virtual ~Sheet() = default: 派生クラスのデストラクタが正しく呼ばれるようにする
    virtual ~Sheet() = default;

    // シート名・行数・列数を取得する純粋仮想関数 (= 0 で純粋仮想関数を宣言)
    // 派生クラスで必ず実装しなければならない
    virtual std::string name()      const = 0;
    virtual uint32_t    row_count() const = 0;
    virtual uint32_t    col_count() const = 0;

    // ---- Read (読み取り系) ----
    /// Returns a blank Cell (not an error) if the address is empty or out of range.
    /// アドレスが空または範囲外の場合は空の Cell を返す (例外は投げない)
    virtual Cell cell(uint32_t row, uint32_t col)  const = 0;    // 行・列インデックスで指定
    virtual Cell cell(const CellAddress& addr)      const = 0;    // CellAddress 構造体で指定
    virtual Cell cell(const std::string& a1)        const = 0;    // A1 記法の文字列で指定

    /// Returns nullopt for blank or out-of-range cells.
    /// 空・範囲外のセルは nullopt を返す (例外が発生しないのが利点)
    virtual std::optional<Cell> try_cell(uint32_t row, uint32_t col) const = 0;

    /// All cells in one row (sorted by column).
    /// 指定した行の全セルを列順に並べたベクタで返す
    virtual std::vector<Cell> row(uint32_t row_idx) const = 0;

    /// All non-blank cells (row-major order).
    /// 非空セルをすべて行優先順で返す
    virtual std::vector<Cell> cells() const = 0;

    // for_each_cell: 全非空セルに対してコールバック関数を呼ぶ
    // std::function<void(const Cell&)> はセルを受け取って何も返さない関数の型
    virtual void for_each_cell(std::function<void(const Cell&)> fn) const = 0;

    // ---- Write (書き込み系) ----
    // 行・列インデックスで値を書き込む
    virtual void set_cell   (uint32_t row, uint32_t col, const CellValue& v) = 0;
    // A1 記法で値を書き込む
    virtual void set_cell   (const std::string& a1,      const CellValue& v) = 0;
    // A1 記法で数式文字列を書き込む (例: "SUM(A1:A10)")
    virtual void set_formula(const std::string& a1, const std::string& formula) = 0;

    // ---- Bulk helpers (一括書き込みヘルパー) ----
    /// Write a 2-D table starting at top_left.  Each inner vector is one row.
    /// top_left を起点に 2次元ベクタのデータを一括書き込みする
    /// 外側のベクタが行、内側のベクタが列に対応する
    void set_table(const std::string& top_left,
                   const std::vector<std::vector<CellValue>>& data);
};

// ============================================================
//  Workbook interface (ワークブックの抽象インターフェース)
// ============================================================
// FileFormat: サポートするファイルフォーマット
enum class FileFormat { XLS, XLSX, Auto };

// OpenOptions: ファイルを開く際のオプション
struct OpenOptions {
    // strict_validation: true の場合、フォーマット違反があれば例外を投げる
    bool strict_validation  {true};   // throw on any format violation
    // preserve_formulas: 数式を保持するかどうか
    bool preserve_formulas  {true};
    // preserve_styles: スタイル(書式)を保持するかどうか
    bool preserve_styles    {true};
    // fail_on_unsupported: 未対応の機能があれば例外を投げるかどうか
    bool fail_on_unsupported{false};  // throw on unrecognised features
};

// SaveOptions: ファイルを保存する際のオプション
struct SaveOptions {
    // format: 保存するフォーマット (Auto = 元のフォーマットを維持)
    FileFormat format    {FileFormat::Auto}; // Auto = keep original format
    // recalculate: 保存時に数式を再計算するかどうか
    bool       recalculate{false};
};

// Workbook: ワークブック(Excelファイル)を表す抽象クラス
class Workbook {
public:
    // デストラクタ: デフォルト実装を使う
    virtual ~Workbook() = default;

    // ファイルフォーマット (XLS or XLSX) を返す
    virtual FileFormat   format()      const = 0;
    // フォーマット名の文字列 ("XLS" or "XLSX") を返す
    virtual std::string  format_name() const = 0;
    // シートの総数を返す
    virtual size_t       sheet_count() const = 0;

    // インデックスでシートを取得 (const と非constの両バージョン)
    virtual Sheet&       sheet(size_t index)             = 0;
    virtual const Sheet& sheet(size_t index)       const = 0;
    // シート名でシートを取得
    virtual Sheet&       sheet(const std::string& name)  = 0;
    // 全シート名をベクタで返す
    virtual std::vector<std::string> sheet_names() const = 0;

    // 新しいシートを追加し参照を返す
    virtual Sheet& add_sheet   (const std::string& name)              = 0;
    // インデックスでシートを削除する
    virtual void   remove_sheet(size_t index)                         = 0;
    // インデックスでシートをリネームする
    virtual void   rename_sheet(size_t index, const std::string& name)= 0;

    // ファイルに保存する (opts はデフォルト値 = {} で省略可能)
    virtual void                 save    (const std::string& path,
                                          const SaveOptions& opts = {}) const = 0;
    // バイト列として取得する (ファイルに書かずメモリ上で保持したい場合に使う)
    virtual std::vector<uint8_t> to_bytes(FileFormat fmt,
                                          const SaveOptions& opts = {}) const = 0;
};

// ============================================================
//  WorkbookFactory (ファクトリクラス: ワークブックの生成を担当)
// ============================================================
class WorkbookFactory {
public:
    /// Open from file path.  Format is detected from magic bytes, never from extension.
    /// ファイルパスからワークブックを開く。フォーマットは拡張子でなくマジックバイトで判定
    static std::unique_ptr<Workbook> open(const std::string& path,
                                          const OpenOptions& opts = {});

    /// Open from an in-memory buffer.
    /// メモリ上のバイト列からワークブックを開く (HTTP レスポンスなど)
    static std::unique_ptr<Workbook> open(const std::vector<uint8_t>& data,
                                          const OpenOptions& opts = {});

    /// Create a new empty workbook.
    /// 新規の空ワークブックを作成する (デフォルトは XLSX 形式)
    static std::unique_ptr<Workbook> create(FileFormat fmt = FileFormat::XLSX);
};

} // namespace excellib
