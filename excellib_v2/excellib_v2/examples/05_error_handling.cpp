/**
 * examples/05_error_handling.cpp
 *
 * エラーハンドリングのパターン集。
 * excellib が送出する各例外の種類と、セルアクセス時の安全なパターンを示す。
 *
 * 実行:
 *   ./05_error_handling
 */
// excellib.hpp: WorkbookFactory, CellValue, 例外クラスなど全クラスを一括インクルード
#include "excellib/excellib.hpp"
// std::cout: 標準出力
#include <iostream>

// excellib 名前空間を使用する宣言
using namespace excellib;

// ============================================================
//  例 1: 例外の種類別ハンドリング
// ============================================================
// excellib が送出しうる例外クラスの階層:
//   ExcelError (基底クラス)
//     ├─ FormatError  : ファイル形式が不正・未対応
//     ├─ ParseError   : バイト列の解析中に異常を検出
//     ├─ RangeError   : シート/セルのインデックスが範囲外
//     ├─ WriteError   : 書き込み時の制約違反 (重複シート名など)
//     └─ IOError      : ファイル読み書きのシステムエラー
// std::bad_variant_access は C++ 標準例外であり ExcelError とは別系統
void example_exceptions() {
    std::cout << "=== 例外ハンドリング ===\n";

    // ---- ケース 1: 壊れたファイル (FormatError または ParseError) ----
    try {
        // ゴミデータ (100 バイトすべて 0xFF) を渡す
        // WorkbookFactory::open() はマジックバイトで ZIP / OLE2 を判定するが
        // 0xFF の連続はどちらにも一致しないため FormatError が送出される
        std::vector<uint8_t> garbage(100, 0xFF);
        WorkbookFactory::open(garbage);
    } catch (const FormatError& e) {
        // FormatError を捕捉: e.what() でエラーメッセージを取得できる
        std::cout << "FormatError: " << e.what() << "\n";
    }

    // ---- ケース 2: 存在しないシートへのアクセス (RangeError) ----
    try {
        // 新規ワークブックにはシートが 0 枚 (add_sheet 前) なので
        // インデックス 99 を指定すると範囲外 → RangeError が送出される
        auto wb = WorkbookFactory::create();
        wb->sheet(99);   // sheet() は 0-indexed; 範囲外なら RangeError
    } catch (const RangeError& e) {
        std::cout << "RangeError: " << e.what() << "\n";
    }

    // ---- ケース 3: 重複したシート名の追加 (WriteError) ----
    try {
        auto wb = WorkbookFactory::create();
        wb->add_sheet("S");   // 1 枚目: OK
        wb->add_sheet("S");   // 2 枚目: 同名のシートは作れない → WriteError
    } catch (const WriteError& e) {
        std::cout << "WriteError: " << e.what() << "\n";
    }

    // ---- ケース 4: CellValue の型ミスマッチ (std::bad_variant_access) ----
    try {
        // CellValue は std::variant<BlankValue, bool, int64_t, double, std::string, ErrorValue>
        // v には std::string{"hello"} が格納されている
        CellValue v = std::string{"hello"};
        // get_int() は std::get<int64_t>(v) を呼び出す
        // v の実際の型は std::string なので std::bad_variant_access が送出される
        get_int(v);
    } catch (const std::bad_variant_access&) {
        // std::bad_variant_access は C++ 標準ライブラリの例外クラス
        // ExcelError の派生ではないため別の catch ブロックが必要
        std::cout << "bad_variant_access: 型が違う\n";
    }

    // ---- ケース 5: 全例外を一括で受け取る (ExcelError の基底クラス) ----
    try {
        // 存在しないファイルを開こうとすると IOError が送出される
        // IOError は ExcelError の派生クラスなので ExcelError& でまとめて捕捉できる
        auto wb = WorkbookFactory::open("存在しないファイル.xlsx");
    } catch (const ExcelError& e) {
        // ExcelError は FormatError / ParseError / RangeError / WriteError / IOError
        // 全ての基底クラス。個別の種類を問わずここで受け取れる
        std::cout << "ExcelError: " << e.what() << "\n";
    }
}

// ============================================================
//  例 2: 安全なセルアクセスパターン
// ============================================================
// CellValue に格納された値を安全に取り出すための 3 つのパターンを示す。
void example_safe_access() {
    std::cout << "\n=== 安全なセルアクセス ===\n";

    // テスト用のワークブックとシートを準備する
    auto wb = WorkbookFactory::create();
    auto& sh = wb->add_sheet("S");
    // A1: 文字列, B1: 整数, C1: 浮動小数点
    sh.set_cell("A1", std::string{"hello"});
    sh.set_cell("B1", int64_t{42});
    sh.set_cell("C1", 3.14);

    // ---- パターン 1: is_* で事前チェック (最も安全で明示的) ----
    // is_string() / is_int() / is_numeric() など型を確認してから get_*() を呼ぶ
    // 型が一致しない場合でも例外が発生しないため、信頼度が低いデータに向く
    Cell c = sh.cell("A1");   // A1 セルの Cell 構造体を取得
    if (is_string(c.value))
        // get_string(): std::get<std::string>(v) のラッパー (型保証済みのときに使う)
        std::cout << "string: " << get_string(c.value) << "\n";
    else if (is_int(c.value))
        // get_int(): std::get<int64_t>(v) のラッパー
        std::cout << "int: "    << get_int(c.value) << "\n";
    else if (is_numeric(c.value))
        // is_numeric(): is_int() || is_double() の合成判定
        // as_double(): int64_t/double どちらも double として返す変換関数
        std::cout << "number: " << as_double(c.value) << "\n";

    // ---- パターン 2: try_cell で空セルを安全に処理 ----
    // cell() は空セルにアクセスすると RangeError を送出することがある
    // try_cell() は空セルのときに std::nullopt を返すため例外が発生しない
    if (auto opt = sh.try_cell(0, 0)) {   // row=0 (1行目), col=0 (A列) → A1
        // opt は std::optional<Cell> なのでアクセスには opt-> / *opt を使う
        std::cout << "A1 exists: " << to_string(opt->value) << "\n";
    }
    if (!sh.try_cell(99, 99)) {           // row=99 (100行目), col=99 (CV列) → 空セル
        // try_cell() が nullopt を返すと bool 変換で false になる
        std::cout << "Z100 is empty\n";
    }

    // ---- パターン 3: std::visit で全型を網羅する ----
    // std::visit は variant の実際の型に応じてラムダを呼び分けるディスパッチ機構
    // すべての型バリアントを if constexpr で網羅するため、型追加時にコンパイルエラーで検知できる
    sh.for_each_cell([](const Cell& c) {
        // std::visit(visitor, variant): variant の実際の型で visitor を呼び出す
        std::visit([&](auto&& v) {
            // std::decay_t<decltype(v)>: 参照・const 修飾を除いた実際の型を取得する
            // これにより T == int64_t, T == double, T == std::string などの比較ができる
            using T = std::decay_t<decltype(v)>;
            std::cout << c.address.to_a1() << ": ";
            // if constexpr: コンパイル時に型を評価するため、
            // 実行時のオーバーヘッドが一切ない (通常の if-else より高速)
            if      constexpr (std::is_same_v<T, BlankValue>)   std::cout << "(blank)";
            // bool は int64_t より先に比較する (bool は数値型に暗黙変換されるため)
            else if constexpr (std::is_same_v<T, bool>)         std::cout << (v ? "TRUE" : "FALSE");
            else if constexpr (std::is_same_v<T, int64_t>)      std::cout << v;
            else if constexpr (std::is_same_v<T, double>)       std::cout << v;
            // std::string は引用符で囲んで出力する
            else if constexpr (std::is_same_v<T, std::string>)  std::cout << '"' << v << '"';
            // ErrorValue: Excel のセルエラー (#DIV/0!, #N/A, #REF! など)
            // v.code にエラーコード文字列が格納されている
            else if constexpr (std::is_same_v<T, ErrorValue>)   std::cout << v.code;
            std::cout << "\n";
        }, c.value);
    });
}

// ============================================================
//  例 3: 数値の集計
// ============================================================
// try_cell() と is_numeric() / as_double() を組み合わせて
// 型が混在する列から数値だけを安全に集計するパターン。
void example_aggregate() {
    std::cout << "\n=== 数値集計 ===\n";

    auto wb = WorkbookFactory::create();
    auto& sh = wb->add_sheet("Data");
    // ヘッダー行 + 4 行のデータを set_table() で書き込む
    sh.set_table("A1", {
        // ヘッダー行: 文字列 (集計対象外)
        {std::string{"Name"}, std::string{"Score"}},
        // データ行: 名前 (文字列) と点数 (int64_t)
        {std::string{"Alice"}, int64_t{95}},
        {std::string{"Bob"},   int64_t{87}},
        {std::string{"Carol"}, int64_t{92}},
        {std::string{"Dave"},  int64_t{78}},
    });

    // ---- B 列 (スコア列) の合計と平均を求める ----
    double sum = 0;   // 合計値の累積変数
    int count = 0;    // 有効な数値セルの個数

    // r = 1 から始めてヘッダー行 (r=0) をスキップする
    // row_count(): 有効データのある行数。ここでは 5 (ヘッダー 1 + データ 4)
    for (uint32_t r = 1; r < sh.row_count(); ++r) {
        // try_cell(r, 1): B 列 (col=1) のセルを取得。空なら nullopt
        if (auto c = sh.try_cell(r, 1)) {   // B列: スコア
            // is_numeric(): int64_t または double ならば true
            // ヘッダー行のような文字列セルはここで除外される
            if (is_numeric(c->value)) {
                // as_double(): int64_t/double を double に変換して加算
                sum += as_double(c->value);
                ++count;
            }
        }
    }
    // ゼロ除算を避けるため count > 0 のときだけ平均を計算する
    if (count > 0)
        std::cout << "合計: " << sum << " / 平均: " << sum/count << "\n";
}

int main() {
    // 3 つのエラーハンドリングパターンをそれぞれ実行する
    example_exceptions();   // 例外の種類別ハンドリング
    example_safe_access();  // 安全なセルアクセス
    example_aggregate();    // 数値集計
    return 0;
}
