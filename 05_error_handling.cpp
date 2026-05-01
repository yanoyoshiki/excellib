/**
 * examples/05_error_handling.cpp
 *
 * エラーハンドリングのパターン集。
 *
 * 実行:
 *   ./05_error_handling
 */

// excellib の全機能をインクルード
#include "excellib/excellib.hpp"
// std::cout など標準入出力を使うため
#include <iostream>

// excellib:: を省略できるようにする
using namespace excellib;

// ---- 例 1: 例外の種類別ハンドリング ----
// void 関数は戻り値なし
void example_exceptions() {
    std::cout << "=== 例外ハンドリング ===\n";

    // 壊れたファイル → ParseError or FormatError
    try {
        // 100バイト分 0xFF を詰めた「ゴミデータ」を作成
        // std::vector<uint8_t>(100, 0xFF) は要素数100、全て 0xFF の配列
        std::vector<uint8_t> garbage(100, 0xFF);
        // このゴミデータをファイルとして開こうとすると FormatError が発生する
        // (マジックバイトが OLE2 にも ZIP にも一致しないため)
        WorkbookFactory::open(garbage);
    // FormatError は ExcelError の派生クラス (workbook.hpp で定義)
    // フォーマットが認識できないときに投げられる
    } catch (const FormatError& e) {
        std::cout << "FormatError: " << e.what() << "\n";
    }

    // 存在しないシート → RangeError
    try {
        // 空のワークブックを作成
        auto wb = WorkbookFactory::create();
        // シートが0枚なのに 99番目にアクセスしようとすると RangeError が発生
        wb->sheet(99);
    // RangeError は範囲外アクセス(シートや行・列のインデックスが存在しない)のとき
    } catch (const RangeError& e) {
        std::cout << "RangeError: " << e.what() << "\n";
    }

    // 重複シート名 → WriteError
    try {
        auto wb = WorkbookFactory::create();
        // "S" という名前のシートを1枚追加
        wb->add_sheet("S");
        // 同じ名前 "S" でもう1枚追加しようとすると WriteError が発生
        wb->add_sheet("S");
    // WriteError は書き込み操作に問題がある場合 (重複シート名など)
    } catch (const WriteError& e) {
        std::cout << "WriteError: " << e.what() << "\n";
    }

    // 型ミスマッチ → std::bad_variant_access
    try {
        // CellValue に文字列 "hello" を格納
        CellValue v = std::string{"hello"};
        // 文字列なのに get_int() で整数として取り出そうとすると例外が発生
        // CellValue は std::variant なので間違った型で get すると bad_variant_access
        get_int(v);
    // std::bad_variant_access は標準ライブラリの例外 (<variant> で定義)
    } catch (const std::bad_variant_access&) {
        std::cout << "bad_variant_access: 型が違う\n";
    }

    // 全例外を一括で受け取る場合
    // ExcelError はすべての excellib 例外の基底クラス
    // → FormatError, ParseError, IOError, RangeError, WriteError をすべて受け取れる
    try {
        // 存在しないファイルを開こうとすると IOError が発生
        auto wb = WorkbookFactory::open("存在しないファイル.xlsx");
    } catch (const ExcelError& e) {
        // e.what() でエラーメッセージを文字列として取得
        std::cout << "ExcelError: " << e.what() << "\n";
    }
}

// ---- 例 2: 安全なセルアクセスパターン ----
void example_safe_access() {
    std::cout << "\n=== 安全なセルアクセス ===\n";

    // テスト用のワークブックとシートを作成
    auto wb = WorkbookFactory::create();
    // "S" という名前のシートを追加し参照を取得
    auto& sh = wb->add_sheet("S");
    // A1 に文字列 "hello" を書き込む
    sh.set_cell("A1", std::string{"hello"});
    // B1 に整数 42 を書き込む
    sh.set_cell("B1", int64_t{42});
    // C1 に小数 3.14 を書き込む
    sh.set_cell("C1", 3.14);

    // ---- パターン 1: is_* で事前チェック（最も安全）----
    // A1 の Cell オブジェクトを取得
    Cell c = sh.cell("A1");
    // is_string() で文字列かどうか確認してから get_string() で取り出す
    if (is_string(c.value))      std::cout << "string: " << get_string(c.value) << "\n";
    // is_int() で整数 (int64_t) かどうか確認してから get_int() で取り出す
    else if (is_int(c.value))    std::cout << "int: "    << get_int(c.value) << "\n";
    // is_numeric() は int64_t または double のいずれかなら true
    // as_double() でどちらの数値型でも double に変換して取得
    else if (is_numeric(c.value))std::cout << "number: " << as_double(c.value) << "\n";

    // ---- パターン 2: try_cell で空セルを安全に処理 ----
    // try_cell() は空・範囲外のセルを例外でなく nullopt で返す
    // (0, 0) は A1 に相当 (0-indexed)
    if (auto opt = sh.try_cell(0, 0)) {  // A1
        // opt.has_value() が true の場合だけここが実行される
        // opt-> で std::optional<Cell> の中身を取り出す
        std::cout << "A1 exists: " << to_string(opt->value) << "\n";
    }
    // 99行99列 (範囲外の空セル) は nullopt が返る
    if (!sh.try_cell(99, 99)) {          // 空のセル
        std::cout << "Z100 is empty\n";
    }

    // ---- パターン 3: std::visit で全型を網羅 ----
    // for_each_cell() で全非空セルを走査
    sh.for_each_cell([](const Cell& c) {
        // std::visit は variant (CellValue) の実際の型に応じてラムダを呼び出す
        // [&] で c をキャプチャ
        std::visit([&](auto&& v) {
            // std::decay_t は参照や const を取り除いた純粋な型を得るための型トレイト
            // decltype(v) は v の型, std::decay_t でその型名 T を得る
            using T = std::decay_t<decltype(v)>;
            // セルアドレスを A1 形式で表示
            std::cout << c.address.to_a1() << ": ";
            // if constexpr は「コンパイル時の条件分岐」。実行時ではなくコンパイル時に型を確定させる
            // std::is_same_v<T, X> で T が X と同じ型か判定
            if      constexpr (std::is_same_v<T, BlankValue>)   std::cout << "(blank)";
            // bool 型は TRUE/FALSE で表示
            else if constexpr (std::is_same_v<T, bool>)         std::cout << (v ? "TRUE" : "FALSE");
            // 整数型はそのまま表示
            else if constexpr (std::is_same_v<T, int64_t>)      std::cout << v;
            // 浮動小数点型はそのまま表示
            else if constexpr (std::is_same_v<T, double>)       std::cout << v;
            // 文字列型はダブルクォートで囲んで表示
            else if constexpr (std::is_same_v<T, std::string>)  std::cout << '"' << v << '"';
            // エラー値 (#REF! など) はコード文字列を表示
            else if constexpr (std::is_same_v<T, ErrorValue>)   std::cout << v.code;
            std::cout << "\n";
        }, c.value);    // c.value (CellValue = std::variant) を visit に渡す
    });
}

// ---- 例 3: 数値の集計 ----
void example_aggregate() {
    std::cout << "\n=== 数値集計 ===\n";

    // テスト用データを持つワークブックを作成
    auto wb = WorkbookFactory::create();
    // "Data" という名前のシートを追加
    auto& sh = wb->add_sheet("Data");
    // A1 起点にヘッダー + スコアデータを書き込む
    sh.set_table("A1", {
        // ヘッダー行
        {std::string{"Name"}, std::string{"Score"}},
        // データ行: 名前と点数
        {std::string{"Alice"}, int64_t{95}},
        {std::string{"Bob"},   int64_t{87}},
        {std::string{"Carol"}, int64_t{92}},
        {std::string{"Dave"},  int64_t{78}},
    });

    // 数値列（B列、1行目はヘッダーなのでスキップ）の合計と平均
    // double で合計を蓄積 (精度を確保するため)
    double sum = 0; int count = 0;
    // r=1 から始める (0行目はヘッダーなのでスキップ)
    for (uint32_t r = 1; r < sh.row_count(); ++r) {
        // B列は列インデックス 1 (0-indexed)
        if (auto c = sh.try_cell(r, 1)) {   // B列
            // セルが数値 (int64_t または double) かどうか確認
            if (is_numeric(c->value)) {
                // as_double() で統一的に double として加算
                sum += as_double(c->value);
                // カウントを1増やす
                ++count;
            }
        }
    }
    // ゼロ除算を防ぐために count > 0 を確認してから平均を計算
    if (count > 0)
        std::cout << "合計: " << sum << " / 平均: " << sum/count << "\n";
}

// プログラムのエントリーポイント
int main() {
    // 例外ハンドリングのデモを実行
    example_exceptions();
    // 安全なセルアクセスのデモを実行
    example_safe_access();
    // 数値集計のデモを実行
    example_aggregate();
    // 正常終了
    return 0;
}
