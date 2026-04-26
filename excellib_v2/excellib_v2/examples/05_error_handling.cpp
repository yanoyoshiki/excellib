/**
 * examples/05_error_handling.cpp
 *
 * エラーハンドリングのパターン集。
 *
 * 実行:
 *   ./05_error_handling
 */
#include "excellib/excellib.hpp"
#include <iostream>

using namespace excellib;

// ---- 例 1: 例外の種類別ハンドリング ----
void example_exceptions() {
    std::cout << "=== 例外ハンドリング ===\n";

    // 壊れたファイル → ParseError or FormatError
    try {
        std::vector<uint8_t> garbage(100, 0xFF);
        WorkbookFactory::open(garbage);
    } catch (const FormatError& e) {
        std::cout << "FormatError: " << e.what() << "\n";
    }

    // 存在しないシート → RangeError
    try {
        auto wb = WorkbookFactory::create();
        wb->sheet(99);
    } catch (const RangeError& e) {
        std::cout << "RangeError: " << e.what() << "\n";
    }

    // 重複シート名 → WriteError
    try {
        auto wb = WorkbookFactory::create();
        wb->add_sheet("S");
        wb->add_sheet("S");
    } catch (const WriteError& e) {
        std::cout << "WriteError: " << e.what() << "\n";
    }

    // 型ミスマッチ → std::bad_variant_access
    try {
        CellValue v = std::string{"hello"};
        get_int(v);
    } catch (const std::bad_variant_access&) {
        std::cout << "bad_variant_access: 型が違う\n";
    }

    // 全例外を一括で受け取る場合
    try {
        auto wb = WorkbookFactory::open("存在しないファイル.xlsx");
    } catch (const ExcelError& e) {
        std::cout << "ExcelError: " << e.what() << "\n";
    }
}

// ---- 例 2: 安全なセルアクセスパターン ----
void example_safe_access() {
    std::cout << "\n=== 安全なセルアクセス ===\n";

    auto wb = WorkbookFactory::create();
    auto& sh = wb->add_sheet("S");
    sh.set_cell("A1", std::string{"hello"});
    sh.set_cell("B1", int64_t{42});
    sh.set_cell("C1", 3.14);

    // ---- パターン 1: is_* で事前チェック（最も安全）----
    Cell c = sh.cell("A1");
    if (is_string(c.value))      std::cout << "string: " << get_string(c.value) << "\n";
    else if (is_int(c.value))    std::cout << "int: "    << get_int(c.value) << "\n";
    else if (is_numeric(c.value))std::cout << "number: " << as_double(c.value) << "\n";

    // ---- パターン 2: try_cell で空セルを安全に処理 ----
    if (auto opt = sh.try_cell(0, 0)) {  // A1
        std::cout << "A1 exists: " << to_string(opt->value) << "\n";
    }
    if (!sh.try_cell(99, 99)) {          // 空のセル
        std::cout << "Z100 is empty\n";
    }

    // ---- パターン 3: std::visit で全型を網羅 ----
    sh.for_each_cell([](const Cell& c) {
        std::visit([&](auto&& v) {
            using T = std::decay_t<decltype(v)>;
            std::cout << c.address.to_a1() << ": ";
            if      constexpr (std::is_same_v<T, BlankValue>)   std::cout << "(blank)";
            else if constexpr (std::is_same_v<T, bool>)         std::cout << (v ? "TRUE" : "FALSE");
            else if constexpr (std::is_same_v<T, int64_t>)      std::cout << v;
            else if constexpr (std::is_same_v<T, double>)       std::cout << v;
            else if constexpr (std::is_same_v<T, std::string>)  std::cout << '"' << v << '"';
            else if constexpr (std::is_same_v<T, ErrorValue>)   std::cout << v.code;
            std::cout << "\n";
        }, c.value);
    });
}

// ---- 例 3: 数値の集計 ----
void example_aggregate() {
    std::cout << "\n=== 数値集計 ===\n";

    auto wb = WorkbookFactory::create();
    auto& sh = wb->add_sheet("Data");
    sh.set_table("A1", {
        {std::string{"Name"}, std::string{"Score"}},
        {std::string{"Alice"}, int64_t{95}},
        {std::string{"Bob"},   int64_t{87}},
        {std::string{"Carol"}, int64_t{92}},
        {std::string{"Dave"},  int64_t{78}},
    });

    // 数値列（B列、1行目はヘッダーなのでスキップ）の合計と平均
    double sum = 0; int count = 0;
    for (uint32_t r = 1; r < sh.row_count(); ++r) {
        if (auto c = sh.try_cell(r, 1)) {   // B列
            if (is_numeric(c->value)) {
                sum += as_double(c->value);
                ++count;
            }
        }
    }
    if (count > 0)
        std::cout << "合計: " << sum << " / 平均: " << sum/count << "\n";
}

int main() {
    example_exceptions();
    example_safe_access();
    example_aggregate();
    return 0;
}
