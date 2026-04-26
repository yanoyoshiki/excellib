/**
 * examples/01_read.cpp
 *
 * XLS / XLSX ファイルの読み取りサンプル。
 *
 * ビルド:
 *   g++ -std=c++17 -Iinclude 01_read.cpp [ライブラリ .o 群] -lz -o 01_read
 *
 * 実行:
 *   ./01_read data.xlsx
 *   ./01_read legacy.xls     # 旧形式も同じ API で読める
 */
#include "excellib/excellib.hpp"
#include <iostream>
#include <iomanip>

using namespace excellib;

int main(int argc, char* argv[]) {
    if (argc < 2) {
        std::cerr << "Usage: " << argv[0] << " <file.xlsx|file.xls>\n";
        return 1;
    }

    try {
        // ---- ファイルを開く ----
        // フォーマットは拡張子ではなくマジックバイトで自動判定
        auto wb = WorkbookFactory::open(argv[1]);

        std::cout << "File   : " << argv[1] << "\n"
                  << "Format : " << wb->format_name() << "\n"
                  << "Sheets : " << wb->sheet_count() << "\n\n";

        // ---- シート一覧 ----
        for (size_t si = 0; si < wb->sheet_count(); ++si) {
            auto& sh = wb->sheet(si);
            std::cout << "=== [" << si << "] " << sh.name()
                      << " (" << sh.row_count() << " rows x "
                      << sh.col_count() << " cols) ===\n";

            // ---- セルアクセス ----

            // 方法 1: A1 記法
            Cell a1 = sh.cell("A1");
            std::cout << "A1 = " << to_string(a1.value);
            if (a1.has_formula()) std::cout << "  [=" << *a1.formula << "]";
            std::cout << "\n";

            // 方法 2: 行・列インデックス（0-based）
            Cell b2 = sh.cell(1, 1);   // B2
            std::cout << "B2 = " << to_string(b2.value) << "\n";

            // 方法 3: 空セルで nullopt を返す安全版
            if (auto c = sh.try_cell(0, 0)) {
                std::cout << "A1 (try) = " << to_string(c->value) << "\n";
            }

            // ---- 型チェックしてから取得 ----
            if (is_string(a1.value))  std::cout << "  → 文字列: " << get_string(a1.value) << "\n";
            if (is_int(a1.value))     std::cout << "  → 整数: "   << get_int(a1.value)    << "\n";
            if (is_double(a1.value))  std::cout << "  → 実数: "   << get_double(a1.value) << "\n";
            if (is_bool(a1.value))    std::cout << "  → 真偽: "   << get_bool(a1.value)   << "\n";
            if (is_numeric(a1.value)) std::cout << "  → 数値: "   << as_double(a1.value)  << "\n";

            // ---- 全非空セルを走査 ----
            std::cout << "\n非空セル:\n";
            int n = 0;
            sh.for_each_cell([&](const Cell& c) {
                if (n++ < 10) {
                    std::cout << std::setw(5) << c.address.to_a1() << " : "
                              << to_string(c.value) << "\n";
                }
            });
            if (n > 10) std::cout << "  ... 他 " << n - 10 << " セル\n";
            std::cout << "\n";
        }

    } catch (const ExcelError& e) {
        std::cerr << "Error: " << e.what() << "\n";
        return 1;
    }
    return 0;
}
