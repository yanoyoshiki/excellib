/**
 * examples/02_write.cpp
 *
 * XLSX ファイルの新規作成サンプル。
 *
 * 実行:
 *   ./02_write
 *   → output.xlsx が生成される
 */
#include "excellib/excellib.hpp"
#include <iostream>

using namespace excellib;

int main() {
    try {
        auto wb = WorkbookFactory::create(FileFormat::XLSX);

        // ---- シート 1: セルを個別に設定 ----
        auto& sales = wb->add_sheet("売上");

        // ヘッダー
        sales.set_cell("A1", std::string{"製品名"});
        sales.set_cell("B1", std::string{"単価"});
        sales.set_cell("C1", std::string{"数量"});
        sales.set_cell("D1", std::string{"合計"});

        // データ
        const struct { const char* name; double price; int qty; } data[] = {
            {"WidgetA",  9.99, 100},
            {"WidgetB",  4.49, 250},
            {"WidgetC", 24.99,  75},
        };
        for (int i = 0; i < 3; ++i) {
            int row = i + 1;  // 2行目から
            sales.set_cell(uint32_t(row), 0, std::string{data[i].name});
            sales.set_cell(uint32_t(row), 1, data[i].price);
            sales.set_cell(uint32_t(row), 2, int64_t{data[i].qty});
            // 数式
            sales.set_formula("D" + std::to_string(row + 1),
                              "B" + std::to_string(row + 1) +
                              "*C" + std::to_string(row + 1));
        }
        sales.set_cell("A5", std::string{"合計"});
        sales.set_formula("D5", "SUM(D2:D4)");

        // ---- シート 2: set_table で 2D データを一括書き込み ----
        auto& report = wb->add_sheet("月次");
        report.set_table("A1", {
            {std::string{"月"},   std::string{"売上"},  std::string{"費用"},  std::string{"利益"}},
            {std::string{"1月"},  int64_t{1200000},  int64_t{800000},  int64_t{400000}},
            {std::string{"2月"},  int64_t{1350000},  int64_t{870000},  int64_t{480000}},
            {std::string{"3月"},  int64_t{1580000},  int64_t{920000},  int64_t{660000}},
            {std::string{"Q1計"}, int64_t{4130000},  int64_t{2590000}, int64_t{1540000}},
        });

        // ---- 保存 ----
        wb->save("output.xlsx");
        std::cout << "output.xlsx を生成しました\n";

        // ---- バイト列として取得（HTTP レスポンスなどに使う場合） ----
        auto bytes = wb->to_bytes(FileFormat::XLSX, {});
        std::cout << "バイト列サイズ: " << bytes.size() << " bytes\n";

        // ---- 再読み込みして確認 ----
        auto wb2 = WorkbookFactory::open(bytes);
        std::cout << "再読み込み: " << wb2->sheet(0).name()
                  << " / A1 = " << get_string(wb2->sheet(0).cell("A1").value) << "\n";

    } catch (const ExcelError& e) {
        std::cerr << "Error: " << e.what() << "\n";
        return 1;
    }
    return 0;
}
