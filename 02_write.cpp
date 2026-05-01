/**
 * examples/02_write.cpp
 *
 * XLSX ファイルの新規作成サンプル。
 *
 * 実行:
 *   ./02_write
 *   → output.xlsx が生成される
 */

// excellib の全機能をまとめて使えるようにするメインヘッダー
#include "excellib/excellib.hpp"
// std::cout など標準入出力を使うためのヘッダー
#include <iostream>

// excellib:: プレフィックスを省略できるようにする
using namespace excellib;

// プログラムの開始地点 (引数なし版)
int main() {
    // try ブロック: 例外が起きても安全に処理するための構文
    try {
        // 新規の XLSX ワークブックを作成する
        // FileFormat::XLSX を指定することで .xlsx 形式を作成
        // auto は型を自動推論 (実際は std::unique_ptr<Workbook>)
        auto wb = WorkbookFactory::create(FileFormat::XLSX);

        // ---- シート 1: セルを個別に設定 ----
        // "売上" という名前のシートを追加し、その参照を取得
        // & をつけることでコピーではなく参照として受け取る (効率的)
        auto& sales = wb->add_sheet("売上");

        // ヘッダー行を設定 (A1〜D1)
        // set_cell("A1", 値) で A1 セルに値を書き込む
        // std::string{...} で文字列リテラルを std::string 型として渡す (型推論のため必要)
        sales.set_cell("A1", std::string{"製品名"});
        sales.set_cell("B1", std::string{"単価"});
        sales.set_cell("C1", std::string{"数量"});
        sales.set_cell("D1", std::string{"合計"});

        // データ行の定義
        // const struct { ... } data[] は無名構造体の配列を定義するC++の書き方
        // 各要素に name (文字列)・price (小数)・qty (整数) を持たせている
        const struct { const char* name; double price; int qty; } data[] = {
            {"WidgetA",  9.99, 100},
            {"WidgetB",  4.49, 250},
            {"WidgetC", 24.99,  75},
        };
        // 3件のデータを 2〜4 行目に書き込む
        for (int i = 0; i < 3; ++i) {
            // 0-indexed なので行インデックスは i+1 (ヘッダーの次の行から)
            int row = i + 1;  // 2行目から
            // 行・列インデックスで書き込む方法 (0-indexed)
            // uint32_t にキャストするのはオーバーロードの型解決のため
            sales.set_cell(uint32_t(row), 0, std::string{data[i].name});
            // price は double 型
            sales.set_cell(uint32_t(row), 1, data[i].price);
            // qty は int → int64_t に変換して渡す (ライブラリの型要件)
            sales.set_cell(uint32_t(row), 2, int64_t{data[i].qty});
            // 数式
            // set_formula() でセルに Excel の数式文字列を書き込む
            // "D" + to_string(row + 1) で "D2", "D3", "D4" のような文字列を組み立てる
            sales.set_formula("D" + std::to_string(row + 1),
                              "B" + std::to_string(row + 1) +
                              "*C" + std::to_string(row + 1));
        }
        // 5行目に「合計」ラベルを書く
        sales.set_cell("A5", std::string{"合計"});
        // D2〜D4 の合計を SUM 関数で計算する数式を D5 に設定
        sales.set_formula("D5", "SUM(D2:D4)");

        // ---- シート 2: set_table で 2D データを一括書き込み ----
        // "月次" という名前のシートを追加
        auto& report = wb->add_sheet("月次");
        // set_table() は A1 を起点に 2次元ベクタのデータを一括で書き込む
        // 外側のベクタが行、内側のベクタが列に対応する
        // CellValue 型は string / int64_t / double / bool などの統一型
        report.set_table("A1", {
            // 1行目: ヘッダー
            {std::string{"月"},   std::string{"売上"},  std::string{"費用"},  std::string{"利益"}},
            // 2行目以降: データ (int64_t で整数を渡す)
            {std::string{"1月"},  int64_t{1200000},  int64_t{800000},  int64_t{400000}},
            {std::string{"2月"},  int64_t{1350000},  int64_t{870000},  int64_t{480000}},
            {std::string{"3月"},  int64_t{1580000},  int64_t{920000},  int64_t{660000}},
            {std::string{"Q1計"}, int64_t{4130000},  int64_t{2590000}, int64_t{1540000}},
        });

        // ---- 保存 ----
        // save() でファイルに書き出す。ファイルが存在しない場合は新規作成される
        wb->save("output.xlsx");
        std::cout << "output.xlsx を生成しました\n";

        // ---- バイト列として取得（HTTP レスポンスなどに使う場合） ----
        // to_bytes() はファイルに保存せず、メモリ上のバイト列として返す
        // std::vector<uint8_t> は符号なし8ビット整数 (=バイト) の動的配列
        // {} は SaveOptions のデフォルト値 (省略しても同じ)
        auto bytes = wb->to_bytes(FileFormat::XLSX, {});
        // bytes.size() でバイト数を確認
        std::cout << "バイト列サイズ: " << bytes.size() << " bytes\n";

        // ---- 再読み込みして確認 ----
        // to_bytes() で得たバイト列を直接 open() に渡せる (ファイルを経由しない)
        auto wb2 = WorkbookFactory::open(bytes);
        // 再読み込んだワークブックの 0番目のシートの A1 セルの文字列値を表示
        std::cout << "再読み込み: " << wb2->sheet(0).name()
                  << " / A1 = " << get_string(wb2->sheet(0).cell("A1").value) << "\n";

    // ExcelError (またはその派生) が発生した場合にここを実行
    } catch (const ExcelError& e) {
        // cerr は標準エラー出力 (通常はターミナルに赤字等で表示)
        std::cerr << "Error: " << e.what() << "\n";
        // 1 を返すことで異常終了をシェルに通知
        return 1;
    }
    // 正常終了
    return 0;
}
