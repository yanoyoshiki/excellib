/**
 * examples/02_write.cpp
 *
 * XLSX ファイルの新規作成サンプル。
 * セルへの個別書き込み・数式設定・テーブル一括書き込み・
 * ファイル保存・バイト列取得・再読み込みの方法を示す。
 *
 * 実行:
 *   ./02_write
 *   → output.xlsx が生成される
 */
// excellib.hpp: WorkbookFactory, Sheet, Cell など全クラスを一括インクルード
#include "excellib/excellib.hpp"
// std::cout / std::cerr: 標準出力 / 標準エラー出力
#include <iostream>

// excellib 名前空間を使用する宣言
using namespace excellib;

int main() {
    try {
        // ============================================================
        //  新規ワークブックの作成
        // ============================================================
        // WorkbookFactory::create(): 新しい空のワークブックを作成する
        // FileFormat::XLSX: XLSX 形式で作成 (デフォルト)
        // 戻り値は std::unique_ptr<Workbook>
        auto wb = WorkbookFactory::create(FileFormat::XLSX);

        // ============================================================
        //  シート 1: セルを個別に設定する
        // ============================================================
        // add_sheet(): シートを追加して参照 (Sheet&) を返す
        auto& sales = wb->add_sheet("売上");

        // ---- ヘッダー行の書き込み ----
        // set_cell(A1記法, 値): 指定セルに値を書き込む
        // std::string{...}: 文字列リテラルを明示的に std::string 型にキャストする
        //   (CellValue は variant なので型を明示する必要がある)
        sales.set_cell("A1", std::string{"製品名"});
        sales.set_cell("B1", std::string{"単価"});
        sales.set_cell("C1", std::string{"数量"});
        sales.set_cell("D1", std::string{"合計"});

        // ---- データ行の書き込み ----
        // 無名構造体の配列でデータを定義する
        const struct { const char* name; double price; int qty; } data[] = {
            {"WidgetA",  9.99, 100},
            {"WidgetB",  4.49, 250},
            {"WidgetC", 24.99,  75},
        };
        for (int i = 0; i < 3; ++i) {
            int row = i + 1;  // ヘッダーが 0 行目なので、データは 1 行目から (0-indexed)
            // set_cell(行インデックス, 列インデックス, 値): 0-indexed でセルを指定
            sales.set_cell(uint32_t(row), 0, std::string{data[i].name}); // A列: 製品名
            sales.set_cell(uint32_t(row), 1, data[i].price);             // B列: 単価 (double)
            sales.set_cell(uint32_t(row), 2, int64_t{data[i].qty});      // C列: 数量 (整数)
            // set_formula(): セルに数式を設定する
            // "D2:D4" はそれぞれ "=B2*C2", "=B3*C3", "=B4*C4" になる
            // row + 1 は 0-indexed → 1-indexed の変換 (A1 記法は 1 始まり)
            sales.set_formula("D" + std::to_string(row + 1),
                              "B" + std::to_string(row + 1) +
                              "*C" + std::to_string(row + 1));
        }
        // 合計行
        sales.set_cell("A5", std::string{"合計"});
        // SUM(D2:D4): D2 から D4 の合計を D5 に設定する
        sales.set_formula("D5", "SUM(D2:D4)");

        // ============================================================
        //  シート 2: set_table() で 2D データを一括書き込みする
        // ============================================================
        // add_sheet(): 2 枚目のシートを追加する
        auto& report = wb->add_sheet("月次");
        // set_table(左上セルのA1記法, 2次元データ): 指定位置を起点に表を書き込む
        // 外側の vector が行、内側の vector が列に対応する
        // 各セルの型は CellValue (variant) に合わせて明示的に指定する
        report.set_table("A1", {
            // ヘッダー行: 全て std::string
            {std::string{"月"},   std::string{"売上"},  std::string{"費用"},  std::string{"利益"}},
            // データ行: 月名は std::string、数値は int64_t
            {std::string{"1月"},  int64_t{1200000},  int64_t{800000},  int64_t{400000}},
            {std::string{"2月"},  int64_t{1350000},  int64_t{870000},  int64_t{480000}},
            {std::string{"3月"},  int64_t{1580000},  int64_t{920000},  int64_t{660000}},
            {std::string{"Q1計"}, int64_t{4130000},  int64_t{2590000}, int64_t{1540000}},
        });

        // ============================================================
        //  ファイルに保存する
        // ============================================================
        // save(): ファイルパスに XLSX として保存する
        // 保存先ディレクトリへの書き込み権限が必要
        wb->save("output.xlsx");
        std::cout << "output.xlsx を生成しました\n";

        // ============================================================
        //  バイト列として取得する (HTTP レスポンスやメモリ転送に使う場合)
        // ============================================================
        // to_bytes(): XLSX のバイト列 (= ZIP アーカイブのバイナリ) を返す
        // ファイルに保存せずにメモリ上で扱いたい場合に便利
        // 第1引数: 出力フォーマット (FileFormat::XLSX のみ対応)
        // 第2引数: SaveOptions (現バージョンでは未使用)
        auto bytes = wb->to_bytes(FileFormat::XLSX, {});
        std::cout << "バイト列サイズ: " << bytes.size() << " bytes\n";

        // ============================================================
        //  バイト列から再読み込みする
        // ============================================================
        // open(バイト列): ファイルに保存せずにメモリ上のバイト列から開く
        // フォーマットはマジックバイトで自動判定される
        auto wb2 = WorkbookFactory::open(bytes);
        // sheet(0): 最初のシートを取得
        // get_string(): CellValue が std::string 型であることが確認済みの場合に使う
        std::cout << "再読み込み: " << wb2->sheet(0).name()
                  << " / A1 = " << get_string(wb2->sheet(0).cell("A1").value) << "\n";

    } catch (const ExcelError& e) {
        // ExcelError: ParseError / IOError / WriteError / RangeError など全例外の基底クラス
        std::cerr << "Error: " << e.what() << "\n";
        return 1;
    }
    return 0;
}
