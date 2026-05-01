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

// excellib の全機能をまとめて使えるようにするメインのヘッダーをインクルード
#include "excellib/excellib.hpp"
// std::cout など標準入出力を使うためのヘッダー
#include <iostream>
// std::setw (表示幅の指定) を使うためのヘッダー
#include <iomanip>

// excellib:: というプレフィックスを省略できるようにする
// (例: excellib::WorkbookFactory → WorkbookFactory と書けるようになる)
using namespace excellib;

// プログラムの開始地点。argc=引数の個数、argv=引数の文字列配列
int main(int argc, char* argv[]) {
    // 引数が1個以下 (プログラム名だけ) の場合は使い方を表示して終了
    if (argc < 2) {
        // argv[0] はこのプログラム自身の名前
        std::cerr << "Usage: " << argv[0] << " <file.xlsx|file.xls>\n";
        // 0以外を返すと「エラーで終了」を意味する
        return 1;
    }

    // try ブロック: 例外(エラー)が起きても安全に処理するための構文
    try {
        // ---- ファイルを開く ----
        // フォーマットは拡張子ではなくマジックバイトで自動判定
        // argv[1] はコマンドライン引数で渡されたファイルパス
        // open() はファイルを読み込んで Workbook オブジェクトを返す
        // auto は型を自動推論 (実際は std::unique_ptr<Workbook> 型)
        auto wb = WorkbookFactory::open(argv[1]);

        // ファイル名・フォーマット・シート数を表示
        // wb-> でポインタ経由のメンバ関数呼び出し
        // format_name() は "XLSX" や "XLS" などの文字列を返す
        // sheet_count() はシートの総数 (整数) を返す
        std::cout << "File   : " << argv[1] << "\n"
                  << "Format : " << wb->format_name() << "\n"
                  << "Sheets : " << wb->sheet_count() << "\n\n";

        // ---- シート一覧 ----
        // 0 から sheet_count()-1 まで繰り返す (シートは 0-indexed)
        for (size_t si = 0; si < wb->sheet_count(); ++si) {
            // si 番目のシートへの参照を取得 (参照なのでコピーは発生しない)
            auto& sh = wb->sheet(si);
            // シートのインデックス・名前・行数・列数を表示
            std::cout << "=== [" << si << "] " << sh.name()
                      << " (" << sh.row_count() << " rows x "
                      << sh.col_count() << " cols) ===\n";

            // ---- セルアクセス ----

            // 方法 1: A1 記法 (Excelと同じ "A1" という文字列でセルを指定)
            Cell a1 = sh.cell("A1");
            // to_string() で CellValue をどんな型でも文字列に変換して表示
            std::cout << "A1 = " << to_string(a1.value);
            // has_formula() が true なら数式文字列も表示する
            // *a1.formula で optional<string> の中身を取り出す
            if (a1.has_formula()) std::cout << "  [=" << *a1.formula << "]";
            std::cout << "\n";

            // 方法 2: 行・列インデックス（0-based）
            // (1, 1) → 2行目・2列目 = Excelのセル "B2" に相当
            Cell b2 = sh.cell(1, 1);   // B2
            std::cout << "B2 = " << to_string(b2.value) << "\n";

            // 方法 3: 空セルで nullopt を返す安全版
            // try_cell() は空やアクセス範囲外のセルを例外ではなく nullopt で返す
            // if (auto c = ...) という書き方は C++17 の「if初期化文」
            if (auto c = sh.try_cell(0, 0)) {
                // c には std::optional<Cell> が入っており、値がある場合だけここを実行
                // c-> で optional の中身の Cell にアクセス
                std::cout << "A1 (try) = " << to_string(c->value) << "\n";
            }

            // ---- 型チェックしてから取得 ----
            // is_string() / is_int() などで型を確認してから get_*() で値を取り出す
            // これにより型ミスマッチの例外 (bad_variant_access) を防げる
            if (is_string(a1.value))  std::cout << "  → 文字列: " << get_string(a1.value) << "\n";
            if (is_int(a1.value))     std::cout << "  → 整数: "   << get_int(a1.value)    << "\n";
            if (is_double(a1.value))  std::cout << "  → 実数: "   << get_double(a1.value) << "\n";
            if (is_bool(a1.value))    std::cout << "  → 真偽: "   << get_bool(a1.value)   << "\n";
            // is_numeric() は int64_t または double のどちらでも true になる
            // as_double() は int64_t も double も統一して double に変換して返す
            if (is_numeric(a1.value)) std::cout << "  → 数値: "   << as_double(a1.value)  << "\n";

            // ---- 全非空セルを走査 ----
            std::cout << "\n非空セル:\n";
            int n = 0;    // 表示件数のカウンター
            // for_each_cell() はシート上のすべての非空セルに対してラムダ関数を呼ぶ
            // [&] はラムダの外側にある変数 (n など) をそのまま参照できるようにする
            sh.for_each_cell([&](const Cell& c) {
                // 最初の10件だけ表示する
                if (n++ < 10) {
                    // setw(5) で5文字幅に揃えて表示
                    // address.to_a1() でセルのアドレスを "A1" 形式の文字列に変換
                    std::cout << std::setw(5) << c.address.to_a1() << " : "
                              << to_string(c.value) << "\n";
                }
            });
            // 10件を超えた分は「他 ○ セル」とまとめて表示
            if (n > 10) std::cout << "  ... 他 " << n - 10 << " セル\n";
            std::cout << "\n";
        }

    // catch: try ブロック内で ExcelError (またはその派生例外) が投げられたときに実行
    // e.what() でエラーメッセージの文字列を取得できる
    } catch (const ExcelError& e) {
        std::cerr << "Error: " << e.what() << "\n";
        // エラーが起きたので 1 (異常終了) を返す
        return 1;
    }
    // 正常終了は 0 を返す (慣例)
    return 0;
}
