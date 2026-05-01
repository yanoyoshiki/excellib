/**
 * examples/01_read.cpp
 *
 * XLS / XLSX ファイルの読み取りサンプル。
 * excellib を使って Excel ファイルを開き、シートとセルのデータを取り出す方法を示す。
 *
 * ビルド:
 *   g++ -std=c++17 -Iinclude 01_read.cpp [ライブラリ .o 群] -o 01_read
 *
 * 実行:
 *   ./01_read data.xlsx
 *   ./01_read legacy.xls     # 旧形式も同じ API で読める
 */
// excellib.hpp: WorkbookFactory, Sheet, Cell など全クラスを一括インクルード
#include "excellib/excellib.hpp"
// std::cout / std::cerr: 標準出力 / 標準エラー出力
#include <iostream>
// std::setw: 出力幅の指定 (セル一覧の整形に使う)
#include <iomanip>

// excellib 名前空間を使用する宣言 (excellib:: を毎回書かなくてよくなる)
using namespace excellib;

int main(int argc, char* argv[]) {
    // コマンドライン引数のチェック: ファイルパスが指定されていなければ使い方を表示
    if (argc < 2) {
        std::cerr << "Usage: " << argv[0] << " <file.xlsx|file.xls>\n";
        return 1;   // 終了コード 1 = エラー
    }

    try {
        // ============================================================
        //  ファイルを開く
        // ============================================================
        // WorkbookFactory::open(): ファイルパスからワークブックを開く
        // フォーマットは拡張子ではなく「マジックバイト」(先頭数バイト) で自動判定する
        //   OLE2 シグネチャ → XLS (BIFF8 形式)
        //   ZIP  シグネチャ → XLSX (OOXML 形式)
        // 戻り値は std::unique_ptr<Workbook> (自動で解放される)
        auto wb = WorkbookFactory::open(argv[1]);

        // フォーマット名・シート数を出力する
        // wb->format_name() は "XLS" または "XLSX" を返す
        std::cout << "File   : " << argv[1] << "\n"
                  << "Format : " << wb->format_name() << "\n"
                  << "Sheets : " << wb->sheet_count() << "\n\n";

        // ============================================================
        //  シート一覧を走査する
        // ============================================================
        // sheet_count(): ワークブック内のシート数を返す
        for (size_t si = 0; si < wb->sheet_count(); ++si) {
            // sheet(インデックス): シートへの参照を返す (0-indexed)
            auto& sh = wb->sheet(si);
            // name(): シート名、row_count()/col_count(): データのある行数・列数
            std::cout << "=== [" << si << "] " << sh.name()
                      << " (" << sh.row_count() << " rows x "
                      << sh.col_count() << " cols) ===\n";

            // ============================================================
            //  セルアクセスの 3 つの方法
            // ============================================================

            // 方法 1: A1 記法 (文字列) でアクセス
            // "A1" → row=0, col=0 (内部は 0-indexed に変換される)
            Cell a1 = sh.cell("A1");
            std::cout << "A1 = " << to_string(a1.value);  // to_string: CellValue を文字列に変換
            // has_formula(): 数式セルなら true を返す
            // formula: std::optional<std::string> → *a1.formula でアクセス
            if (a1.has_formula()) std::cout << "  [=" << *a1.formula << "]";
            std::cout << "\n";

            // 方法 2: 行・列インデックス (0-based) でアクセス
            // cell(row, col): row=1 (2行目), col=1 (B列) → B2 セル
            Cell b2 = sh.cell(1, 1);
            std::cout << "B2 = " << to_string(b2.value) << "\n";

            // 方法 3: try_cell() — 空セルのときに nullopt を返す安全版
            // std::optional<Cell> を返すため、空セルにアクセスしても例外が発生しない
            if (auto c = sh.try_cell(0, 0)) {
                // c が nullopt でなければ (= セルにデータがある) 値を出力
                std::cout << "A1 (try) = " << to_string(c->value) << "\n";
            }

            // ============================================================
            //  型チェックしてから値を取得する
            // ============================================================
            // CellValue は std::variant<BlankValue, bool, int64_t, double, string, ErrorValue>
            // is_*() で型を確認してから get_*() で値を取得するのが安全
            if (is_string(a1.value))  std::cout << "  → 文字列: " << get_string(a1.value) << "\n";
            if (is_int(a1.value))     std::cout << "  → 整数: "   << get_int(a1.value)    << "\n";
            if (is_double(a1.value))  std::cout << "  → 実数: "   << get_double(a1.value) << "\n";
            if (is_bool(a1.value))    std::cout << "  → 真偽: "   << get_bool(a1.value)   << "\n";
            // is_numeric(): is_int() または is_double() が true の場合に true を返す
            // as_double(): int64_t または double を double に変換 (文字列なら FormatError)
            if (is_numeric(a1.value)) std::cout << "  → 数値: "   << as_double(a1.value)  << "\n";

            // ============================================================
            //  全非空セルを走査する
            // ============================================================
            std::cout << "\n非空セル:\n";
            int n = 0;
            // for_each_cell(): 空でないセルを全て順番にラムダに渡す
            // ラムダのキャプチャ: n をカウンターとして使う
            sh.for_each_cell([&](const Cell& c) {
                if (n++ < 10) {
                    // to_a1(): CellAddress (0-indexed) を "B3" のような A1 記法に変換
                    std::cout << std::setw(5) << c.address.to_a1() << " : "
                              << to_string(c.value) << "\n";
                }
            });
            // 10 件以上あれば残り件数を表示する
            if (n > 10) std::cout << "  ... 他 " << n - 10 << " セル\n";
            std::cout << "\n";
        }

    } catch (const ExcelError& e) {
        // ExcelError は全 excellib 例外の基底クラス
        // ParseError / FormatError / IOError / RangeError など全てをここで受け取れる
        std::cerr << "Error: " << e.what() << "\n";
        return 1;
    }
    return 0;   // 終了コード 0 = 正常終了
}
