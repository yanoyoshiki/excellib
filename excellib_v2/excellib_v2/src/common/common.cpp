// workbook.hpp と print_settings.hpp の実装ファイル
// CellAddress の変換・CellValue の文字列化・PrintArea のパースなどを実装する
#include "excellib/workbook.hpp"
#include "excellib/print_settings.hpp"
// isalpha / isdigit / toupper などの文字分類関数
#include <cctype>
// std::reverse: 文字列の逆順化に使う
#include <algorithm>
// std::ostringstream: 文字列を組み立てるためのストリーム
#include <sstream>

namespace excellib {

// ============================================================
//  CellAddress
// ============================================================
// from_a1(): "B3" のような A1 記法の文字列をパースして {row=2, col=1} を返す
// Excelの A1 記法: アルファベット部分が列 (A=1, B=2, ..., Z=26, AA=27, ...)
//                  数字部分が行 (1始まり)
// このライブラリは 0-indexed なので返り値は row-1, col-1 になる
CellAddress CellAddress::from_a1(const std::string& a1) {
    // 空文字列は不正
    if (a1.empty())
        throw FormatError("Empty A1 reference");

    uint32_t col = 0;    // 列番号 (計算中)
    size_t   i   = 0;    // 現在の解析位置

    // アルファベット部分を読み込んで列番号を計算
    // (例) "AB" → 1*26 + 2 = 28
    while (i < a1.size() && std::isalpha(static_cast<unsigned char>(a1[i]))) {
        // 小文字を大文字に変換 (A=65, B=66, ...)
        char c = static_cast<char>(std::toupper(static_cast<unsigned char>(a1[i])));
        // 26進数として計算 (A=1, B=2, ... Z=26, AA=27, ...)
        col = col * 26 + static_cast<uint32_t>(c - 'A' + 1);
        ++i;
    }
    // アルファベットが1文字もなかった場合は不正
    if (i == 0)
        throw FormatError("No column letters in A1 reference: " + a1);
    // Excel の最大列数は 16384 (XFD)
    if (col == 0 || col > 16384)
        throw FormatError("Column out of Excel range in: " + a1);

    // アルファベットの後に数字がなかった場合は不正
    if (i == a1.size())
        throw FormatError("No row digits in A1 reference: " + a1);

    // 数字部分を読み込んで行番号を計算
    uint32_t row = 0;
    while (i < a1.size()) {
        unsigned char c = static_cast<unsigned char>(a1[i]);
        // 数字以外の文字があれば不正
        if (!std::isdigit(c))
            throw FormatError("Non-digit in row part of A1 reference: " + a1);
        // 10進数として計算
        row = row * 10 + static_cast<uint32_t>(c - '0');
        ++i;
    }
    // Excel の最大行数は 1048576
    if (row == 0 || row > 1048576)
        throw FormatError("Row out of Excel range in: " + a1);

    // 1-indexed → 0-indexed に変換して返す
    return CellAddress{row - 1, col - 1};
}

// to_a1(): 0-indexed の {row, col} を "B3" のような A1 記法の文字列に変換する
std::string CellAddress::to_a1() const {
    // 列番号を列文字列に変換
    std::string col_str;
    // col は 0-indexed なので +1 して 1-indexed に戻す
    uint32_t c = col + 1;
    while (c > 0) {
        --c;    // 0-indexed に変換 (A=0, B=1, ...)
        // 'A' + (c % 26) で余り分の列文字を取得
        col_str += static_cast<char>('A' + (c % 26));
        c /= 26;    // 次の桁へ
    }
    // 逆順に作成したので reverse で正順に直す (例: "BA" → "AB")
    std::reverse(col_str.begin(), col_str.end());
    // 列文字列 + (行番号+1) を連結して返す (0-indexed → 1-indexed)
    return col_str + std::to_string(row + 1);
}

// ============================================================
//  CellValue to_string
// ============================================================
// to_string(): CellValue (variant 型) をどんな型でも文字列に変換して返す
std::string to_string(const CellValue& v) {
    // std::visit はラムダを variant の実際の型に合わせて呼び出す
    return std::visit([](auto&& val) -> std::string {
        // std::decay_t<decltype(val)> で val の純粋な型名 T を取得
        using T = std::decay_t<decltype(val)>;
        // コンパイル時型分岐 (if constexpr) で各型に対応する変換を行う
        if constexpr (std::is_same_v<T, BlankValue>)
            return "";    // 空セルは空文字列
        else if constexpr (std::is_same_v<T, bool>)
            return val ? "TRUE" : "FALSE";    // bool は "TRUE"/"FALSE"
        else if constexpr (std::is_same_v<T, int64_t>)
            return std::to_string(val);    // 整数は to_string で変換
        else if constexpr (std::is_same_v<T, double>) {
            // double は ostringstream を使う (精度の問題があるため to_string は使わない)
            std::ostringstream os;
            os << val;
            return os.str();
        }
        else if constexpr (std::is_same_v<T, std::string>)
            return val;    // 文字列はそのまま返す
        else if constexpr (std::is_same_v<T, ErrorValue>)
            return val.code;    // エラー値はコード文字列 ("#REF!" など) を返す
    }, v);
}

// ============================================================
//  Sheet::set_table helper
// ============================================================
// set_table(): 2次元データをシートの指定位置に一括書き込みする
// top_left: 左上セルの A1 記法 (例: "B2")
// data: 2次元のセル値リスト (外側が行、内側が列)
void Sheet::set_table(const std::string& top_left,
                      const std::vector<std::vector<CellValue>>& data) {
    // top_left の A1 記法を CellAddress (0-indexed) に変換
    CellAddress origin = CellAddress::from_a1(top_left);
    // 行ループ: r = 0, 1, 2, ...
    for (size_t r = 0; r < data.size(); ++r)
        // 列ループ: c = 0, 1, 2, ...
        for (size_t c = 0; c < data[r].size(); ++c)
            // origin を基点に r行c列オフセットした位置に値を書き込む
            set_cell(static_cast<uint32_t>(origin.row + r),
                     static_cast<uint32_t>(origin.col + c),
                     data[r][c]);
}

// ============================================================
//  PrintArea
// ============================================================
// from_range(): "A1:Z50" のような範囲文字列を PrintArea に変換する
PrintArea PrintArea::from_range(const std::string& range) {
    // ':' を探して範囲の区切りを見つける
    auto colon = range.find(':');
    PrintArea pa;
    if (colon == std::string::npos) {
        // ':' がなければ単一セルの指定
        auto addr = CellAddress::from_a1(range);
        // 開始と終了を同じ値にする
        pa.first_row = pa.last_row = addr.row;
        pa.first_col = pa.last_col = addr.col;
    } else {
        // ':' で分割して開始・終了をパース
        auto s = CellAddress::from_a1(range.substr(0, colon));    // ':' より前
        auto e = CellAddress::from_a1(range.substr(colon + 1));   // ':' より後
        // 逆順の指定 (例: "Z50:A1") を正規化: min/max で常に小さい値を first にする
        pa.first_row = std::min(s.row, e.row);
        pa.first_col = std::min(s.col, e.col);
        pa.last_row  = std::max(s.row, e.row);
        pa.last_col  = std::max(s.col, e.col);
    }
    return pa;
}

// to_range(): PrintArea を "A1:Z50" のような範囲文字列に変換する
std::string PrintArea::to_range() const {
    // 開始アドレスを A1 記法に変換
    // + ":" + 終了アドレスを A1 記法に変換
    return CellAddress{first_row, first_col}.to_a1()
         + ":" +
           CellAddress{last_row, last_col}.to_a1();
}

} // namespace excellib
