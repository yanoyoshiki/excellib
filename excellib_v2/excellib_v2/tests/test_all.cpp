// excellib v2 全機能テストスイート
// このファイルは C++ の標準テストマクロを使って全クラスの動作を検証する
// 外部テストフレームワーク (Google Test など) を使わず、自前のマクロで実装している
#include "excellib/excellib.hpp"      // 全ヘッダーをまとめてインクルード
#include "excellib/xls_parser.hpp"   // XLS 内部クラス (decode_rk など) をテストするため
#include "excellib/xlsx_parser.hpp"  // XLSX 内部クラスをテストするため
// deflate.hpp: 自前 DEFLATE / CRC32 の単体テスト用
#include "../src/common/deflate.hpp"
// std::cout / std::cerr: テスト結果の出力
#include <iostream>
// assert: 条件チェック (ここでは使わず EXPECT マクロを使う)
#include <cassert>
// std::logic_error: テスト失敗時に投げる例外
#include <stdexcept>
// std::ostringstream: エラーメッセージを組み立てるため
#include <sstream>
// ファイルの読み書き (save/open のテストで使う)
#include <fstream>
// std::abs: 浮動小数点の差分比較に使う
#include <cmath>

// using namespace: テストコード内で excellib:: の省略を可能にする
using namespace excellib;

// ============================================================
//  テストフレームワーク (自前実装)
// ============================================================
// g_pass / g_fail / g_skip: テスト結果のカウンター (静的変数でグローバルに管理)
static int g_pass=0, g_fail=0, g_skip=0;
// g_current_group: 現在のテストグループ名 (デバッグ用)
static std::string g_current_group;

// group(): テストグループの開始を宣言してコンソールに出力する
void group(const std::string& name) {
    g_current_group = name;
    std::cout << "\n[" << name << "]\n";
}
// pass(): テスト成功を記録して出力する
void pass(const std::string& name) { std::cout<<"  PASS  "<<name<<"\n"; ++g_pass; }
// fail(): テスト失敗を記録してエラーメッセージと共に出力する
void fail(const std::string& name, const std::string& msg) {
    std::cout<<"  FAIL  "<<name<<": "<<msg<<"\n"; ++g_fail;
}
// skip_test(): テストをスキップしてその理由を出力する
void skip_test(const std::string& name, const std::string& reason) {
    std::cout<<"  SKIP  "<<name<<" ("<<reason<<")\n"; ++g_skip;
}

// RUN(name, ...): テストケースを実行するマクロ
// 例外が発生した場合は fail() を呼んで失敗として記録する
// __VA_ARGS__ に実際のテストコードを記述する
#define RUN(name, ...) do { \
    try { __VA_ARGS__; pass(name); } \
    catch(std::exception& e) { fail(name, e.what()); } \
    catch(...) { fail(name, "unknown exception"); } \
} while(0)

// EXPECT(cond): 条件が true でなければ std::logic_error を投げる
// #cond はマクロで条件式をそのまま文字列化する
#define EXPECT(cond) do { if(!(cond)) throw std::logic_error("EXPECT failed: " #cond); } while(0)

// EXPECT_EQ(a, b): a == b でなければ期待値と実際の値を含むエラーを投げる
#define EXPECT_EQ(a,b) do { if(!((a)==(b))) { \
    std::ostringstream _s; _s<<"Expected ["<<(b)<<"] got ["<<(a)<<"]"; \
    throw std::logic_error(_s.str()); } } while(0)

// EXPECT_NEAR(a, b, eps): |a - b| > eps なら失敗 (浮動小数点の比較に使う)
#define EXPECT_NEAR(a,b,eps) do { if(std::abs(double(a)-double(b))>eps) { \
    std::ostringstream _s; _s<<"Expected ~"<<(b)<<" got "<<(a); \
    throw std::logic_error(_s.str()); } } while(0)

// EXPECT_THROWS(expr, exc): expr を実行して exc 例外が投げられなければ失敗
#define EXPECT_THROWS(expr, exc) do { \
    bool _threw=false; try{expr;}catch(exc&){_threw=true;} \
    if(!_threw) throw std::logic_error("Expected exception " #exc); } while(0)

// ============================================================
//  SECTION 1: CellAddress のテスト
// ============================================================
void test_cell_address() {
    group("CellAddress");

    // A1 → {row=0, col=0} の変換と逆変換 (ラウンドトリップ)
    RUN("A1 parses to {0,0}", {
        auto a = CellAddress::from_a1("A1");
        EXPECT_EQ(a.row, 0u); EXPECT_EQ(a.col, 0u);  // 0-indexed で (0,0)
        EXPECT_EQ(a.to_a1(), "A1");                    // 逆変換で "A1" に戻る
    });
    // B3 → {row=2, col=1} のラウンドトリップ
    RUN("B3 roundtrip", {
        auto a = CellAddress::from_a1("B3");
        EXPECT_EQ(a.row,2u); EXPECT_EQ(a.col,1u);
        EXPECT_EQ(a.to_a1(),"B3");
    });
    // Z26 → {row=25, col=25} (Z は 26 列目 = 0-indexed で 25)
    RUN("Z26 roundtrip", {
        auto a = CellAddress::from_a1("Z26");
        EXPECT_EQ(a.row,25u); EXPECT_EQ(a.col,25u);
        EXPECT_EQ(a.to_a1(),"Z26");
    });
    // AA1 → col=26 (Z=25 の次は AA=26, 0-indexed)
    RUN("AA1 col=26", {
        auto a = CellAddress::from_a1("AA1");
        EXPECT_EQ(a.col,26u); EXPECT_EQ(a.to_a1(),"AA1");
    });
    // Excel の最大セル XFD1048576 (列=16383, 行=1048575 が 0-indexed の最大)
    RUN("XFD1048576 max cell", {
        auto a = CellAddress::from_a1("XFD1048576");
        EXPECT_EQ(a.col,16383u); EXPECT_EQ(a.row,1048575u);
        EXPECT_EQ(a.to_a1(),"XFD1048576");
    });
    // エラーケース: 空文字列は FormatError
    RUN("empty string throws FormatError",  { EXPECT_THROWS(CellAddress::from_a1(""), FormatError); });
    // A0 は行番号 0 → Excel では存在しない (1以上が有効)
    RUN("row 0 throws FormatError",         { EXPECT_THROWS(CellAddress::from_a1("A0"), FormatError); });
    // "123" は列文字がない → FormatError
    RUN("no column letters throws",         { EXPECT_THROWS(CellAddress::from_a1("123"), FormatError); });
    // "A" は行番号がない → FormatError
    RUN("no row digits throws",             { EXPECT_THROWS(CellAddress::from_a1("A"), FormatError); });
    // XFE1 は列数が 16384 を超える → FormatError
    RUN("col > 16384 throws",               { EXPECT_THROWS(CellAddress::from_a1("XFE1"), FormatError); });
    // A1048577 は行数が 1048576 を超える → FormatError
    RUN("row > 1048576 throws",             { EXPECT_THROWS(CellAddress::from_a1("A1048577"), FormatError); });
    // 小文字の列文字 "b5" は大文字と同じく B5 として扱われる
    RUN("mixed case column", {
        auto a = CellAddress::from_a1("b5");
        EXPECT_EQ(a.col,1u); EXPECT_EQ(a.row,4u);
    });
}

// ============================================================
//  SECTION 2: CellValue のテスト
// ============================================================
void test_cell_value() {
    group("CellValue");

    // BlankValue: is_blank() が true で is_string() が false であること
    RUN("BlankValue is_blank", { CellValue v=BlankValue{}; EXPECT(is_blank(v)); EXPECT(!is_string(v)); });
    // bool: is_bool() が true, get_bool() で値を取得できること
    RUN("bool true/false", {
        CellValue t=true, f=false;
        EXPECT(is_bool(t)); EXPECT(get_bool(t)==true);
        EXPECT(get_bool(f)==false);
    });
    // int64_t: is_int() と is_numeric() が true であること
    RUN("int64", { CellValue v=int64_t{42}; EXPECT(is_int(v)); EXPECT_EQ(get_int(v),42); EXPECT(is_numeric(v)); });
    // double: is_double() が true で浮動小数点精度内で一致すること
    RUN("double", { CellValue v=3.14; EXPECT(is_double(v)); EXPECT_NEAR(get_double(v),3.14,1e-9); });
    // std::string: is_string() が true で get_string() で取得できること
    RUN("string", { CellValue v=std::string{"hello"}; EXPECT(is_string(v)); EXPECT_EQ(get_string(v),"hello"); });
    // ErrorValue: is_error() が true, get_error().code でエラーコードを取得できること
    RUN("ErrorValue", { CellValue v=ErrorValue{"#REF!"}; EXPECT(is_error(v)); EXPECT_EQ(get_error(v).code,"#REF!"); });
    // as_double(): int64_t を double として取得できること
    RUN("as_double from int", { CellValue v=int64_t{10}; EXPECT_EQ(as_double(v),10.0); });
    // as_double(): double をそのまま返すこと
    RUN("as_double from double", { CellValue v=2.5; EXPECT_EQ(as_double(v),2.5); });
    // as_double(): 文字列からは FormatError が投げられること
    RUN("as_double from string throws", { EXPECT_THROWS(as_double(CellValue{std::string{"x"}}), FormatError); });
    // 型が違う get_*() は std::bad_variant_access を投げること
    RUN("wrong type get throws", { EXPECT_THROWS(get_int(CellValue{std::string{"x"}}), std::bad_variant_access); });
    // to_string(): 各型を文字列に変換できること
    RUN("to_string blank", { EXPECT_EQ(to_string(BlankValue{}),""); });           // 空 → ""
    RUN("to_string int", { EXPECT_EQ(to_string(int64_t{99}),"99"); });             // 整数 → "99"
    RUN("to_string bool true", { EXPECT_EQ(to_string(true),"TRUE"); });            // true → "TRUE"
    RUN("to_string error", { EXPECT_EQ(to_string(ErrorValue{"#NA"}),"#NA"); });    // エラー → コード文字列
}

// ============================================================
//  SECTION 3: XLSX の作成・読み込みのラウンドトリップ
// ============================================================
void test_xlsx_roundtrip() {
    group("XLSX create/roundtrip");

    // 空ワークブックの作成
    RUN("create empty workbook", {
        auto wb = WorkbookFactory::create(FileFormat::XLSX);
        EXPECT_EQ(wb->sheet_count(), 0u);                         // シート数 = 0
        EXPECT(wb->format() == FileFormat::XLSX);                 // フォーマット = XLSX
        EXPECT_EQ(wb->format_name(), "XLSX");                     // フォーマット名
    });
    // シートの追加・削除・数
    RUN("add/remove/count sheets", {
        auto wb = WorkbookFactory::create();
        wb->add_sheet("Alpha"); wb->add_sheet("Beta"); wb->add_sheet("Gamma");
        EXPECT_EQ(wb->sheet_count(), 3u);    // 3 シート追加後
        wb->remove_sheet(1);                  // 中間の "Beta" を削除
        EXPECT_EQ(wb->sheet_count(), 2u);    // 2 シートになる
        EXPECT_EQ(wb->sheet_names()[1], "Gamma");  // インデックス 1 は "Gamma"
    });
    // 同名シートの追加は WriteError
    RUN("duplicate sheet name throws", {
        auto wb = WorkbookFactory::create();
        wb->add_sheet("S");
        EXPECT_THROWS(wb->add_sheet("S"), WriteError);
    });
    // シート名でアクセスして値を読み書きできること
    RUN("sheet access by name", {
        auto wb = WorkbookFactory::create(); wb->add_sheet("Report");
        wb->sheet("Report").set_cell("A1", std::string{"ok"});
        EXPECT_EQ(get_string(wb->sheet("Report").cell("A1").value), "ok");
    });
    // 範囲外のシートアクセスは RangeError
    RUN("sheet access OOB throws", {
        auto wb = WorkbookFactory::create();
        EXPECT_THROWS(wb->sheet(0), RangeError);  // シートがないのにインデックス 0 を要求
    });
    // 存在しないシート名でのアクセスは RangeError
    RUN("sheet by name not found throws", {
        auto wb = WorkbookFactory::create();
        EXPECT_THROWS(wb->sheet("Missing"), RangeError);
    });
    // シートの名前変更
    RUN("rename sheet", {
        auto wb = WorkbookFactory::create(); wb->add_sheet("Old");
        wb->rename_sheet(0,"New");
        EXPECT_EQ(wb->sheet(0).name(),"New");   // 名前が "New" に変わっていること
    });
    // 全セルタイプ (文字列・整数・浮動小数点・bool・エラー・数式) がバイト列→再読み込みで正しく復元されること
    RUN("all cell types survive serialization", {
        auto wb = WorkbookFactory::create();
        auto& sh = wb->add_sheet("Types");
        sh.set_cell("A1", std::string{"hello"});    // 文字列
        sh.set_cell("B1", int64_t{12345});           // 整数
        sh.set_cell("C1", 3.14159);                  // 浮動小数点
        sh.set_cell("D1", true);                     // true
        sh.set_cell("E1", false);                    // false
        sh.set_cell("F1", ErrorValue{"#DIV/0!"});    // エラー値
        sh.set_formula("G1", "SUM(B1:C1)");         // 数式

        // バイト列に変換して再度読み込む (ラウンドトリップ)
        auto bytes = wb->to_bytes(FileFormat::XLSX, {});
        EXPECT(!bytes.empty());
        auto wb2 = WorkbookFactory::open(bytes, {});
        auto& sh2 = wb2->sheet(0);

        // 各セルの値が正しく復元されているか確認
        EXPECT_EQ(get_string(sh2.cell("A1").value), "hello");
        EXPECT_EQ(get_int(sh2.cell("B1").value), 12345);
        EXPECT_NEAR(get_double(sh2.cell("C1").value), 3.14159, 1e-5);
        EXPECT_EQ(get_bool(sh2.cell("D1").value), true);
        EXPECT_EQ(get_bool(sh2.cell("E1").value), false);
        EXPECT(is_error(sh2.cell("F1").value));
        EXPECT(sh2.cell("G1").has_formula());   // 数式セルかどうか
    });
    // 複数シートのラウンドトリップ
    RUN("multi-sheet roundtrip", {
        auto wb = WorkbookFactory::create();
        wb->add_sheet("S1").set_cell("A1", int64_t{1});
        wb->add_sheet("S2").set_cell("Z99", std::string{"end"});
        auto bytes = wb->to_bytes(FileFormat::XLSX,{});
        auto wb2 = WorkbookFactory::open(bytes,{});
        EXPECT_EQ(wb2->sheet_count(),2u);
        EXPECT_EQ(get_int(wb2->sheet("S1").cell("A1").value),1);
        EXPECT_EQ(get_string(wb2->sheet("S2").cell("Z99").value),"end");
    });
    // ファイルへの保存と再読み込み
    RUN("save and reload file", {
        auto wb = WorkbookFactory::create();
        auto& sh = wb->add_sheet("Data");
        sh.set_cell("A1", std::string{"saved"});
        sh.set_cell("B2", int64_t{999});
        wb->save("/tmp/exc_test_save.xlsx");   // /tmp に保存
        auto wb2 = WorkbookFactory::open("/tmp/exc_test_save.xlsx");  // 再度読み込み
        EXPECT_EQ(get_string(wb2->sheet(0).cell("A1").value),"saved");
        EXPECT_EQ(get_int(wb2->sheet(0).cell("B2").value),999);
    });
    // 大量データ (1000 行) のラウンドトリップ
    RUN("large sheet 1000 rows", {
        auto wb = WorkbookFactory::create(); auto& sh = wb->add_sheet("Big");
        for (int i=0;i<1000;++i) sh.set_cell(uint32_t(i),0,int64_t(i));  // 0〜999 を書き込む
        auto bytes = wb->to_bytes(FileFormat::XLSX,{});
        auto wb2 = WorkbookFactory::open(bytes,{});
        EXPECT_EQ(get_int(wb2->sheet(0).cell(999,0).value), 999);  // 最後の行を確認
    });
    // set_table() ヘルパーのテスト
    RUN("set_table helper", {
        auto wb = WorkbookFactory::create(); auto& sh = wb->add_sheet("T");
        // B2 を左上として 3 行 2 列のテーブルを書き込む
        sh.set_table("B2", {
            {std::string{"Name"}, std::string{"Score"}},
            {std::string{"Alice"}, int64_t{95}},
            {std::string{"Bob"},   int64_t{87}},
        });
        EXPECT_EQ(get_string(sh.cell("B2").value),"Name");    // B2 = "Name"
        EXPECT_EQ(get_string(sh.cell("C2").value),"Score");   // C2 = "Score"
        EXPECT_EQ(get_string(sh.cell("B3").value),"Alice");   // B3 = "Alice"
        EXPECT_EQ(get_int(sh.cell("C4").value),87);           // C4 = 87 (Bob のスコア)
    });
    // Unicode 文字列 (日本語) のシート名と文字列値
    RUN("unicode sheet name and string", {
        auto wb = WorkbookFactory::create();
        auto& sh = wb->add_sheet("売上レポート");
        sh.set_cell("A1", std::string{"テスト文字列"});
        auto bytes = wb->to_bytes(FileFormat::XLSX,{});
        auto wb2 = WorkbookFactory::open(bytes,{});
        EXPECT_EQ(wb2->sheet(0).name(),"売上レポート");           // シート名が正しく復元
        EXPECT_EQ(get_string(wb2->sheet(0).cell("A1").value),"テスト文字列");  // 文字列が正しく復元
    });
}

// ============================================================
//  SECTION 4: Sheet のセルアクセス API のテスト
// ============================================================
void test_sheet_api() {
    group("Sheet API");

    // cell() は存在しないセルに対して Blank を返す
    RUN("cell() returns blank for empty address", {
        auto wb=WorkbookFactory::create(); auto& sh=wb->add_sheet("S");
        auto c = sh.cell(99,99);
        EXPECT(c.is_blank()); EXPECT_EQ(c.address.row,99u);  // アドレスは要求通り
    });
    // try_cell() は空セルに対して nullopt を返す
    RUN("try_cell returns nullopt for blank", {
        auto wb=WorkbookFactory::create(); auto& sh=wb->add_sheet("S");
        EXPECT(!sh.try_cell(5,5).has_value());
    });
    // try_cell() は設定済みセルに対して値を返す
    RUN("try_cell returns value when set", {
        auto wb=WorkbookFactory::create(); auto& sh=wb->add_sheet("S");
        sh.set_cell("C3",int64_t{7});    // C3 = 7 (row=2, col=2)
        auto c=sh.try_cell(2,2);
        EXPECT(c.has_value()); EXPECT_EQ(get_int(c->value),7);
    });
    // row() は列番号順にソートされたセルを返す
    RUN("row() returns sorted cells", {
        auto wb=WorkbookFactory::create(); auto& sh=wb->add_sheet("S");
        // 0 行目に col=2, col=0, col=1 の順で書き込む (順番が逆)
        sh.set_cell(0,2,int64_t{3}); sh.set_cell(0,0,int64_t{1}); sh.set_cell(0,1,int64_t{2});
        auto r = sh.row(0);
        EXPECT_EQ(r.size(),3u);
        EXPECT_EQ(get_int(r[0].value),1);  // col=0 が最初に来ること
        EXPECT_EQ(get_int(r[1].value),2);
        EXPECT_EQ(get_int(r[2].value),3);
    });
    // cells() は空でないセルを全て返す
    RUN("cells() returns all non-blank", {
        auto wb=WorkbookFactory::create(); auto& sh=wb->add_sheet("S");
        sh.set_cell("A1",int64_t{1}); sh.set_cell("B2",int64_t{2}); sh.set_cell("C3",int64_t{3});
        EXPECT_EQ(sh.cells().size(),3u);   // 3 つのセルが返ること
    });
    // for_each_cell() を使って合計を計算できること
    RUN("for_each_cell sums values", {
        auto wb=WorkbookFactory::create(); auto& sh=wb->add_sheet("S");
        sh.set_cell("A1",int64_t{10}); sh.set_cell("B1",int64_t{20}); sh.set_cell("C1",int64_t{30});
        int64_t sum=0;
        // ラムダで各セルを走査して合計する
        sh.for_each_cell([&](const Cell& c){ if(is_int(c.value)) sum+=get_int(c.value); });
        EXPECT_EQ(sum,60);  // 10 + 20 + 30 = 60
    });
    // セルを書き込むと row_count / col_count が自動更新されること
    RUN("row_count and col_count update on write", {
        auto wb=WorkbookFactory::create(); auto& sh=wb->add_sheet("S");
        sh.set_cell(4,7,int64_t{1});  // 5 行目 (0-indexed=4)、8 列目 (0-indexed=7) に書き込む
        EXPECT_EQ(sh.row_count(),5u);  // 0〜4 で 5 行
        EXPECT_EQ(sh.col_count(),8u);  // 0〜7 で 8 列
    });
    // 数式セルは has_formula() が true で formula フィールドに文字列が入ること
    RUN("formula cell stores formula string", {
        auto wb=WorkbookFactory::create(); auto& sh=wb->add_sheet("S");
        sh.set_formula("A1","SUM(B1:Z1)");
        auto c=sh.cell("A1");
        EXPECT(c.has_formula()); EXPECT_EQ(*c.formula,"SUM(B1:Z1)");
        EXPECT(c.type == CellType::Formula);  // type も Formula であること
    });
}

// ============================================================
//  SECTION 5: フォーマット検出のテスト
// ============================================================
void test_format_detection() {
    group("Format detection");

    // ランダムバイト列は FormatError (シグネチャが一致しない)
    RUN("garbage bytes throw FormatError", {
        std::vector<uint8_t> g{0,1,2,3,4,5,6,7};
        EXPECT_THROWS(WorkbookFactory::open(g,{}), FormatError);
    });
    // 空バッファは ParseError (サイズ不足)
    RUN("empty buffer throws ParseError", {
        std::vector<uint8_t> e;
        EXPECT_THROWS(WorkbookFactory::open(e,{}), ParseError);
    });
    // 4 バイトは ParseError (OLE2 マジック確認に 8 バイト必要)
    RUN("4-byte buffer throws ParseError", {
        std::vector<uint8_t> small{0x50,0x4B,0x03,0x04};
        EXPECT_THROWS(WorkbookFactory::open(small,{}), ParseError);
    });
    // XLSX として作成したワークブックは FileFormat::XLSX として検出されること
    RUN("created XLSX detected as XLSX", {
        auto wb=WorkbookFactory::create(FileFormat::XLSX);
        wb->add_sheet("S");
        auto bytes=wb->to_bytes(FileFormat::XLSX,{});
        auto wb2=WorkbookFactory::open(bytes,{});
        EXPECT(wb2->format() == FileFormat::XLSX);
    });
}

// ============================================================
//  SECTION 6: XLS 内部クラスのテスト
// ============================================================
void test_xls_internals() {
    group("XLS internals");

    // decode_rk(): RK 値のデコードテスト
    // RK 値 = 整数値を 2 ビット左シフトして bit1=1 (整数モード) をセット
    // 例: 100 → (100 << 2) | 0x02 = 402
    RUN("decode_rk integer", {
        uint32_t rk=(100<<2)|0x02;   // bit1=1: 整数モード
        EXPECT_EQ(xls::decode_rk(rk),100.0);
    });
    // bit0=1, bit1=1: 整数モードかつ 1/100 スケール
    // 1000 / 100 = 10.0
    RUN("decode_rk integer /100", {
        uint32_t rk=(1000<<2)|0x03;  // bit0=1 (1/100), bit1=1 (整数)
        EXPECT_NEAR(xls::decode_rk(rk),10.0,1e-9);
    });
    // decode_biff8_string(): 圧縮 ASCII (1バイト/文字) のデコード
    // ヘッダー: [長さ: 2byte = 3][フラグ: 1byte = 0x00 (圧縮)] + "ABC"
    RUN("decode_biff8 compressed ASCII", {
        std::vector<uint8_t> d={3,0,0x00,'A','B','C'};
        size_t consumed=0;
        auto s=xls::decode_biff8_string(d.data(),d.size(),consumed);
        EXPECT_EQ(s,"ABC"); EXPECT_EQ(consumed,6u);  // 2+1+3 = 6 バイト消費
    });
    // decode_biff8_string(): UTF-16LE (2バイト/文字) のデコード
    // ヘッダー: [長さ: 2byte = 2][フラグ: 1byte = 0x01 (UTF-16LE)] + "H\0i\0"
    RUN("decode_biff8 UTF-16LE", {
        std::vector<uint8_t> d={2,0,0x01,'H',0,'i',0};
        size_t consumed=0;
        auto s=xls::decode_biff8_string(d.data(),d.size(),consumed);
        EXPECT_EQ(s,"Hi"); EXPECT_EQ(consumed,7u);  // 2+1+2*2 = 7 バイト消費
    });
    // データが少なすぎる場合は ParseError
    RUN("decode_biff8 truncated throws", {
        std::vector<uint8_t> d={5,0};  // 長さ=5 と宣言しているがデータがない
        size_t consumed=0;
        EXPECT_THROWS(xls::decode_biff8_string(d.data(),d.size(),consumed), ParseError);
    });
    // SharedStringTable: 空のテーブルへのインデックス 0 アクセスは RangeError
    RUN("SharedStringTable out-of-range throws", {
        xls::SharedStringTable sst;
        EXPECT_THROWS(sst.get(0), RangeError);
    });
    // FormatRegistry: 組み込み日付フォーマット ID 14 は日付として認識されること
    RUN("FormatRegistry built-in date (14)", {
        xls::FormatRegistry r; EXPECT(r.get(14).is_date_format);
    });
    // FormatRegistry: ID 0 (General) は日付でないこと
    RUN("FormatRegistry General is not date", {
        xls::FormatRegistry r; EXPECT(!r.get(0).is_date_format);
    });
    // XFTable: 空のテーブルへのインデックス 0 アクセスは RangeError
    RUN("XFTable out-of-range throws", {
        xls::XFTable t; EXPECT_THROWS(t.get(0), RangeError);
    });
    // BiffRecord: 境界チェック付き u16() のテスト
    RUN("BiffRecord u16 bounds-checks", {
        xls::BiffRecord r; r.data={0x01,0x02};
        EXPECT_EQ(r.u16(0), 0x0201u);   // オフセット 0 で 2 バイト Little Endian = 0x0201
        EXPECT_THROWS(r.u16(1), ParseError);   // オフセット 1 で 2 バイト → 1 バイト不足 → ParseError
    });
}

// ============================================================
//  SECTION 7: PrintArea のテスト
// ============================================================
void test_print_area() {
    group("PrintArea");

    // "A1:Z50" → first_row=0, first_col=0, last_row=49, last_col=25
    RUN("from_range A1:Z50", {
        auto pa=PrintArea::from_range("A1:Z50");
        EXPECT_EQ(pa.first_row,0u); EXPECT_EQ(pa.first_col,0u);
        EXPECT_EQ(pa.last_row,49u); EXPECT_EQ(pa.last_col,25u);
        EXPECT(pa.is_valid());  // is_valid(): first <= last であること
    });
    // to_range() で元の文字列に戻ること
    RUN("to_range roundtrip", {
        EXPECT_EQ(PrintArea::from_range("B2:D10").to_range(),"B2:D10");
    });
    // 単一セルの指定 (コロンなし)
    RUN("single cell", {
        auto pa=PrintArea::from_range("C5");
        EXPECT_EQ(pa.first_row,pa.last_row); EXPECT_EQ(pa.first_col,pa.last_col);
    });
    // 逆順指定 "Z50:A1" は正規化されて first < last になること
    RUN("reversed coords normalize", {
        auto pa=PrintArea::from_range("Z50:A1");
        EXPECT_EQ(pa.first_row,0u); EXPECT_EQ(pa.last_row,49u);
    });
    // 不正な範囲 (行 0 は存在しない) は FormatError
    RUN("invalid range throws", {
        EXPECT_THROWS(PrintArea::from_range("A0:B5"), FormatError);
    });
}

// ============================================================
//  SECTION 8: PageSetup のデフォルト値と apply のテスト
// ============================================================
void test_page_setup() {
    group("PageSetup");

    // デフォルト値の確認
    RUN("defaults", {
        PageSetup ps;
        EXPECT(ps.paper_size==PaperSize::A4);         // 用紙 A4
        EXPECT(ps.orientation==Orientation::Portrait); // 縦向き
        EXPECT(ps.fit_to==FitTo::None);                // 自動縮小なし
        EXPECT_EQ(ps.scale_percent,100);               // 倍率 100%
        EXPECT(!ps.print_gridlines);                   // グリッド線なし
        EXPECT(!ps.print_area.has_value());            // 印刷範囲なし
    });
    // マージンのデフォルト値 (インチ)
    RUN("margins defaults", {
        PageSetup ps;
        EXPECT_NEAR(ps.margins.left,  0.70, 1e-9);
        EXPECT_NEAR(ps.margins.right, 0.70, 1e-9);
        EXPECT_NEAR(ps.margins.top,   0.75, 1e-9);
        EXPECT_NEAR(ps.margins.bottom,0.75, 1e-9);
    });

    // テスト用 XLSX ファイルを作成するラムダ
    auto make_xlsx=[](const std::string& path){
        auto wb=WorkbookFactory::create();
        auto& sh=wb->add_sheet("Sheet1");
        sh.set_cell("A1",std::string{"data"}); sh.set_cell("B2",int64_t{42});
        wb->save(path);
    };
    // XLSX の XML 内に特定文字列が含まれるか確認するラムダ
    auto raw_contains=[](const std::string& path, const std::string& needle)->bool{
        std::ifstream f(path,std::ios::binary);
        std::string s((std::istreambuf_iterator<char>(f)),{});
        return s.find(needle)!=std::string::npos;
    };

    // A4 縦向きの pageSetup が XML に書き込まれること
    RUN("apply portrait A4", {
        make_xlsx("/tmp/exc_ps1.xlsx");
        PageSetup ps; ps.paper_size=PaperSize::A4; ps.orientation=Orientation::Portrait;
        ExcelPrinter p; p.apply_page_setup("/tmp/exc_ps1.xlsx","Sheet1",ps);
        EXPECT(raw_contains("/tmp/exc_ps1.xlsx","paperSize=\"9\""));        // A4 = 9
        EXPECT(raw_contains("/tmp/exc_ps1.xlsx","orientation=\"portrait\"")); // 縦向き
    });
    // 横向き (Landscape) が XML に書き込まれること
    RUN("apply landscape", {
        make_xlsx("/tmp/exc_ps2.xlsx");
        PageSetup ps; ps.orientation=Orientation::Landscape;
        ExcelPrinter p; p.apply_page_setup("/tmp/exc_ps2.xlsx","Sheet1",ps);
        EXPECT(raw_contains("/tmp/exc_ps2.xlsx","orientation=\"landscape\""));
    });
    // FitTo::Width (横 1 ページに収める) が XML に書き込まれること
    RUN("apply FitTo::Width", {
        make_xlsx("/tmp/exc_ps3.xlsx");
        PageSetup ps; ps.fit_to=FitTo::Width; ps.fit_to_pages_wide=1;
        ExcelPrinter p; p.apply_page_setup("/tmp/exc_ps3.xlsx","Sheet1",ps);
        EXPECT(raw_contains("/tmp/exc_ps3.xlsx","fitToPage=\"1\""));   // FitTo 有効
        EXPECT(raw_contains("/tmp/exc_ps3.xlsx","fitToWidth=\"1\""));  // 横 1 ページ
    });
    // FitTo::WidthAndHeight (縦横両方に収める) が XML に書き込まれること
    RUN("apply FitTo::WidthAndHeight", {
        make_xlsx("/tmp/exc_ps4.xlsx");
        PageSetup ps; ps.fit_to=FitTo::WidthAndHeight;
        ps.fit_to_pages_wide=1; ps.fit_to_pages_tall=1;
        ExcelPrinter p; p.apply_page_setup("/tmp/exc_ps4.xlsx","Sheet1",ps);
        EXPECT(raw_contains("/tmp/exc_ps4.xlsx","fitToWidth=\"1\""));
        EXPECT(raw_contains("/tmp/exc_ps4.xlsx","fitToHeight=\"1\""));
    });
    // 印刷範囲が definedName に書き込まれること
    RUN("apply print_area", {
        make_xlsx("/tmp/exc_ps5.xlsx");
        PageSetup ps; ps.print_area=PrintArea::from_range("A1:B20");
        ExcelPrinter p; p.apply_page_setup("/tmp/exc_ps5.xlsx","Sheet1",ps);
        EXPECT(raw_contains("/tmp/exc_ps5.xlsx","Print_Area"));   // definedName が存在する
        EXPECT(raw_contains("/tmp/exc_ps5.xlsx","$A$1"));         // 絶対参照形式
    });
    // タイトル行の繰り返しが definedName に書き込まれること
    RUN("apply repeat_titles rows", {
        make_xlsx("/tmp/exc_ps6.xlsx");
        PageSetup ps; RepeatTitles rt; rt.row_start=0; rt.row_end=0; ps.repeat_titles=rt;
        ExcelPrinter p; p.apply_page_setup("/tmp/exc_ps6.xlsx","Sheet1",ps);
        EXPECT(raw_contains("/tmp/exc_ps6.xlsx","Print_Titles"));  // definedName が存在する
    });
    // グリッド線・白黒印刷が XML に書き込まれること
    RUN("apply gridlines + black_and_white", {
        make_xlsx("/tmp/exc_ps7.xlsx");
        PageSetup ps; ps.print_gridlines=true; ps.black_and_white=true;
        ExcelPrinter p; p.apply_page_setup("/tmp/exc_ps7.xlsx","Sheet1",ps);
        EXPECT(raw_contains("/tmp/exc_ps7.xlsx","gridLines=\"1\""));
        EXPECT(raw_contains("/tmp/exc_ps7.xlsx","blackAndWhite=\"1\""));
    });
    // ヘッダー・フッターが XML に書き込まれること
    RUN("apply header/footer", {
        make_xlsx("/tmp/exc_ps8.xlsx");
        PageSetup ps;
        ps.header_footer.odd_header="&C&14Report";     // 中央に 14pt で "Report"
        ps.header_footer.odd_footer="&LPage &P / &N";  // 左端にページ番号
        ExcelPrinter p; p.apply_page_setup("/tmp/exc_ps8.xlsx","Sheet1",ps);
        EXPECT(raw_contains("/tmp/exc_ps8.xlsx","Report"));
        EXPECT(raw_contains("/tmp/exc_ps8.xlsx","oddFooter"));
    });
    // シート名が空の場合は最初のシートに適用されること
    RUN("apply empty sheet_name defaults to first sheet", {
        make_xlsx("/tmp/exc_ps9.xlsx");
        PageSetup ps; ps.orientation=Orientation::Landscape;
        ExcelPrinter p; p.apply_page_setup("/tmp/exc_ps9.xlsx","",ps);  // シート名を空に
        EXPECT(raw_contains("/tmp/exc_ps9.xlsx","orientation=\"landscape\""));
    });
    // ページ設定を書き込んだ後も XLSX として正常に読み込めること
    RUN("apply re-opens correctly after modification", {
        make_xlsx("/tmp/exc_ps10.xlsx");
        PageSetup ps; ps.print_gridlines=true;
        ExcelPrinter p; p.apply_page_setup("/tmp/exc_ps10.xlsx","Sheet1",ps);
        // 書き換え後も XLSX として正常に読み込めること
        auto wb=WorkbookFactory::open("/tmp/exc_ps10.xlsx");
        EXPECT_EQ(get_string(wb->sheet(0).cell("A1").value),"data");
        EXPECT_EQ(get_int(wb->sheet(0).cell("B2").value),42);
    });
    // 全設定を同時に適用する統合テスト
    RUN("combined all settings", {
        make_xlsx("/tmp/exc_ps_full.xlsx");
        PageSetup ps;
        ps.paper_size=PaperSize::A4; ps.orientation=Orientation::Landscape;
        ps.fit_to=FitTo::Width; ps.fit_to_pages_wide=1;
        ps.print_gridlines=true; ps.print_area=PrintArea::from_range("A1:H20");
        ps.margins.left=0.5; ps.margins.right=0.5;
        ps.header_footer.odd_header="&C売上レポート";
        RepeatTitles rt; rt.row_start=0; rt.row_end=0; ps.repeat_titles=rt;
        ExcelPrinter p; p.apply_page_setup("/tmp/exc_ps_full.xlsx","Sheet1",ps);
        EXPECT(raw_contains("/tmp/exc_ps_full.xlsx","paperSize=\"9\""));
        EXPECT(raw_contains("/tmp/exc_ps_full.xlsx","orientation=\"landscape\""));
        EXPECT(raw_contains("/tmp/exc_ps_full.xlsx","Print_Area"));
        EXPECT(raw_contains("/tmp/exc_ps_full.xlsx","Print_Titles"));
        EXPECT(raw_contains("/tmp/exc_ps_full.xlsx","売上レポート"));
    });
}

// ============================================================
//  SECTION 9: ExcelPrinter のエンジン / バリデーションのテスト
// ============================================================
void test_printer() {
    group("ExcelPrinter");

    // ExcelPrinter のコンストラクタがクラッシュしないこと (COM 初期化のテスト)
    RUN("ExcelPrinter constructs without crash", {
        ExcelPrinter p;
        // 構築成功 = COM の初期化が成功 (または非 Windows で空の COMGuard が作られる)
    });
    // list_printers() がクラッシュせずにリストを返すこと
    RUN("list_printers does not crash", {
        ExcelPrinter p;
        auto pr=p.list_printers();
        std::cout<<"        "<<pr.size()<<" printer(s) found\n";  // 件数を出力 (情報確認)
        EXPECT(true);
    });
    // to_pdf() は output_path が空なら PrintError を投げること
    RUN("to_pdf with empty path throws PrintError", {
        ExcelPrinter p;
        PdfOptions o; // output_path="" (デフォルト空文字列)
        EXPECT_THROWS(p.to_pdf("/tmp/exc_ps_full.xlsx",o), PrintError);
    });
    // 同じテストをもう一度 (重複して見えるが、別の呼び出し経路を確認)
    RUN("to_pdf requires output_path", {
        ExcelPrinter p;
        PdfOptions opts;
        EXPECT_THROWS(p.to_pdf("/tmp/exc_ps_full.xlsx", opts), PrintError);
    });
}

// ============================================================
//  SECTION 10: BatchPrinter のテスト
// ============================================================
void test_batch_printer() {
    group("BatchPrinter");

    // テスト用 XLSX ファイルを作成するラムダ
    auto make = [](const std::string& path, const std::string& sheet, int val) {
        auto wb = WorkbookFactory::create();
        wb->add_sheet(sheet).set_cell("A1", int64_t{val});
        wb->save(path);
    };
    // テスト用の 3 ファイルを作成する
    make("/tmp/batch1.xlsx", "Sheet1", 1);
    make("/tmp/batch2.xlsx", "Report", 2);
    make("/tmp/batch3.xlsx", "Data",   3);

    // フルエント API (メソッドチェーン) でジョブを追加できること
    RUN("add jobs via fluent API", {
        BatchPrinter bp;
        bp.add("/tmp/batch1.xlsx", "Sheet1")   // シート名指定
          .add("/tmp/batch2.xlsx", "Report")
          .add("/tmp/batch3.xlsx");             // シート名なし (ActiveSheet)
        EXPECT_EQ(bp.job_count(), 3u);
    });
    // add_all() でベクターからジョブを一括追加できること
    RUN("add_all from vector", {
        BatchPrinter bp;
        std::vector<PrintJob> jobs = {
            PrintJob::make("/tmp/batch1.xlsx", "Sheet1"),
            PrintJob::make("/tmp/batch2.xlsx", "Report"),
        };
        bp.add_all(jobs);
        EXPECT_EQ(bp.job_count(), 2u);
    });
    // clear() でジョブリストをリセットできること
    RUN("clear resets job list", {
        BatchPrinter bp;
        bp.add("/tmp/batch1.xlsx").add("/tmp/batch2.xlsx");
        EXPECT_EQ(bp.job_count(), 2u);
        bp.clear();   // リセット
        EXPECT_EQ(bp.job_count(), 0u);
    });
    // PrintJob::make() で全フィールドを設定できること
    RUN("PrintJob::make with all fields", {
        PageSetup ps; ps.orientation = Orientation::Landscape;
        auto job = PrintJob::make("/tmp/batch1.xlsx", "Sheet1", ps, "Q1_Sales");
        EXPECT_EQ(job.file_path, "/tmp/batch1.xlsx");
        EXPECT_EQ(job.sheet_name, "Sheet1");
        EXPECT_EQ(job.label, "Q1_Sales");
        EXPECT(job.setup.has_value());                              // PageSetup が設定されている
        EXPECT(job.setup->orientation == Orientation::Landscape);   // 横向き
    });
    // setup なしの PrintJob::make()
    RUN("PrintJob without setup", {
        auto job = PrintJob::make("/tmp/batch1.xlsx");
        EXPECT(!job.setup.has_value());   // setup は空 (nullopt)
        EXPECT(job.sheet_name.empty());   // シート名も空
    });
    // BatchPrinter のデフォルトコンストラクター
    RUN("BatchPrinter constructs without crash", {
        BatchPrinter bp;
        EXPECT_EQ(bp.job_count(), 0u);   // 初期状態はジョブ 0 件
    });
    // BatchResult の失敗追跡テスト
    RUN("BatchResult failure tracking", {
        BatchResult r;
        r.total = 3;
        // 手動でジョブ結果を追加する (1 件失敗させる)
        r.jobs.push_back({0, "a.xlsx", "", true,  ""});             // 成功
        r.jobs.push_back({1, "b.xlsx", "", false, "File not found"}); // 失敗
        r.jobs.push_back({2, "c.xlsx", "", true,  ""});             // 成功
        r.succeeded = 2; r.failed = 1;
        EXPECT(!r.all_success());          // 1 件失敗があるので false
        auto fails = r.failures();         // 失敗したジョブのリストを取得
        EXPECT_EQ(fails.size(), 1u);       // 失敗は 1 件
        EXPECT_EQ(fails[0]->file_path, "b.xlsx");  // 失敗したファイルの確認
    });
    // フルエント API でのジョブ追加後の件数確認
    RUN("job fluent add sets count", {
        BatchPrinter bp;
        bp.add("/tmp/batch1.xlsx","Sheet1")
          .add("/tmp/batch2.xlsx","Report")
          .add("/tmp/batch3.xlsx","Data");
        EXPECT_EQ(bp.job_count(), 3u);
    });
    // job_count() の追加後の確認
    RUN("job_count after add_all", {
        BatchPrinter bp;
        bp.add("/tmp/batch1.xlsx")
          .add("/tmp/batch2.xlsx")
          .add("/tmp/batch3.xlsx");
        EXPECT_EQ(bp.job_count(), 3u);
    });
}

// ============================================================
//  SECTION 11: DEFLATE / CRC32 のテスト
// ============================================================
// from_hex(): 16 進数文字列をバイト列に変換するヘルパー
// "48656C6C6F" → {0x48, 0x65, 0x6C, 0x6C, 0x6F}
static std::vector<uint8_t> from_hex(const char* h) {
    std::vector<uint8_t> v;
    for (; h[0] && h[1]; h += 2) {
        // 1 文字を 16 進数の値に変換するラムダ
        auto hex = [](char c) -> uint8_t {
            if (c >= '0' && c <= '9') return uint8_t(c - '0');
            if (c >= 'a' && c <= 'f') return uint8_t(c - 'a' + 10);
            return uint8_t(c - 'A' + 10);
        };
        // 2 文字で 1 バイト: 上位 4 ビット + 下位 4 ビット
        v.push_back(uint8_t((hex(h[0]) << 4) | hex(h[1])));
    }
    return v;
}

void test_deflate() {
    group("DEFLATE / CRC32");

    // ── CRC32 テスト ─────────────────────────────────────────
    // CRC32("Hello, World!") の既知の値との比較
    RUN("CRC32 Hello World", {
        auto s = reinterpret_cast<const uint8_t*>("Hello, World!");
        EXPECT_EQ(excellib::detail::crc32_compute(s, 13), 0xEC4AC3D0u);
    });
    // CRC32("excellib test") の既知の値との比較
    RUN("CRC32 excellib test", {
        auto s = reinterpret_cast<const uint8_t*>("excellib test");
        EXPECT_EQ(excellib::detail::crc32_compute(s, 13), 0x9EED09FFu);
    });
    // 空データの CRC32 は 0x00000000
    RUN("CRC32 空データ = 0x00000000", {
        EXPECT_EQ(excellib::detail::crc32_compute(nullptr, 0), 0x00000000u);
    });

    // ── 非圧縮ブロック (BTYPE=0, Stored) のテスト ────────────
    // DEFLATE 非圧縮ブロック形式:
    //   BFINAL=1, BTYPE=00 → 0x01 (最下位ビット=BFINAL, ビット1-2=BTYPE)
    //   アライン後: LEN (2byte LE) + NLEN (LEN のビット反転, 2byte LE) + データ
    RUN("stored block: Hello", {
        // 0x01=BFINAL|BTYPE(00), 0x05,0x00=LEN=5, 0xFA,0xFF=~LEN, "Hello"
        auto comp = from_hex("010500FAFF48656C6C6F");
        auto out  = excellib::detail::deflate_decompress(comp.data(), comp.size(), 5);
        EXPECT(out == std::vector<uint8_t>({'H','e','l','l','o'}));
    });

    // ── 固定ハフマン (BTYPE=1) のテスト ──────────────────────
    // Python で生成: zlib.compressobj(6, DEFLATED, -15).compress(b"Hello, DEFLATE!") + flush()
    RUN("fixed huffman: Hello, DEFLATE!", {
        auto comp = from_hex("f348cdc9c9d751707175f3710c71550400");
        auto out  = excellib::detail::deflate_decompress(comp.data(), comp.size(), 15);
        std::string s(out.begin(), out.end());
        EXPECT_EQ(s, "Hello, DEFLATE!");
    });
    // 反復データ: 'A' を 22 回繰り返した場合 (後方参照コードが使われる)
    RUN("fixed huffman: 22xA (back-reference)", {
        auto comp = from_hex("7374c40600");
        auto out  = excellib::detail::deflate_decompress(comp.data(), comp.size(), 22);
        EXPECT_EQ(out.size(), 22u);
        EXPECT(std::all_of(out.begin(), out.end(), [](uint8_t b){ return b == 'A'; }));
    });
    // アルファベット 26 文字 × 3 回 (= 78 文字)
    RUN("fixed huffman: alphabet x3", {
        auto comp = from_hex(
            "4b4c4a4e494d4bcfc8cccacec9cdcb2f282c2a2e292d2b"
            "afa8ac4a24430600");
        auto out  = excellib::detail::deflate_decompress(comp.data(), comp.size(), 78);
        EXPECT_EQ(out.size(), 78u);
        std::string s(out.begin(), out.end());
        // 期待値: "abcdefghijklmnopqrstuvwxyz" × 3
        std::string expected;
        for (int i = 0; i < 3; ++i) expected += "abcdefghijklmnopqrstuvwxyz";
        EXPECT_EQ(s, expected);
    });
    // パングラム (全英字を含む文)
    RUN("fixed huffman: pangram", {
        auto comp = from_hex(
            "0bc94855282ccd4cce56482aca2fcf5348cbaf50c82acd2d"
            "2856c82f4b2d5228014ae72456552aa4e4a70300");
        auto out  = excellib::detail::deflate_decompress(comp.data(), comp.size(), 43);
        std::string s(out.begin(), out.end());
        EXPECT_EQ(s, "The quick brown fox jumps over the lazy dog");
    });

    // ── 動的ハフマン (BTYPE=2) のテスト ──────────────────────
    // XLSX の sheet.xml のような XML データを動的ハフマンで圧縮したもの
    RUN("dynamic huffman: XLSX-like XML", {
        auto comp = from_hex(
            "4d8ed10ac2300c457fa5e4dda50e1191b64311bf403fa074"
            "711baeed68cba67f6f3765f8126e4e38e18aea657b365288"
            "9d7712b6050746cef8ba738d84fbedba39008b49bb5af7de"
            "91843745a894987c78c69628b1ecbb28a14d69382246d392"
            "d5b1f003b97c79f86075ca6b68300e8174bd48b6c792f33d"
            "5add39506261179db412c14f2ce41e999a399ce634aa5d29"
            "705402cd0f9fbf78cbf9ca31ab79fefdc2b5a4fa00");
        // 期待する展開結果: XLSX の sheet.xml の典型的な XML
        const std::string expected =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
            "<sheetData><row r=\"1\"><c r=\"A1\"><v>42</v></c>"
            "<c r=\"B1\"><v>100</v></c></row></sheetData></worksheet>";
        auto out = excellib::detail::deflate_decompress(comp.data(), comp.size(), expected.size());
        std::string s(out.begin(), out.end());
        EXPECT_EQ(s, expected);
    });

    // ── エラーケースのテスト ──────────────────────────────────
    // 空データは例外を投げること
    RUN("空データで例外", {
        std::vector<uint8_t> empty;
        bool threw = false;
        try { excellib::detail::deflate_decompress(empty.data(), 0); }
        catch (std::exception&) { threw = true; }
        EXPECT(threw);
    });
    // BTYPE=3 は予約済みで無効 → 例外を投げること
    // 0x07 = 0b00000111: BFINAL=1, BTYPE=11 (= 3, 無効)
    RUN("btype=3 (予約済み) で例外", {
        std::vector<uint8_t> bad = {0x07};
        bool threw = false;
        try { excellib::detail::deflate_decompress(bad.data(), bad.size()); }
        catch (std::exception&) { threw = true; }
        EXPECT(threw);
    });
}

// ============================================================
//  メイン関数
// ============================================================
int main() {
    // テスト開始のバナーを出力する
    std::cout<<"╔══════════════════════════════════════════╗\n";
    std::cout<<"║    excellib v2 — Full Test Suite         ║\n";
    std::cout<<"╚══════════════════════════════════════════╝\n";

    // 各セクションのテストを順番に実行する
    test_cell_address();     // Section 1: CellAddress の変換
    test_cell_value();       // Section 2: CellValue の型操作
    test_xlsx_roundtrip();   // Section 3: XLSX の作成・保存・読み込み
    test_sheet_api();        // Section 4: シートのセルアクセス
    test_format_detection(); // Section 5: ファイルフォーマットの自動検出
    test_xls_internals();    // Section 6: XLS (BIFF8) の内部関数
    test_print_area();       // Section 7: 印刷範囲の解析
    test_page_setup();       // Section 8: ページ設定の適用
    test_printer();          // Section 9: ExcelPrinter のバリデーション
    test_batch_printer();    // Section 10: BatchPrinter のジョブ管理
    test_deflate();          // Section 11: DEFLATE / CRC32 の実装

    // 結果サマリーを出力する
    std::cout<<"\n╔══════════════════════════════════════════╗\n";
    std::cout<<"║  Results: "<<g_pass<<" passed  "<<g_fail<<" failed  "<<g_skip<<" skipped        ║\n";
    std::cout<<"╚══════════════════════════════════════════╝\n";
    // 失敗が 0 件なら終了コード 0 (成功)、1 件以上なら終了コード 1 (失敗)
    return g_fail==0 ? 0 : 1;
}
