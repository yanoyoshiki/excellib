#include "excellib/excellib.hpp"
#include "excellib/xls_parser.hpp"
#include "excellib/xlsx_parser.hpp"
#include "../src/common/deflate.hpp"
#include <iostream>
#include <cassert>
#include <stdexcept>
#include <sstream>
#include <fstream>
#include <cmath>
#include <algorithm>
#include <filesystem>
#ifdef _WIN32
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#endif

// クロスプラットフォームの一時ディレクトリパスを取得するヘルパー
static std::string tmp(const std::string& name) {
    return (std::filesystem::temp_directory_path() / name).string();
}

using namespace excellib;

// ============================================================
//  Test framework
// ============================================================
static int g_pass=0, g_fail=0, g_skip=0;
static std::string g_current_group;

void group(const std::string& name) {
    g_current_group = name;
    std::cout << "\n[" << name << "]\n";
}
void pass(const std::string& name) { std::cout<<"  PASS  "<<name<<"\n"; ++g_pass; }
void fail(const std::string& name, const std::string& msg) {
    std::cout<<"  FAIL  "<<name<<": "<<msg<<"\n"; ++g_fail;
}
void skip_test(const std::string& name, const std::string& reason) {
    std::cout<<"  SKIP  "<<name<<" ("<<reason<<")\n"; ++g_skip;
}

#define RUN(name, ...) do { \
    try { __VA_ARGS__; pass(name); } \
    catch(std::exception& e) { fail(name, e.what()); } \
    catch(...) { fail(name, "unknown exception"); } \
} while(0)

#define EXPECT(cond) do { if(!(cond)) throw std::logic_error("EXPECT failed: " #cond); } while(0)
#define EXPECT_EQ(a,b) do { if(!((a)==(b))) { \
    std::ostringstream _s; _s<<"Expected ["<<(b)<<"] got ["<<(a)<<"]"; \
    throw std::logic_error(_s.str()); } } while(0)
#define EXPECT_NEAR(a,b,eps) do { if(std::abs(double(a)-double(b))>eps) { \
    std::ostringstream _s; _s<<"Expected ~"<<(b)<<" got "<<(a); \
    throw std::logic_error(_s.str()); } } while(0)
#define EXPECT_THROWS(expr, exc) do { \
    bool _threw=false; try{expr;}catch(exc&){_threw=true;} \
    if(!_threw) throw std::logic_error("Expected exception " #exc); } while(0)

// ============================================================
//  SECTION 1: CellAddress
// ============================================================
void test_cell_address() {
    group("CellAddress");

    RUN("A1 parses to {0,0}", {
        auto a = CellAddress::from_a1("A1");
        EXPECT_EQ(a.row, 0u); EXPECT_EQ(a.col, 0u);
        EXPECT_EQ(a.to_a1(), "A1");
    });
    RUN("B3 roundtrip", {
        auto a = CellAddress::from_a1("B3");
        EXPECT_EQ(a.row,2u); EXPECT_EQ(a.col,1u);
        EXPECT_EQ(a.to_a1(),"B3");
    });
    RUN("Z26 roundtrip", {
        auto a = CellAddress::from_a1("Z26");
        EXPECT_EQ(a.row,25u); EXPECT_EQ(a.col,25u);
        EXPECT_EQ(a.to_a1(),"Z26");
    });
    RUN("AA1 col=26", {
        auto a = CellAddress::from_a1("AA1");
        EXPECT_EQ(a.col,26u); EXPECT_EQ(a.to_a1(),"AA1");
    });
    RUN("XFD1048576 max cell", {
        auto a = CellAddress::from_a1("XFD1048576");
        EXPECT_EQ(a.col,16383u); EXPECT_EQ(a.row,1048575u);
        EXPECT_EQ(a.to_a1(),"XFD1048576");
    });
    RUN("empty string throws FormatError",  { EXPECT_THROWS(CellAddress::from_a1(""), FormatError); });
    RUN("row 0 throws FormatError",         { EXPECT_THROWS(CellAddress::from_a1("A0"), FormatError); });
    RUN("no column letters throws",         { EXPECT_THROWS(CellAddress::from_a1("123"), FormatError); });
    RUN("no row digits throws",             { EXPECT_THROWS(CellAddress::from_a1("A"), FormatError); });
    RUN("col > 16384 throws",               { EXPECT_THROWS(CellAddress::from_a1("XFE1"), FormatError); });
    RUN("row > 1048576 throws",             { EXPECT_THROWS(CellAddress::from_a1("A1048577"), FormatError); });
    RUN("mixed case column", {
        auto a = CellAddress::from_a1("b5");
        EXPECT_EQ(a.col,1u); EXPECT_EQ(a.row,4u);
    });
}

// ============================================================
//  SECTION 2: CellValue
// ============================================================
void test_cell_value() {
    group("CellValue");

    RUN("BlankValue is_blank", { CellValue v=BlankValue{}; EXPECT(is_blank(v)); EXPECT(!is_string(v)); });
    RUN("bool true/false", {
        CellValue t=true, f=false;
        EXPECT(is_bool(t)); EXPECT(get_bool(t)==true);
        EXPECT(get_bool(f)==false);
    });
    RUN("int64", { CellValue v=int64_t{42}; EXPECT(is_int(v)); EXPECT_EQ(get_int(v),42); EXPECT(is_numeric(v)); });
    RUN("double", { CellValue v=3.14; EXPECT(is_double(v)); EXPECT_NEAR(get_double(v),3.14,1e-9); });
    RUN("string", { CellValue v=std::string{"hello"}; EXPECT(is_string(v)); EXPECT_EQ(get_string(v),"hello"); });
    RUN("ErrorValue", { CellValue v=ErrorValue{"#REF!"}; EXPECT(is_error(v)); EXPECT_EQ(get_error(v).code,"#REF!"); });
    RUN("as_double from int", { CellValue v=int64_t{10}; EXPECT_EQ(as_double(v),10.0); });
    RUN("as_double from double", { CellValue v=2.5; EXPECT_EQ(as_double(v),2.5); });
    RUN("as_double from string throws", { EXPECT_THROWS(as_double(CellValue{std::string{"x"}}), FormatError); });
    RUN("wrong type get throws", { EXPECT_THROWS(get_int(CellValue{std::string{"x"}}), std::bad_variant_access); });
    RUN("to_string blank", { EXPECT_EQ(to_string(BlankValue{}),""); });
    RUN("to_string int", { EXPECT_EQ(to_string(int64_t{99}),"99"); });
    RUN("to_string bool true", { EXPECT_EQ(to_string(true),"TRUE"); });
    RUN("to_string error", { EXPECT_EQ(to_string(ErrorValue{"#NA"}),"#NA"); });
}

// ============================================================
//  SECTION 3: XLSX create / read roundtrip
// ============================================================
void test_xlsx_roundtrip() {
    group("XLSX create/roundtrip");

    RUN("create empty workbook", {
        auto wb = WorkbookFactory::create(FileFormat::XLSX);
        EXPECT_EQ(wb->sheet_count(), 0u);
        EXPECT(wb->format() == FileFormat::XLSX);
        EXPECT_EQ(wb->format_name(), "XLSX");
    });
    RUN("add/remove/count sheets", {
        auto wb = WorkbookFactory::create();
        wb->add_sheet("Alpha"); wb->add_sheet("Beta"); wb->add_sheet("Gamma");
        EXPECT_EQ(wb->sheet_count(), 3u);
        wb->remove_sheet(1);
        EXPECT_EQ(wb->sheet_count(), 2u);
        EXPECT_EQ(wb->sheet_names()[1], "Gamma");
    });
    RUN("duplicate sheet name throws", {
        auto wb = WorkbookFactory::create();
        wb->add_sheet("S");
        EXPECT_THROWS(wb->add_sheet("S"), WriteError);
    });
    RUN("sheet access by name", {
        auto wb = WorkbookFactory::create(); wb->add_sheet("Report");
        wb->sheet("Report").set_cell("A1", std::string{"ok"});
        EXPECT_EQ(get_string(wb->sheet("Report").cell("A1").value), "ok");
    });
    RUN("sheet access OOB throws", {
        auto wb = WorkbookFactory::create();
        EXPECT_THROWS(wb->sheet(0), RangeError);
    });
    RUN("sheet by name not found throws", {
        auto wb = WorkbookFactory::create();
        EXPECT_THROWS(wb->sheet("Missing"), RangeError);
    });
    RUN("rename sheet", {
        auto wb = WorkbookFactory::create(); wb->add_sheet("Old");
        wb->rename_sheet(0,"New");
        EXPECT_EQ(wb->sheet(0).name(),"New");
    });
    RUN("all cell types survive serialization", {
        auto wb = WorkbookFactory::create();
        auto& sh = wb->add_sheet("Types");
        sh.set_cell("A1", std::string{"hello"});
        sh.set_cell("B1", int64_t{12345});
        sh.set_cell("C1", 3.14159);
        sh.set_cell("D1", true);
        sh.set_cell("E1", false);
        sh.set_cell("F1", ErrorValue{"#DIV/0!"});
        sh.set_formula("G1", "SUM(B1:C1)");

        auto bytes = wb->to_bytes(FileFormat::XLSX, {});
        EXPECT(!bytes.empty());
        auto wb2 = WorkbookFactory::open(bytes, {});
        auto& sh2 = wb2->sheet(0);

        EXPECT_EQ(get_string(sh2.cell("A1").value), "hello");
        EXPECT_EQ(get_int(sh2.cell("B1").value), 12345);
        EXPECT_NEAR(get_double(sh2.cell("C1").value), 3.14159, 1e-5);
        EXPECT_EQ(get_bool(sh2.cell("D1").value), true);
        EXPECT_EQ(get_bool(sh2.cell("E1").value), false);
        EXPECT(is_error(sh2.cell("F1").value));
        EXPECT(sh2.cell("G1").has_formula());
    });
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
    RUN("save and reload file", {
        auto wb = WorkbookFactory::create();
        auto& sh = wb->add_sheet("Data");
        sh.set_cell("A1", std::string{"saved"});
        sh.set_cell("B2", int64_t{999});
        wb->save(tmp("exc_test_save.xlsx"));
        auto wb2 = WorkbookFactory::open(tmp("exc_test_save.xlsx"));
        EXPECT_EQ(get_string(wb2->sheet(0).cell("A1").value),"saved");
        EXPECT_EQ(get_int(wb2->sheet(0).cell("B2").value),999);
    });
    RUN("large sheet 1000 rows", {
        auto wb = WorkbookFactory::create(); auto& sh = wb->add_sheet("Big");
        for (int i=0;i<1000;++i) sh.set_cell(uint32_t(i),0,int64_t(i));
        auto bytes = wb->to_bytes(FileFormat::XLSX,{});
        auto wb2 = WorkbookFactory::open(bytes,{});
        EXPECT_EQ(get_int(wb2->sheet(0).cell(999,0).value), 999);
    });
    RUN("set_table helper", {
        auto wb = WorkbookFactory::create(); auto& sh = wb->add_sheet("T");
        sh.set_table("B2", {
            {std::string{"Name"}, std::string{"Score"}},
            {std::string{"Alice"}, int64_t{95}},
            {std::string{"Bob"},   int64_t{87}},
        });
        EXPECT_EQ(get_string(sh.cell("B2").value),"Name");
        EXPECT_EQ(get_string(sh.cell("C2").value),"Score");
        EXPECT_EQ(get_string(sh.cell("B3").value),"Alice");
        EXPECT_EQ(get_int(sh.cell("C4").value),87);
    });
    RUN("unicode sheet name and string", {
        auto wb = WorkbookFactory::create();
        auto& sh = wb->add_sheet("売上レポート");
        sh.set_cell("A1", std::string{"テスト文字列"});
        auto bytes = wb->to_bytes(FileFormat::XLSX,{});
        auto wb2 = WorkbookFactory::open(bytes,{});
        EXPECT_EQ(wb2->sheet(0).name(),"売上レポート");
        EXPECT_EQ(get_string(wb2->sheet(0).cell("A1").value),"テスト文字列");
    });
}

// ============================================================
//  SECTION 4: Sheet cell access API
// ============================================================
void test_sheet_api() {
    group("Sheet API");

    RUN("cell() returns blank for empty address", {
        auto wb=WorkbookFactory::create(); auto& sh=wb->add_sheet("S");
        auto c = sh.cell(99,99);
        EXPECT(c.is_blank()); EXPECT_EQ(c.address.row,99u);
    });
    RUN("try_cell returns nullopt for blank", {
        auto wb=WorkbookFactory::create(); auto& sh=wb->add_sheet("S");
        EXPECT(!sh.try_cell(5,5).has_value());
    });
    RUN("try_cell returns value when set", {
        auto wb=WorkbookFactory::create(); auto& sh=wb->add_sheet("S");
        sh.set_cell("C3",int64_t{7});
        auto c=sh.try_cell(2,2);
        EXPECT(c.has_value()); EXPECT_EQ(get_int(c->value),7);
    });
    RUN("row() returns sorted cells", {
        auto wb=WorkbookFactory::create(); auto& sh=wb->add_sheet("S");
        sh.set_cell(0,2,int64_t{3}); sh.set_cell(0,0,int64_t{1}); sh.set_cell(0,1,int64_t{2});
        auto r = sh.row(0);
        EXPECT_EQ(r.size(),3u);
        EXPECT_EQ(get_int(r[0].value),1);
        EXPECT_EQ(get_int(r[1].value),2);
        EXPECT_EQ(get_int(r[2].value),3);
    });
    RUN("cells() returns all non-blank", {
        auto wb=WorkbookFactory::create(); auto& sh=wb->add_sheet("S");
        sh.set_cell("A1",int64_t{1}); sh.set_cell("B2",int64_t{2}); sh.set_cell("C3",int64_t{3});
        EXPECT_EQ(sh.cells().size(),3u);
    });
    RUN("for_each_cell sums values", {
        auto wb=WorkbookFactory::create(); auto& sh=wb->add_sheet("S");
        sh.set_cell("A1",int64_t{10}); sh.set_cell("B1",int64_t{20}); sh.set_cell("C1",int64_t{30});
        int64_t sum=0;
        sh.for_each_cell([&](const Cell& c){ if(is_int(c.value)) sum+=get_int(c.value); });
        EXPECT_EQ(sum,60);
    });
    RUN("row_count and col_count update on write", {
        auto wb=WorkbookFactory::create(); auto& sh=wb->add_sheet("S");
        sh.set_cell(4,7,int64_t{1});
        EXPECT_EQ(sh.row_count(),5u);
        EXPECT_EQ(sh.col_count(),8u);
    });
    RUN("formula cell stores formula string", {
        auto wb=WorkbookFactory::create(); auto& sh=wb->add_sheet("S");
        sh.set_formula("A1","SUM(B1:Z1)");
        auto c=sh.cell("A1");
        EXPECT(c.has_formula()); EXPECT_EQ(*c.formula,"SUM(B1:Z1)");
        EXPECT(c.type == CellType::Formula);
    });
}

// ============================================================
//  SECTION 5: Format detection
// ============================================================
void test_format_detection() {
    group("Format detection");

    RUN("garbage bytes throw FormatError", {
        std::vector<uint8_t> g{0,1,2,3,4,5,6,7};
        EXPECT_THROWS(WorkbookFactory::open(g,{}), FormatError);
    });
    RUN("empty buffer throws ParseError", {
        std::vector<uint8_t> e;
        EXPECT_THROWS(WorkbookFactory::open(e,{}), ParseError);
    });
    RUN("4-byte buffer throws ParseError", {
        std::vector<uint8_t> small{0x50,0x4B,0x03,0x04};
        EXPECT_THROWS(WorkbookFactory::open(small,{}), ParseError);
    });
    RUN("created XLSX detected as XLSX", {
        auto wb=WorkbookFactory::create(FileFormat::XLSX);
        wb->add_sheet("S");
        auto bytes=wb->to_bytes(FileFormat::XLSX,{});
        auto wb2=WorkbookFactory::open(bytes,{});
        EXPECT(wb2->format() == FileFormat::XLSX);
    });
}

// ============================================================
//  SECTION 6: XLS internals
// ============================================================
void test_xls_internals() {
    group("XLS internals");

    RUN("decode_rk integer", {
        uint32_t rk=(100<<2)|0x02;
        EXPECT_EQ(xls::decode_rk(rk),100.0);
    });
    RUN("decode_rk integer /100", {
        uint32_t rk=(1000<<2)|0x03;
        EXPECT_NEAR(xls::decode_rk(rk),10.0,1e-9);
    });
    RUN("decode_biff8 compressed ASCII", {
        std::vector<uint8_t> d={3,0,0x00,'A','B','C'};
        size_t consumed=0;
        auto s=xls::decode_biff8_string(d.data(),d.size(),consumed);
        EXPECT_EQ(s,"ABC"); EXPECT_EQ(consumed,6u);
    });
    RUN("decode_biff8 UTF-16LE", {
        std::vector<uint8_t> d={2,0,0x01,'H',0,'i',0};
        size_t consumed=0;
        auto s=xls::decode_biff8_string(d.data(),d.size(),consumed);
        EXPECT_EQ(s,"Hi"); EXPECT_EQ(consumed,7u);
    });
    RUN("decode_biff8 truncated throws", {
        std::vector<uint8_t> d={5,0};
        size_t consumed=0;
        EXPECT_THROWS(xls::decode_biff8_string(d.data(),d.size(),consumed), ParseError);
    });
    RUN("SharedStringTable out-of-range throws", {
        xls::SharedStringTable sst;
        EXPECT_THROWS(sst.get(0), RangeError);
    });
    RUN("FormatRegistry built-in date (14)", {
        xls::FormatRegistry r; EXPECT(r.get(14).is_date_format);
    });
    RUN("FormatRegistry General is not date", {
        xls::FormatRegistry r; EXPECT(!r.get(0).is_date_format);
    });
    RUN("XFTable out-of-range throws", {
        xls::XFTable t; EXPECT_THROWS(t.get(0), RangeError);
    });
    RUN("BiffRecord u16 bounds-checks", {
        xls::BiffRecord r; r.data={0x01,0x02};
        EXPECT_EQ(r.u16(0), 0x0201u);
        EXPECT_THROWS(r.u16(1), ParseError);
    });
}

// ============================================================
//  SECTION 7: PrintArea
// ============================================================
void test_print_area() {
    group("PrintArea");

    RUN("from_range A1:Z50", {
        auto pa=PrintArea::from_range("A1:Z50");
        EXPECT_EQ(pa.first_row,0u); EXPECT_EQ(pa.first_col,0u);
        EXPECT_EQ(pa.last_row,49u); EXPECT_EQ(pa.last_col,25u);
        EXPECT(pa.is_valid());
    });
    RUN("to_range roundtrip", {
        EXPECT_EQ(PrintArea::from_range("B2:D10").to_range(),"B2:D10");
    });
    RUN("single cell", {
        auto pa=PrintArea::from_range("C5");
        EXPECT_EQ(pa.first_row,pa.last_row); EXPECT_EQ(pa.first_col,pa.last_col);
    });
    RUN("reversed coords normalize", {
        auto pa=PrintArea::from_range("Z50:A1");
        EXPECT_EQ(pa.first_row,0u); EXPECT_EQ(pa.last_row,49u);
    });
    RUN("invalid range throws", {
        EXPECT_THROWS(PrintArea::from_range("A0:B5"), FormatError);
    });
}

// ============================================================
//  SECTION 8: PageSetup defaults & apply
// ============================================================
void test_page_setup() {
    group("PageSetup");

    RUN("defaults", {
        PageSetup ps;
        EXPECT(ps.paper_size==PaperSize::A4);
        EXPECT(ps.orientation==Orientation::Portrait);
        EXPECT(ps.fit_to==FitTo::None);
        EXPECT_EQ(ps.scale_percent,100);
        EXPECT(!ps.print_gridlines);
        EXPECT(!ps.print_area.has_value());
    });
    RUN("margins defaults", {
        PageSetup ps;
        EXPECT_NEAR(ps.margins.left,  0.70, 1e-9);
        EXPECT_NEAR(ps.margins.right, 0.70, 1e-9);
        EXPECT_NEAR(ps.margins.top,   0.75, 1e-9);
        EXPECT_NEAR(ps.margins.bottom,0.75, 1e-9);
    });

    // -- apply_page_setup tests (modify XLSX in-place) --
    auto make_xlsx=[](const std::string& path){
        auto wb=WorkbookFactory::create();
        auto& sh=wb->add_sheet("Sheet1");
        sh.set_cell("A1",std::string{"data"}); sh.set_cell("B2",int64_t{42});
        wb->save(path);
    };
    auto raw_contains=[](const std::string& path, const std::string& needle)->bool{
        std::ifstream f(path,std::ios::binary);
        std::string s((std::istreambuf_iterator<char>(f)),{});
        return s.find(needle)!=std::string::npos;
    };

    RUN("apply portrait A4", {
        make_xlsx(tmp("exc_ps1.xlsx"));
        PageSetup ps; ps.paper_size=PaperSize::A4; ps.orientation=Orientation::Portrait;
        ExcelPrinter p; p.apply_page_setup(tmp("exc_ps1.xlsx"),"Sheet1",ps);
        EXPECT(raw_contains(tmp("exc_ps1.xlsx"),"paperSize=\"9\""));
        EXPECT(raw_contains(tmp("exc_ps1.xlsx"),"orientation=\"portrait\""));
    });
    RUN("apply landscape", {
        make_xlsx(tmp("exc_ps2.xlsx"));
        PageSetup ps; ps.orientation=Orientation::Landscape;
        ExcelPrinter p; p.apply_page_setup(tmp("exc_ps2.xlsx"),"Sheet1",ps);
        EXPECT(raw_contains(tmp("exc_ps2.xlsx"),"orientation=\"landscape\""));
    });
    RUN("apply FitTo::Width", {
        make_xlsx(tmp("exc_ps3.xlsx"));
        PageSetup ps; ps.fit_to=FitTo::Width; ps.fit_to_pages_wide=1;
        ExcelPrinter p; p.apply_page_setup(tmp("exc_ps3.xlsx"),"Sheet1",ps);
        EXPECT(raw_contains(tmp("exc_ps3.xlsx"),"fitToPage=\"1\""));
        EXPECT(raw_contains(tmp("exc_ps3.xlsx"),"fitToWidth=\"1\""));
    });
    RUN("apply FitTo::WidthAndHeight", {
        make_xlsx(tmp("exc_ps4.xlsx"));
        PageSetup ps; ps.fit_to=FitTo::WidthAndHeight;
        ps.fit_to_pages_wide=1; ps.fit_to_pages_tall=1;
        ExcelPrinter p; p.apply_page_setup(tmp("exc_ps4.xlsx"),"Sheet1",ps);
        EXPECT(raw_contains(tmp("exc_ps4.xlsx"),"fitToWidth=\"1\""));
        EXPECT(raw_contains(tmp("exc_ps4.xlsx"),"fitToHeight=\"1\""));
    });
    RUN("apply print_area", {
        make_xlsx(tmp("exc_ps5.xlsx"));
        PageSetup ps; ps.print_area=PrintArea::from_range("A1:B20");
        ExcelPrinter p; p.apply_page_setup(tmp("exc_ps5.xlsx"),"Sheet1",ps);
        EXPECT(raw_contains(tmp("exc_ps5.xlsx"),"Print_Area"));
        EXPECT(raw_contains(tmp("exc_ps5.xlsx"),"$A$1"));
    });
    RUN("apply repeat_titles rows", {
        make_xlsx(tmp("exc_ps6.xlsx"));
        PageSetup ps; RepeatTitles rt; rt.row_start=0; rt.row_end=0; ps.repeat_titles=rt;
        ExcelPrinter p; p.apply_page_setup(tmp("exc_ps6.xlsx"),"Sheet1",ps);
        EXPECT(raw_contains(tmp("exc_ps6.xlsx"),"Print_Titles"));
    });
    RUN("apply gridlines + black_and_white", {
        make_xlsx(tmp("exc_ps7.xlsx"));
        PageSetup ps; ps.print_gridlines=true; ps.black_and_white=true;
        ExcelPrinter p; p.apply_page_setup(tmp("exc_ps7.xlsx"),"Sheet1",ps);
        EXPECT(raw_contains(tmp("exc_ps7.xlsx"),"gridLines=\"1\""));
        EXPECT(raw_contains(tmp("exc_ps7.xlsx"),"blackAndWhite=\"1\""));
    });
    RUN("apply header/footer", {
        make_xlsx(tmp("exc_ps8.xlsx"));
        PageSetup ps;
        ps.header_footer.odd_header="&C&14Report";
        ps.header_footer.odd_footer="&LPage &P / &N";
        ExcelPrinter p; p.apply_page_setup(tmp("exc_ps8.xlsx"),"Sheet1",ps);
        EXPECT(raw_contains(tmp("exc_ps8.xlsx"),"Report"));
        EXPECT(raw_contains(tmp("exc_ps8.xlsx"),"oddFooter"));
    });
    RUN("apply empty sheet_name defaults to first sheet", {
        make_xlsx(tmp("exc_ps9.xlsx"));
        PageSetup ps; ps.orientation=Orientation::Landscape;
        ExcelPrinter p; p.apply_page_setup(tmp("exc_ps9.xlsx"),"",ps);
        EXPECT(raw_contains(tmp("exc_ps9.xlsx"),"orientation=\"landscape\""));
    });
    RUN("apply re-opens correctly after modification", {
        make_xlsx(tmp("exc_ps10.xlsx"));
        PageSetup ps; ps.print_gridlines=true;
        ExcelPrinter p; p.apply_page_setup(tmp("exc_ps10.xlsx"),"Sheet1",ps);
        // Must still be a valid XLSX after modification
        auto wb=WorkbookFactory::open(tmp("exc_ps10.xlsx"));
        EXPECT_EQ(get_string(wb->sheet(0).cell("A1").value),"data");
        EXPECT_EQ(get_int(wb->sheet(0).cell("B2").value),42);
    });
    RUN("combined all settings", {
        make_xlsx(tmp("exc_ps_full.xlsx"));
        PageSetup ps;
        ps.paper_size=PaperSize::A4; ps.orientation=Orientation::Landscape;
        ps.fit_to=FitTo::Width; ps.fit_to_pages_wide=1;
        ps.print_gridlines=true; ps.print_area=PrintArea::from_range("A1:H20");
        ps.margins.left=0.5; ps.margins.right=0.5;
        ps.header_footer.odd_header="&C売上レポート";
        RepeatTitles rt; rt.row_start=0; rt.row_end=0; ps.repeat_titles=rt;
        ExcelPrinter p; p.apply_page_setup(tmp("exc_ps_full.xlsx"),"Sheet1",ps);
        EXPECT(raw_contains(tmp("exc_ps_full.xlsx"),"paperSize=\"9\""));
        EXPECT(raw_contains(tmp("exc_ps_full.xlsx"),"orientation=\"landscape\""));
        EXPECT(raw_contains(tmp("exc_ps_full.xlsx"),"Print_Area"));
        EXPECT(raw_contains(tmp("exc_ps_full.xlsx"),"Print_Titles"));
        EXPECT(raw_contains(tmp("exc_ps_full.xlsx"),"売上レポート"));
    });
}

// ============================================================
//  SECTION 9: ExcelPrinter engine / validation
// ============================================================
void test_printer() {
    group("ExcelPrinter");

    RUN("ExcelPrinter constructs without crash", {
        ExcelPrinter p;
        // Construction succeeds — Excel COM initialized
    });
    RUN("list_printers does not crash", {
        ExcelPrinter p;
        auto pr=p.list_printers();
        std::cout<<"        "<<pr.size()<<" printer(s) found\n";
        EXPECT(true);
    });
    RUN("to_pdf with empty path throws PrintError", {
        ExcelPrinter p;
        PdfOptions o; // output_path=""
        EXPECT_THROWS(p.to_pdf(tmp("exc_ps_full.xlsx"),o), PrintError);
    });
    RUN("to_pdf requires output_path", {
        ExcelPrinter p;
        PdfOptions opts; // output_path=""
        EXPECT_THROWS(p.to_pdf(tmp("exc_ps_full.xlsx"), opts), PrintError);
    });
}

// ============================================================
//  SECTION 10: BatchPrinter
// ============================================================
void test_batch_printer() {
    group("BatchPrinter");

    auto make = [](const std::string& path, const std::string& sheet, int val) {
        auto wb = WorkbookFactory::create();
        wb->add_sheet(sheet).set_cell("A1", int64_t{val});
        wb->save(path);
    };
    make(tmp("batch1.xlsx"), "Sheet1", 1);
    make(tmp("batch2.xlsx"), "Report", 2);
    make(tmp("batch3.xlsx"), "Data",   3);

    RUN("add jobs via fluent API", {
        BatchPrinter bp;
        bp.add(tmp("batch1.xlsx"), "Sheet1")
          .add(tmp("batch2.xlsx"), "Report")
          .add(tmp("batch3.xlsx"));
        EXPECT_EQ(bp.job_count(), 3u);
    });
    RUN("add_all from vector", {
        BatchPrinter bp;
        std::vector<PrintJob> jobs = {
            PrintJob::make(tmp("batch1.xlsx"), "Sheet1"),
            PrintJob::make(tmp("batch2.xlsx"), "Report"),
        };
        bp.add_all(jobs);
        EXPECT_EQ(bp.job_count(), 2u);
    });
    RUN("clear resets job list", {
        BatchPrinter bp;
        bp.add(tmp("batch1.xlsx")).add(tmp("batch2.xlsx"));
        EXPECT_EQ(bp.job_count(), 2u);
        bp.clear();
        EXPECT_EQ(bp.job_count(), 0u);
    });
    RUN("PrintJob::make with all fields", {
        PageSetup ps; ps.orientation = Orientation::Landscape;
        auto job = PrintJob::make(tmp("batch1.xlsx"), "Sheet1", ps, "Q1_Sales");
        EXPECT_EQ(job.file_path, tmp("batch1.xlsx"));
        EXPECT_EQ(job.sheet_name, "Sheet1");
        EXPECT_EQ(job.label, "Q1_Sales");
        EXPECT(job.setup.has_value());
        EXPECT(job.setup->orientation == Orientation::Landscape);
    });
    RUN("PrintJob without setup", {
        auto job = PrintJob::make(tmp("batch1.xlsx"));
        EXPECT(!job.setup.has_value());
        EXPECT(job.sheet_name.empty());
    });
    RUN("BatchPrinter constructs without crash", {
        BatchPrinter bp;
        EXPECT_EQ(bp.job_count(), 0u);
    });
    RUN("BatchResult failure tracking", {
        BatchResult r;
        r.total = 3;
        r.jobs.push_back({0, "a.xlsx", "", true,  ""});
        r.jobs.push_back({1, "b.xlsx", "", false, "File not found"});
        r.jobs.push_back({2, "c.xlsx", "", true,  ""});
        r.succeeded = 2; r.failed = 1;
        EXPECT(!r.all_success());
        auto fails = r.failures();
        EXPECT_EQ(fails.size(), 1u);
        EXPECT_EQ(fails[0]->file_path, "b.xlsx");
    });
    RUN("job fluent add sets count", {
        BatchPrinter bp;
        bp.add(tmp("batch1.xlsx"),"Sheet1")
          .add(tmp("batch2.xlsx"),"Report")
          .add(tmp("batch3.xlsx"),"Data");
        EXPECT_EQ(bp.job_count(), 3u);
    });
    RUN("job_count after add_all", {
        BatchPrinter bp;
        bp.add(tmp("batch1.xlsx"))
          .add(tmp("batch2.xlsx"))
          .add(tmp("batch3.xlsx"));
        EXPECT_EQ(bp.job_count(), 3u);
    });
}

// ============================================================
//  SECTION 11: DEFLATE / CRC32
// ============================================================
static std::vector<uint8_t> from_hex(const char* h) {
    std::vector<uint8_t> v;
    for (; h[0] && h[1]; h += 2) {
        auto hex = [](char c) -> uint8_t {
            if (c >= '0' && c <= '9') return uint8_t(c - '0');
            if (c >= 'a' && c <= 'f') return uint8_t(c - 'a' + 10);
            return uint8_t(c - 'A' + 10);
        };
        v.push_back(uint8_t((hex(h[0]) << 4) | hex(h[1])));
    }
    return v;
}

void test_deflate() {
    group("DEFLATE / CRC32");

    // ── CRC32 ─────────────────────────────────────────────
    RUN("CRC32 Hello World", {
        auto s = reinterpret_cast<const uint8_t*>("Hello, World!");
        EXPECT_EQ(excellib::detail::crc32_compute(s, 13), 0xEC4AC3D0u);
    });
    RUN("CRC32 excellib test", {
        auto s = reinterpret_cast<const uint8_t*>("excellib test");
        EXPECT_EQ(excellib::detail::crc32_compute(s, 13), 0x9EED09FFu);
    });
    RUN("CRC32 空データ = 0x00000000", {
        EXPECT_EQ(excellib::detail::crc32_compute(nullptr, 0), 0x00000000u);
    });

    // ── 非圧縮ブロック (btype=0) ──────────────────────────
    RUN("stored block: Hello", {
        // BFINAL=1 BTYPE=00 + align + LEN=5 NLEN=~5 + "Hello"
        auto comp = from_hex("010500FAFF48656C6C6F");
        auto out  = excellib::detail::deflate_decompress(comp.data(), comp.size(), 5);
        EXPECT(out == std::vector<uint8_t>({'H','e','l','l','o'}));
    });

    // ── 固定ハフマン (btype=1) ────────────────────────────
    // Python: zlib.compressobj(6, DEFLATED, -15).compress(b"Hello, DEFLATE!") + flush()
    RUN("fixed huffman: Hello, DEFLATE!", {
        auto comp = from_hex("f348cdc9c9d751707175f3710c71550400");
        auto out  = excellib::detail::deflate_decompress(comp.data(), comp.size(), 15);
        std::string s(out.begin(), out.end());
        EXPECT_EQ(s, "Hello, DEFLATE!");
    });
    // 反復データ (後方参照を含む): b"A"*22
    RUN("fixed huffman: 22xA (back-reference)", {
        auto comp = from_hex("7374c40600");
        auto out  = excellib::detail::deflate_decompress(comp.data(), comp.size(), 22);
        EXPECT_EQ(out.size(), 22u);
        EXPECT(std::all_of(out.begin(), out.end(), [](uint8_t b){ return b == 'A'; }));
    });
    // アルファベット×3 (反復パターン)
    RUN("fixed huffman: alphabet x3", {
        auto comp = from_hex(
            "4b4c4a4e494d4bcfc8cccacec9cdcb2f282c2a2e292d2b"
            "afa8ac4a24430600");
        auto out  = excellib::detail::deflate_decompress(comp.data(), comp.size(), 78);
        EXPECT_EQ(out.size(), 78u);
        std::string s(out.begin(), out.end());
        std::string expected;
        for (int i = 0; i < 3; ++i) expected += "abcdefghijklmnopqrstuvwxyz";
        EXPECT_EQ(s, expected);
    });
    // pangram (多種文字)
    RUN("fixed huffman: pangram", {
        auto comp = from_hex(
            "0bc94855282ccd4cce56482aca2fcf5348cbaf50c82acd2d"
            "2856c82f4b2d5228014ae72456552aa4e4a70300");
        auto out  = excellib::detail::deflate_decompress(comp.data(), comp.size(), 43);
        std::string s(out.begin(), out.end());
        EXPECT_EQ(s, "The quick brown fox jumps over the lazy dog");
    });

    // ── 動的ハフマン (btype=2) ────────────────────────────
    // Python: zlib.compressobj(6, DEFLATED, -15).compress(xml2_bytes) + flush()
    RUN("dynamic huffman: XLSX-like XML", {
        auto comp = from_hex(
            "4d8ed10ac2300c457fa5e4dda50e1191b64311bf403fa074"
            "711baeed68cba67f6f3765f8126e4e38e18aea657b365288"
            "9d7712b6050746cef8ba738d84fbedba39008b49bb5af7de"
            "91843745a894987c78c69628b1ecbb28a14d69382246d392"
            "d5b1f003b97c79f86075ca6b68300e8174bd48b6c792f33d"
            "5add39506261179db412c14f2ce41e999a399ce634aa5d29"
            "705402cd0f9fbf78cbf9ca31ab79fefdc2b5a4fa00");
        const std::string expected =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
            "<sheetData><row r=\"1\"><c r=\"A1\"><v>42</v></c>"
            "<c r=\"B1\"><v>100</v></c></row></sheetData></worksheet>";
        auto out = excellib::detail::deflate_decompress(comp.data(), comp.size(), expected.size());
        std::string s(out.begin(), out.end());
        EXPECT_EQ(s, expected);
    });

    // ── エラーケース ──────────────────────────────────────
    RUN("空データで例外", {
        std::vector<uint8_t> empty;
        bool threw = false;
        try { excellib::detail::deflate_decompress(empty.data(), 0); }
        catch (std::exception&) { threw = true; }
        EXPECT(threw);
    });
    RUN("btype=3 (予約済み) で例外", {
        // BFINAL=1 BTYPE=11 → 0b00000111 = 0x07
        std::vector<uint8_t> bad = {0x07};
        bool threw = false;
        try { excellib::detail::deflate_decompress(bad.data(), bad.size()); }
        catch (std::exception&) { threw = true; }
        EXPECT(threw);
    });
}

// ============================================================
//  SECTION 12: 結合セル / set_row / col / 警告コールバック
// ============================================================
void test_new_features() {
    group("New Features v3.0.0");

    RUN("merge_cells_roundtrip", {
        auto wb = WorkbookFactory::create();
        auto& sh = wb->add_sheet("S");
        sh.set_cell("A1", std::string{"merged"});
        sh.merge("A1:C3");
        EXPECT(sh.merged_ranges().size() == 1);
        auto bytes = wb->to_bytes(FileFormat::XLSX, {});
        auto wb2 = WorkbookFactory::open(bytes);
        EXPECT(wb2->sheet(0).merged_ranges().size() == 1);
        auto r = wb2->sheet(0).merged_ranges()[0];
        EXPECT_EQ(r.row1, uint32_t(0)); EXPECT_EQ(r.col1, uint32_t(0));
        EXPECT_EQ(r.row2, uint32_t(2)); EXPECT_EQ(r.col2, uint32_t(2));
    });

    RUN("merge_unmerge", {
        auto wb = WorkbookFactory::create();
        auto& sh = wb->add_sheet("S");
        sh.merge("B2:D4");
        EXPECT_EQ(sh.merged_ranges().size(), 1u);
        auto rng = sh.merged_ranges()[0];
        EXPECT_EQ(rng.row1, 1u); EXPECT_EQ(rng.col1, 1u);
        EXPECT_EQ(rng.row2, 3u); EXPECT_EQ(rng.col2, 3u);
        sh.unmerge(rng);
        EXPECT_EQ(sh.merged_ranges().size(), 0u);
    });

    RUN("CellRange::from_a1 and to_a1", {
        auto r = CellRange::from_a1("A1:C3");
        EXPECT_EQ(r.row1, 0u); EXPECT_EQ(r.col1, 0u);
        EXPECT_EQ(r.row2, 2u); EXPECT_EQ(r.col2, 2u);
        EXPECT_EQ(r.to_a1(), "A1:C3");
    });

    RUN("set_cell_overloads", {
        auto wb = WorkbookFactory::create();
        auto& sh = wb->add_sheet("S");
        sh.set_cell("A1", "hello");
        sh.set_cell("B1", 42);
        sh.set_cell("C1", 3.14);
        EXPECT(is_string(sh.cell("A1").value));
        EXPECT(is_int(sh.cell("B1").value));
        EXPECT(is_double(sh.cell("C1").value));
        EXPECT_EQ(get_string(sh.cell("A1").value), std::string("hello"));
        EXPECT_EQ(get_int(sh.cell("B1").value), int64_t(42));
    });

    RUN("set_row_and_col", {
        auto wb = WorkbookFactory::create();
        auto& sh = wb->add_sheet("S");
        sh.set_row(0, {std::string("a"), int64_t(1), 2.0});
        EXPECT_EQ(get_string(sh.cell("A1").value), std::string("a"));
        EXPECT_EQ(get_int(sh.cell("B1").value), int64_t(1));
        auto c = sh.col(0);
        EXPECT(c.size() >= 1);
        EXPECT_EQ(get_string(c[0].value), std::string("a"));
    });

    RUN("warning_callback", {
        std::vector<ParseWarning> warns;
        OpenOptions opts;
        opts.on_warning = [&](const ParseWarning& w) { warns.push_back(w); };
        auto wb = WorkbookFactory::create();
        EXPECT(warns.empty());
    });

    RUN("xls_merge_write_throws", {
        auto wb = WorkbookFactory::create();
        auto& sh = wb->add_sheet("S");
        sh.set_cell("A1", std::string{"data"});
        auto bytes = wb->to_bytes(FileFormat::XLSX, {});
        auto wb2 = WorkbookFactory::open(bytes);
        auto bytes2 = wb2->to_bytes(FileFormat::XLSX, {});
        EXPECT(!bytes2.empty());
    });

    RUN("XLS write throws WriteError", {
        auto wb = WorkbookFactory::create();
        wb->add_sheet("S");
        auto bytes = wb->to_bytes(FileFormat::XLSX, {});
        auto xls_wb = WorkbookFactory::open(bytes);
        EXPECT_THROWS(xls_wb->to_bytes(FileFormat::XLS, {}), WriteError);
    });
}

// ============================================================
//  Main
// ============================================================
int main() {
#ifdef _WIN32
    SetConsoleOutputCP(CP_UTF8);
#endif
    std::cout<<"==========================================\n";
    std::cout<<"    excellib v2 - Full Test Suite         \n";
    std::cout<<"==========================================\n";

    test_cell_address();
    test_cell_value();
    test_xlsx_roundtrip();
    test_sheet_api();
    test_format_detection();
    test_xls_internals();
    test_print_area();
    test_page_setup();
    test_printer();
    test_batch_printer();
    test_deflate();
    test_new_features();

    std::cout<<"\n==========================================\n";
    std::cout<<"  Results: "<<g_pass<<" passed  "<<g_fail<<" failed  "<<g_skip<<" skipped\n";
    std::cout<<"==========================================\n";
    return g_fail==0 ? 0 : 1;
}
