#pragma once
/**
 * excellib — Unified Excel (xls/xlsx) read/write/print library
 * C++17, zero external runtime dependencies beyond zlib
 *
 * Quick start:
 *   #include "excellib/excellib.hpp"
 *   auto wb = excellib::WorkbookFactory::open("data.xlsx");
 */

// workbook.hpp: CellValue / Cell / Sheet / Workbook / WorkbookFactory など
// ファイルの読み書きに必要なすべての型と API を定義している
#include "excellib/workbook.hpp"

// print_settings.hpp: PageSetup / PaperSize / Orientation など
// ページ設定・印刷オプションに必要な型を定義している
#include "excellib/print_settings.hpp"

// excel_printer.hpp: ExcelPrinter クラス
// PDF 出力・プリンター送信・ページ設定の書き込みを担当する
#include "excellib/excel_printer.hpp"

// batch_printer.hpp: BatchPrinter / PrintJob / BatchResult クラス
// 複数ファイルの一括印刷・PDF 化に特化したクラスを定義している
#include "excellib/batch_printer.hpp"
