#pragma once
/**
 * excellib — Unified Excel (xls/xlsx) read/write/print library
 * C++17, zero external runtime dependencies beyond zlib
 *
 * Quick start:
 *   #include "excellib/excellib.hpp"
 *   auto wb = excellib::WorkbookFactory::open("data.xlsx");
 */
#include "excellib/workbook.hpp"
#include "excellib/print_settings.hpp"
#include "excellib/excel_printer.hpp"
#include "excellib/batch_printer.hpp"
