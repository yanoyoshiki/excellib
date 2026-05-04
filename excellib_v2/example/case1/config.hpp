#pragma once
/**
 * case1 — 設定ファイル
 *
 * 環境に合わせて以下の定数を編集してください。
 * 列構成・パステンプレートが変わっても、このファイルだけ書き換えれば対応できます。
 */
#include <string_view>
#include <cstdint>

namespace case1::config {

// ============================================================
//  パステンプレート
// ============================================================
//   {INPUT} を GUI 入力で置換します。
//   拡張子はつけずに、xlsx → xls の順で自動試行します。
//
//   例:
//     R"(C:\reports\{INPUT}\master)"
//   → ユーザーが "2026Q1" と入力したら:
//     "C:\reports\2026Q1\master.xlsx" を試し、なければ ".xls" を試す
//
constexpr std::string_view PATH_TEMPLATE = R"(C:\path\to\{INPUT}\master)";

// ============================================================
//  マスターシート設定
// ============================================================
constexpr std::string_view MASTER_SHEET_NAME = "";    // 空文字列 = 先頭シート
constexpr uint32_t MASTER_START_ROW = 1;              // 0-based (1 = 2行目から)

// ============================================================
//  マスターの列構成 (0-based)
// ============================================================
//   列が増減した場合はここを編集
constexpr uint32_t COL_FILE_PATH  = 0;   // A列: 印刷対象 Excel のパス
constexpr uint32_t COL_SHEET_NAME = 1;   // B列: 印刷対象シート名
constexpr uint32_t COL_PRINT_COLS = 2;   // C列: 印刷列範囲 "A:D" など (空なら全列)
constexpr uint32_t COL_LABEL      = 3;   // D列: ラベル (PDF 名・進捗表示)

// 空白判定に使う列（この列が空白になったら読み取り終了）
constexpr uint32_t COL_KEY = COL_FILE_PATH;

// ============================================================
//  デフォルトのページ設定
// ============================================================
//   ジョブ個別に GUI で上書き可能
constexpr bool      DEFAULT_FIT_TO_WIDTH = true;
constexpr uint16_t  DEFAULT_FIT_PAGES_WIDE = 1;
constexpr bool      DEFAULT_LANDSCAPE = false;

// ============================================================
//  GUI 表記
// ============================================================
constexpr std::string_view APP_TITLE = "case1 — エクセル一括印刷";
constexpr std::string_view INPUT_LABEL = "可変文字列";
constexpr std::string_view INPUT_HINT  = "ここにマスター Excel パスの可変部分を入力";

}  // namespace case1::config
