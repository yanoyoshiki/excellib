# CHANGELOG

## [3.1.0] - 2026-05-03

### Added
- `PrintArea::columns_only(string_view)` — "A:D" 形式で列範囲のみの印刷エリアを作成
- `PrintArea::columns_only(uint32_t, uint32_t)` — 0-based インデックスで列範囲を指定
- `PrintArea::is_columns_only()` — 列のみ印刷エリアかどうかを判定
- `PrintArea::to_range()` が `col_range_only=true` の場合 `"$A:$D"` 形式を返すよう更新
- `Sheet::read_until_blank(start, Direction)` — 空白セルが出るまでセルを取得
- `Direction` 列挙型 — `Direction::Down` / `Direction::Right`
- `BatchPrinter::request_cancel()` — バッチ実行中のキャンセル要求（スレッドセーフ）
- `BatchPrinter::is_cancel_requested()` — キャンセル要求の確認
- `BatchPrinter::reset_cancel()` — キャンセルフラグのリセット
- `BatchPrinter::from_failures(result)` — 失敗/キャンセルジョブのみ再試行する BatchPrinter を生成
- `BatchResult::cancelled` — バッチがキャンセルされたかどうかのフラグ
- `BatchResult::cancelled_count` — キャンセルされたジョブ数
- `JobResult::cancelled` — 個別ジョブがキャンセルされたかどうかのフラグ
- `JobResult::original_job` — 再試行用に元の PrintJob を保持

---

## [3.0.0] - 2026-05-03

### Breaking Changes
- `OpenOptions` から `preserve_formulas`, `preserve_styles`, `fail_on_unsupported` を削除（実装されていなかったため）
- `SaveOptions` から `recalculate` を削除（Excel エンジンなしでは不可能なため）
- XLS 書き込み（`save`/`to_bytes` に `FileFormat::XLS` を渡すこと）は `WriteError` が発生する（以前と同様だが API から明示化）
- XLS の `Cell::formula` が `"<formula>"` の代わりに `std::nullopt` を返すようになった

### Added
- `ParseWarning` 構造体と `WarningCallback` typedef を追加
- `OpenOptions::on_warning` と `OpenOptions::strict` を追加
- `Sheet::merge(range)` / `unmerge(range)` / `merged_ranges()` を追加（XLSX 読み書き・XLS 読み込みのみ）
- `Sheet::set_row(row, values)` を追加
- `Sheet::col(col_idx)` を追加
- `set_cell(addr, const char*)` / `set_cell(addr, int)` / `set_cell(addr, long)` オーバーロードを追加
- XLSX open → save で charts, drawings, images, comments などの未知エントリを保持するようになった（passthrough mode）
- XLSX open → save で元の styles.xml を保持するようになった
- `CellRange` 構造体を追加（`from_a1()` / `to_a1()` 対応）
- VBScript 生成でシート名・ヘッダ・フッタの特殊文字をエスケープするようになった
- Windows 環境の VBScript を CP932 で出力するようになった（日本語文字化け修正）

### Fixed
- XLS MULRK レコードの size_t アンダーフローによるクラッシュを修正
- XLSX パーサの複数箇所で `catch(...)` が警告なく握りつぶしていた問題を修正
- `XlsxWorkbook::to_bytes(FileFormat::XLS, ...)` が WriteError を正しく throw するように修正

---

## v2.0.0

### 新機能
- `BatchPrinter` — 複数ファイル・複数シートを Excel 1 起動で一括処理
  - `print_all()` — プリンターに一括送信
  - `to_pdf_merged()` — 全シートを 1 PDF に結合
  - `to_pdf_individual()` — ジョブごとに個別 PDF
- `Sheet::set_table()` — 2D ベクターを一括書き込み
- `to_string(CellValue)` — 全型を文字列に変換するユーティリティ
- `ExcelPrinter` を Excel COM 専用に整理（LibreOffice 依存を完全排除）

### バグ修正
- **ZIP セントラルディレクトリ `mod_date` 欠落**
  `apply_page_setup` 後に XLSX を再オープンすると「ZIP 破損」エラーになっていた問題を修正
- **BIFF8 DIMENSION レコードの境界チェック**
  `data.size() >= 8` を確認しながら `u16(8)` を読んでいた越境アクセスを修正（`>= 10` に変更）
- **XLS MULRK レコードの size_t アンダーフロー**
  `(data.size() - 4) / 6` の計算でデータが 4 バイト未満の場合にアンダーフローしていた問題を修正
- **XLSX ZIP EOCD 探索ループ**
  `i > 0` の条件で先頭バイトを見落とす可能性があった問題を修正（`ptrdiff_t` に変更）
- **XLSX stoul/stoi による空文字列クラッシュ**
  空の属性値を `std::stoul` に渡してクラッシュする問題を修正
- **シートパス解決**
  `worksheets/sheet1.xml` の前置き `xl/` が欠落するケースを修正（3 パターン対応）
- **`rename_sheet` の const name_ フィールド問題**
  全セルコピー方式で解決

### 変更
- `PrintEngine` enum を削除（ExcelCOM 一択に統一）
- `is_libreoffice_available()` / `find_libreoffice_path()` を削除
- `ExcelPrinter` のコンストラクタ引数を削除（`ExcelPrinter()` のみ）

---

## v1.0.0

- XLS（BIFF8）読み込みパーサー実装
  - OLE2 コンテナ解析
  - SST / XF / RK / MULRK / NUMBER / LABELSST / BOOLERR / FORMULA 対応
- XLSX（OOXML）読み書きパーサー実装
  - 自前 ZIP 読み書き（zlib のみ依存）
  - 自前 XML 解析（外部パーサー不使用）
- `WorkbookFactory` によるフォーマット自動判定（マジックバイト）
- `ExcelPrinter` による印刷・PDF 出力（Excel COM）
- ページ設定の XLSX XML 注入（`apply_page_setup`）
- 92 件のテストスイート
