# CHANGELOG

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
