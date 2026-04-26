# コントリビューティングガイド

## ブランチ戦略

```
main        本番リリース済みコード
develop     開発中の統合ブランチ
feature/*   新機能
fix/*       バグ修正
docs/*      ドキュメントのみの変更
```

## 開発の流れ

```bash
# develop から feature ブランチを切る
git checkout develop
git checkout -b feature/your-feature

# 実装 → テスト → コミット
git add -p          # 変更を確認しながらステージング
git commit -m "feat: add your feature"

# develop にマージ
git checkout develop
git merge --no-ff feature/your-feature
```

## コミットメッセージ規約

[Conventional Commits](https://www.conventionalcommits.org/) に従う。

```
feat:     新機能
fix:      バグ修正
docs:     ドキュメントのみの変更
refactor: 挙動を変えないリファクタリング
test:     テストの追加・修正
chore:    ビルド設定・依存ライブラリの変更
```

例：
```
feat: add Sheet::find_cell() for searching by value
fix: correct MULRK size_t underflow when record < 10 bytes
docs: add PageSetup examples to API_REFERENCE.md
```

## テスト

コードを変更したら必ずテストを実行すること。

```bash
# ビルド
g++ -std=c++17 -Wno-trigraphs -Iinclude \
    src/common/common.cpp src/common/factory.cpp \
    src/xls/xls_parser.cpp src/xlsx/xlsx_parser.cpp \
    src/print/xlsx_page_setup.cpp src/print/excel_printer.cpp \
    src/print/batch_printer.cpp \
    tests/test_all.cpp -lz -o build/test_all

# 実行（ネットワーク完全遮断で確認する場合）
unshare -n ./build/test_all

# 全テスト通過が必須
```

新機能には対応するテストを `tests/test_all.cpp` に追加すること。

## コードスタイル

`.clang-format` の設定に従う。

```bash
clang-format -i src/**/*.cpp include/**/*.hpp
```

## プルリクエスト前チェックリスト

- [ ] `tests/test_all.cpp` が全て PASS する
- [ ] 新機能には対応するテストを追加した
- [ ] `docs/API_REFERENCE.md` を更新した（API 変更の場合）
- [ ] `docs/CHANGELOG.md` に変更内容を記載した
- [ ] ビルド警告がないこと（`-Wall -Wextra`）
