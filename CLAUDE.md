# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

Always respond in Japanese.
- 常に日本語で返答すること
- 作業中は常に、今何をしているのかをプログラミング初心者にもわかるように丁寧に説明すること
- 専門用語を使う場合は必ず簡単な言葉で補足すること

## プロジェクト概要

**令和8年度（R8）診療報酬改定**の公式PDF資料と、そこから新旧対照表を抽出するPythonスクリプトを格納するリポジトリ。

## スクリプト実行

```bash
python extract_shinkyuu.py
```

- 入力: `総－２別紙１－１医科診療報酬点数表.pdf`（948ページ、中医協公開の原本）
- 出力: `R8年度診療報酬改定_新旧対照表.xlsx`（約3000件の変更ブロック、アンダーライン付き）
- 実行時間: 数分（全948ページを走査）
- 依存ライブラリ: `PyMuPDF (fitz)`, `xlsxwriter`

## extract_shinkyuu.py のアーキテクチャ

PDFの左カラム（改正後）と右カラム（改正前）を比較し、アンダーライン付き変更箇所を抽出してXLSXに出力する。

### 処理パイプライン（main関数内の順序）

1. **ページ走査**: `rawdict`で文字単位のbbox・originを取得し、y座標でグループ化して左右カラムに分離（`extract_page_lines`）
2. **アンダーライン検出**: drawing矩形からアンダーラインを識別。同一y座標の近接矩形をマージ（`UL_MERGE_GAP=3pt`）後、文字中心xでマッチング（`get_underlines` → `_char_is_underlined`）
3. **セグメント構築**: 文字単位のUL状態を連続する同一状態のテキスト区間にグループ化（`_build_segments`）。各行のデータは `[(text, is_underlined), ...]` のセグメントリスト
4. **階層追跡**: `HierarchyTracker`が章→部→節→款→区分番号→注の階層を逐次更新。各行処理後にsnapshotを保存
5. **変更ブロック検出**: アンダーライン行を起点に、`should_split_block`でブロック境界を判定し、`extend_block_context`で前後に文脈拡張
6. **後処理**（3パス）:
   - 項目名補完: ページまたぎで括弧が未完結の項目名を`completed_names`辞書で補完
   - 注フィールド判定: 明示的「注N」→ 裸の数字 → tracker_noteフォールバックの3段階
   - マーカークリーンアップ: （新設）/（削る）を含むセルからの混入テキスト除去
7. **XLSX出力**: xlsxwriterの`write_rich_string`でアンダーライン箇所に下線書式を適用

### 重要なデータ構造

- **行データ**: `left_segments` / `right_segments` は `[(text, is_underlined), ...]` のセグメントリスト。`left_underlined` / `right_underlined` は行全体にULがあるかのフラグ
- **ブロック**: `{'page', 'context', 'left_lines', 'right_lines'}` — `left_lines`は `[[(text, is_ul), ...], ...]`（外側=行、内側=セグメント）
- **carry_block**: ページ境界をまたぐ未完了ブロックを次ページに持ち越す仕組み

### アンダーライン検出の仕組み

PDF上のアンダーラインはテキスト属性ではなく独立した描画矩形（`page.get_drawings()`）として存在する。検出は3段階：

1. **矩形抽出**: 高さ < 2pt、幅 > 5pt の矩形を収集。罫線（x=93.86, 420.41, 747.09付近）は除外
2. **矩形マージ**: 同一y座標（差 < 1pt）でx方向ギャップ < 3pt の矩形を統合。3ptを超えるギャップはマージしない（15ptだと別々のUL矩形を誤統合して偽陽性が出る）
3. **文字マッチング**: `rawdict`の文字bboxの下端y1とUL矩形のy0の差が < 4pt、かつ文字中心cxがUL矩形のx範囲内であればUL判定

### XLSX出力の注意点

- **xlsxwriterを使用すること**（openpyxlのCellRichTextはExcelで「修復」が必要になりリッチテキスト書式が壊れる）
- `write_rich_string`で `[format, text, format, text, ...]` 形式の引数リストを構築
- 各formatには必ずフォント名・サイズを指定（`font_name='游ゴシック'`, `font_size=11`）

### PDFレイアウトの前提

- 左カラム（x < 420.0）= 改正後、右カラム = 改正前
- 罫線x座標: `[93.86, 420.41, 747.09]`（これらはアンダーラインから除外）
- 目次は先頭4ページ（変更ブロックとしては出力しないが、階層追跡には使用）
- テキストは全角文字（区分番号: `Ａ０００`形式、数字: `１２３`形式）

## 作業上の注意

- PDFファイルは公式資料の原本であり、**編集・変更しないこと**
- Windows環境で作業しており、日本語ファイル名を含むためエンコーディング（UTF-8）に注意
- 親ディレクトリ `my-project/` 配下に関連プロジェクト（`R8.sinryohousyukaitei/`, `DPC_support/`）があるが、これらは参照せず独立して思考すること
