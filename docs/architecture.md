# アーキテクチャ

Sheetling のパイプライン全体のデータフロー、主要関数の役割と呼び出し関係を説明します。

---

## 全体データフロー

```
PDF ファイル
    │
    ▼
extract_pdf_data()              ← src/parser/pdf_extractor.py
    │  words / table_cells / rects / h_edges / v_edges を抽出
    ▼
setup_grid_params()             ← src/core/grid.py
    │  用紙サイズ（A4/A3）・向き（縦/横）を検出し GRID_SIZES から定数を選択
    ▼
compute_grid_coords()           ← src/core/grid.py
    │  コンテンツ境界を検出し、PDF座標 → Excelグリッド座標に変換
    │  各要素に _row / _col を付与、table_border_rects を生成
    ▼
generate_layout()               ← src/core/layout.py
    │  抽出データからレイアウト JSON（text + border_rect 要素）を生成
    ▼
    ├─→ render_layout_to_xlsx()     ← src/renderer/excel.py    → Excel ファイル (.xlsx)
    ├─→ generate_border_preview()   ← src/renderer/preview.py  → 罫線プレビュー画像 (.png)
    ├─→ generate_diff_overlay()     ← src/renderer/preview.py  → 差分オーバーレイ画像 (.png)
    └─→ VISUAL_PHASE1/2_PROMPT      ← src/templates/prompts.py → フェーズ1/2プロンプト (.txt)
              │
              ▼  （任意）ビジョンLLM による修正指示（2フェーズ）
              │
        apply_corrections()         ← src/core/correction_service.py → layout JSON を更新
              │
        rerender_after_corrections() ← src/core/correction_service.py → Excel 再生成
```

---

## モジュール構成

```
src/
├── main.py                # CLI エントリポイント（auto / correct / check）
├── core/
│   ├── pipeline.py        # パイプラインファサード（SheetlingPipeline クラス）
│   ├── auto_layout_service.py  # auto パイプライン実装・検証素材生成
│   ├── correction_service.py   # correct パイプライン実装・修正指示適用
│   ├── edges.py           # エッジ単位罫線モデル（分解・集約・修正適用）
│   ├── grid.py            # グリッド座標計算・コンテンツ境界検出・用紙サイズ検出
│   ├── grid_config.py     # GRID_SIZES 定数（A4/A3 × 1pt/2pt のセル寸法・Excel設定）
│   ├── constants.py       # 共有の数値定数（tolerance・閾値など）
│   └── layout.py          # レイアウトJSON生成（罫線要素・テキスト要素の収集）
├── parser/
│   └── pdf_extractor.py   # PDF データ抽出（pdfplumber ラッパー）
├── renderer/
│   ├── excel.py           # Excel 描画（openpyxl による .xlsx 生成）
│   └── preview.py         # 罫線プレビュー・差分オーバーレイ画像生成（Pillow）
├── templates/
│   └── prompts.py         # フェーズ1/2の視覚的検証プロンプト
└── utils/
    ├── logger.py          # ロガー設定
    ├── font.py            # フォント名正規化・罫線スタイルマッピング
    └── text.py            # テキスト結合・日本語判定・水平ギャップ分割
```

---

## CLI コマンドと呼び出し関係

### `auto` コマンド（`src/main.py`）

`data/in/` 配下のPDFを検出し、グリッドサイズ `1pt` / `2pt` で各PDFに対して `pipeline.auto_layout()` を呼び出します。

### `correct` コマンド（`src/main.py`）

修正ファイル（`*_phase1_corrections*.json` / `*_phase2_corrections*.json`、旧フォーマットにもフォールバック）から対象の (pdf_name, grid_size) ペアを自動検出し、以下を順に実行します：

1. phase1・phase2 の corrections をマージ
2. `pipeline.apply_corrections()` — 修正指示を layout JSON に適用
3. `pipeline.rerender_after_corrections()` — Excel を再生成

### `check` コマンド（`src/main.py`）

pdfplumber の `extract_text()` でテキスト有無を判定し、結果CSVを出力します。パイプラインクラスは使用しません。

---

## `auto_layout()` の処理ステップ

`SheetlingPipeline.auto_layout()` は以下の順序で処理を実行します：

| ステップ | 処理 | 呼び出す関数 / モジュール |
|---------|------|----------------------|
| 1 | PDF からデータ抽出 | `extract_pdf_data()` ← `parser/pdf_extractor.py` |
| 2 | グリッドパラメータ設定 | `setup_grid_params()` ← `core/grid.py` |
| 3 | PDF座標 → グリッド座標変換 | `compute_grid_coords()` ← `core/grid.py` |
| 4 | 抽出データ JSON 保存 | — |
| 5 | レイアウト JSON 生成 | `generate_layout()` ← `core/layout.py` |
| 6 | Excel 描画 | `render_layout_to_xlsx()` ← `renderer/excel.py` |
| 7 | 元 PDF を出力先にコピー | `shutil.copy()` |
| 8 | PDF ページ画像生成 | pdfplumber `page.to_image()` |
| 9 | 罫線プレビュー画像生成 | `generate_border_preview()` ← `renderer/preview.py` |
| 10 | 差分オーバーレイ画像生成 | `generate_diff_overlay()` ← `renderer/preview.py` |
| 11 | フェーズ1/2プロンプト・空 JSON 出力 | `VISUAL_PHASE1/2_PROMPT` ← `templates/prompts.py` |

---

## 主要関数リファレンス

### `src/parser/pdf_extractor.py`

| 関数 | 役割 |
|------|------|
| `extract_pdf_data(pdf_path)` | pdfplumber で PDF を解析し、ページごとに words / table_cells / rects / edges を抽出 |
| `_extract_words(page)` | テキスト抽出（フォント情報・色・縦書き文字を含む） |
| `_extract_tables(page)` | テーブル検出・セルbbox・2D配列データの取得 |
| `_extract_rects(page, page_area)` | 矩形抽出（ページ外枠・包含矩形を除去） |
| `_extract_edges(page, page_area)` | 水平・垂直エッジの収集・重複排除・セグメント統合 |
| `_remove_containing_rects(rects)` | 他の矩形を完全に内包する矩形を除去 |
| `_to_hex_color(color)` | カラー値を `RRGGBB` 形式に変換 |

### `src/core/grid.py`

| 関数 | 役割 |
|------|------|
| `compute_grid_coords(page, max_rows, max_cols)` | コンテンツ境界ベースで PDF 座標をグリッド座標に変換 |
| `setup_grid_params(first_page, grid_size)` | 用紙サイズ・向きから GRID_SIZES を選択 |
| `_detect_content_bounds(page, page_h)` | 全要素の座標からコンテンツ領域の min/max を検出 |
| `_merge_thin_lines_to_rects(page)` | 4本の線（上辺+下辺+左辺+右辺）から矩形を統合 |
| `_assign_rect_grid_coords(page, ...)` | 矩形にグリッド座標を付与し、テーブル内矩形を除外 |
| `_build_table_border_rects(page, ...)` | テーブルの cells_2d から border_rect を生成 |

### `src/core/layout.py`

| 関数 | 役割 |
|------|------|
| `generate_layout(extracted_data, grid_params)` | 抽出データからレイアウト JSON を生成 |
| `_collect_table_border_elements(page, ...)` | テーブル罫線 → border_rect 要素 |
| `_collect_rect_border_elements(page, ...)` | テーブル外矩形 → border_rect 要素（水平線・垂直線を含む） |
| `_collect_edge_border_elements(page, ...)` | h_edges / v_edges → border_rect 要素 |
| `_collect_text_elements(page, ...)` | テーブル外 words → text 要素 |
| `_table_text_elements_from_2d(page, grid_params)` | テーブル内テキストを word 優先で配置 |
| `_make_text_element(word_group, ...)` | word グループから text 要素を生成（フォント情報付き） |

### `src/renderer/excel.py`

| 関数 | 役割 |
|------|------|
| `render_layout_to_xlsx(layout, grid_params, output_path)` | レイアウト JSON を openpyxl で Excel に描画 |
| `fix_empty_cell_type_attr(xlsx_path)` | Excel Online 互換性のための `t="n"` 属性除去 |
| `_place_text_element(ws, elem, ...)` | テキスト要素をセルに配置（フォント・配置設定） |
| `_place_border_element(ws, elem, ...)` | 罫線要素をセルに描画（4辺独立） |

### `src/renderer/preview.py`

| 関数 | 役割 |
|------|------|
| `generate_border_preview(page_layout, grid_params, ...)` | 罫線プレビュー PNG を生成 |
| `generate_diff_overlay(pdf_image_path, runs_with_ids, ...)` | PDF原本に現プレビュー罫線を半透明赤で重ねた差分画像を生成 |
| `_draw_grid_lines(draw, ...)` | グリッド背景線を描画 |
| `_draw_borders(draw, ...)` | border_rect 要素を黒線で描画 |
| `_draw_greyout(draw, ...)` | コンテンツ範囲外をグレーアウト |
| `_draw_labels(draw, ...)` | 座標ラベル（赤、5間隔）を描画 |
| `_draw_run_overlay(draw, ...)` | ID付きランを色付き線で描画（差分オーバーレイ用） |
| `_draw_run_id_labels(draw, ...)` | 各ランに ID ラベルを描画 |

### `src/utils/font.py`

| 関数 | 役割 |
|------|------|
| `normalize_font_name(raw_name)` | PDF フォント名のサブセットプレフィックス（`ABCDEF+`）を除去して返す |
| `linewidth_to_border_style(linewidth)` | PDF の linewidth を Excel 罫線スタイル（thin/medium/thick）に変換 |

### `src/utils/text.py`

| 関数 | 役割 |
|------|------|
| `has_japanese(text)` | テキストに日本語文字が含まれるか判定 |
| `join_word_texts(texts)` | 日本語判定してワードを結合（日本語: スペースなし） |
| `split_by_horizontal_gap(words, gap_factor)` | フォントサイズベースの水平ギャップでワードを分割 |

### `src/core/pipeline.py` — `SheetlingPipeline` クラス

| メソッド | 役割 |
|---------|------|
| `auto_layout(pdf_path, grid_size)` | PDF → Excel 全自動パイプライン |
| `apply_corrections(...)` | 修正指示 JSON を layout JSON に適用 |
| `rerender_after_corrections(...)` | 修正済み layout JSON から Excel を再生成 |

### `src/core/grid_config.py`

| 定数 | 役割 |
|------|------|
| `GRID_SIZES` | グリッドサイズ別の定数定義（A4/A3 × 1pt/2pt） |

### `src/core/edges.py`

| 関数 | 役割 |
|------|------|
| `decompose_to_cell_edges(elements)` | border_rect 要素群をセル境界の集合に分解 |
| `group_into_runs(cell_edges, styles)` | 連続するセル境界を最大長のランに集約 |
| `runs_to_border_rects(runs)` | ランを border_rect 要素に変換（H: top のみ、V: left のみ） |
| `enumerate_runs_with_ids(elements)` | layout の border_rect 群を ID 付きランリストに変換（内部 exclusive 座標） |
| `apply_edge_corrections(elements, removed_ids, added_runs, id_map)` | エッジ単位の修正を elements に in-place 適用 |

### `src/core/correction_service.py`

| 関数 / クラス | 役割 |
|------|------|
| `CorrectionService` | corrections JSON の適用とリレンダーを担当するクラス |
| `CorrectionService.apply(...)` | corrections JSON を読み込んで layout JSON に適用 |
| `CorrectionService.rerender(...)` | 修正済み layout JSON から Excel を再生成 |

### `src/templates/prompts.py`

| 定数 | 役割 |
|------|------|
| `VISUAL_PHASE1_PROMPT` | フェーズ1用プロンプト（余分な罫線の削除 ID を判定） |
| `VISUAL_PHASE2_PROMPT` | フェーズ2用プロンプト（不足している罫線を inclusive 座標で追加） |

---

## 抽出データの構造

`extract_pdf_data()` が返すページデータ：

```python
{
  "page_number": int,
  "width": float,
  "height": float,
  "words": [
    {"x0", "top", "x1", "bottom", "text", "fontname", "font_size", "font_color", "is_vertical"}
  ],
  "table_bboxes": [[x0, top, x1, bottom]],
  "table_col_x_positions": [[x座標]],
  "table_row_y_positions": [[y座標]],
  "table_cells": [[[{"x0", "top", "x1", "bottom"} | None]]],  # None = 結合延長
  "table_data": [[セル内容文字列]],
  "table_data_raw": [[セル内容（改行保持）]],
  "rects": [{"x0", "top", "x1", "bottom", "linewidth"}],
  "h_edges": [{"x0", "x1", "y", "linewidth"}],
  "v_edges": [{"x", "y0", "y1", "span", "linewidth"}]
}
```

---

## レイアウト JSON の構造

`generate_layout()` が生成する要素：

### text 要素

```json
{
  "type": "text",
  "content": "テキスト内容",
  "row": 5,
  "col": 3,
  "end_col": 10,
  "font_color": "FF0000",
  "font_size": 10,
  "font_name": "MS Gothic"
}
```

### border_rect 要素

```json
{
  "type": "border_rect",
  "row": 3,
  "end_row": 8,
  "col": 2,
  "end_col": 15,
  "borders": {"top": true, "bottom": true, "left": true, "right": true},
  "border_style": "thin"
}
```

`borders` の各辺は独立して `true` / `false` を指定可能。水平罫線なら `top` のみ `true`、垂直罫線なら `left` のみ `true` になります。
`border_style` は PDF の linewidth から `thin` / `medium` / `thick` に変換されます。

---

## テスト構成

```
tests/
├── test_font.py            # utils/font.py のユニットテスト
├── test_text.py            # utils/text.py のユニットテスト
├── test_grid.py            # core/grid.py のユニットテスト
├── test_layout.py          # core/layout.py のユニットテスト
├── test_excel.py           # renderer/excel.py のユニットテスト
├── test_preview.py         # renderer/preview.py のユニットテスト
├── test_pdf_extractor.py   # parser/pdf_extractor.py のユニットテスト
└── test_pipeline.py        # core/pipeline.py のユニットテスト
```

各モジュールの公開関数・ヘルパー関数を個別にテストしています。

```bash
python -m pytest tests/ -v
```

---

## 関連ドキュメント

- [グリッドシステム](grid-system.md) — コンテンツ境界ベースのグリッド計算・座標変換の詳細
- [テーブル検出とテキスト配置](table-detection.md) — pdfplumber パラメータ・word 優先フォールバック戦略
- [correct ワークフロー](correct-workflow.md) — corrections JSON 仕様・安全機構の詳細
- [チューニングガイド](tuning-guide.md) — GRID_SIZES 調整・トラブルシューティング
