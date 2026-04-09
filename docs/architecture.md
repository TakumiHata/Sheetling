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
_setup_grid_params()            ← src/core/pipeline.py
    │  用紙サイズ（A4/A3）・向き（縦/横）を検出し GRID_SIZES から定数を選択
    ▼
_compute_grid_coords()          ← src/core/pipeline.py
    │  コンテンツ境界を検出し、PDF座標 → Excelグリッド座標に変換
    │  各要素に _row / _col を付与、table_border_rects を生成
    ▼
_auto_generate_layout()         ← src/core/pipeline.py
    │  抽出データからレイアウト JSON（text + border_rect 要素）を生成
    ▼
    ├─→ _render_layout_to_xlsx()    → Excel ファイル (.xlsx)
    ├─→ _generate_border_preview()  → 罫線プレビュー画像 (.png)
    └─→ VISUAL_REVIEW_PROMPT        → 視覚的検証プロンプト (.txt)
              │
              ▼  （任意）ビジョンLLM による修正指示
              │
        apply_corrections()         → layout JSON を更新
              │
        rerender_after_corrections() → Excel 再生成
```

---

## モジュール構成

```
src/
├── main.py              # CLI エントリポイント（auto / correct / check）
├── core/
│   └── pipeline.py      # パイプライン全体の制御・座標計算・レイアウト生成・Excel描画
├── parser/
│   └── pdf_extractor.py # PDF データ抽出（pdfplumber ラッパー）
├── templates/
│   └── prompts.py       # グリッドサイズ定数・視覚的検証プロンプト
└── utils/
    └── logger.py        # ロガー設定
```

---

## CLI コマンドと呼び出し関係

### `auto` コマンド（`src/main.py`）

`data/in/` 配下のPDFを検出し、グリッドサイズ `1pt` / `2pt` で各PDFに対して `pipeline.auto_layout()` を呼び出します。

### `correct` コマンド（`src/main.py`）

修正ファイル（`*_visual_corrections*.json`）から対象の (pdf_name, grid_size) ペアを自動検出し、以下を順に実行します：

1. `pipeline.apply_corrections()` — 修正指示を layout JSON に適用
2. `pipeline.rerender_after_corrections()` — Excel を再生成

### `check` コマンド（`src/main.py`）

pdfplumber の `extract_text()` でテキスト有無を判定し、結果CSVを出力します。パイプラインクラスは使用しません。

---

## `auto_layout()` の処理ステップ

`SheetlingPipeline.auto_layout()` は以下の順序で処理を実行します：

| ステップ | 処理 | 呼び出す関数 |
|---------|------|-----------|
| 1 | PDF からデータ抽出 | `extract_pdf_data()` |
| 2 | グリッドパラメータ設定 | `_setup_grid_params()` |
| 3 | PDF座標 → グリッド座標変換 | `_compute_grid_coords()` |
| 4 | 抽出データ JSON 保存 | — |
| 5 | レイアウト JSON 生成 | `_auto_generate_layout()` |
| 6 | Excel 描画 | `_render_layout_to_xlsx()` |
| 7 | 元 PDF を出力先にコピー | `shutil.copy()` |
| 8 | PDF ページ画像生成 | pdfplumber `page.to_image()` |
| 9 | 罫線プレビュー画像生成 | `_generate_border_preview()` |
| 10 | 検証プロンプト・空テンプレート出力 | `VISUAL_REVIEW_PROMPT` |

---

## 主要関数リファレンス

### `src/parser/pdf_extractor.py`

| 関数 | 役割 |
|------|------|
| `extract_pdf_data(pdf_path)` | pdfplumber で PDF を解析し、ページごとに words / table_cells / rects / edges を抽出 |
| `_remove_containing_rects(rects)` | 他の矩形を完全に内包する矩形を除去 |
| `_to_hex_color(color)` | カラー値を `#RRGGBB` 形式に変換 |

### `src/core/pipeline.py` — ヘルパー関数

| 関数 | 役割 |
|------|------|
| `_normalize_font_name()` | PDF フォント名を Excel 用に正規化（MS Gothic 等） |
| `_compute_grid_coords()` | コンテンツ境界ベースで PDF 座標をグリッド座標に変換 |
| `_setup_grid_params()` | 用紙サイズ・向きから GRID_SIZES を選択 |
| `_auto_generate_layout()` | 抽出データからレイアウト JSON を生成 |
| `_render_layout_to_xlsx()` | レイアウト JSON を openpyxl で Excel に描画 |
| `_generate_border_preview()` | 罫線プレビュー PNG を生成（グレーアウト・ラベル付き） |
| `_fix_empty_cell_type_attr()` | Excel Online 互換性のための属性修正 |
| `_split_by_horizontal_gap()` | フォントサイズベースの水平ギャップでワードを分割 |
| `_join_word_texts()` | 日本語判定してワードを結合（日本語: スペースなし） |
| `_table_text_elements_from_2d()` | テーブル内テキストを word 優先で配置 |
| `_has_japanese()` | テキストに日本語が含まれるか判定 |

### `src/core/pipeline.py` — `SheetlingPipeline` クラス

| メソッド | 役割 |
|---------|------|
| `auto_layout(pdf_path, grid_size)` | PDF → Excel 全自動パイプライン |
| `apply_corrections(...)` | 修正指示 JSON を layout JSON に適用 |
| `rerender_after_corrections(...)` | 修正済み layout JSON から Excel を再生成 |

### `src/templates/prompts.py`

| 定数 | 役割 |
|------|------|
| `GRID_SIZES` | グリッドサイズ別の定数定義（A4/A3 × 1pt/2pt） |
| `VISUAL_REVIEW_PROMPT` | ビジョン LLM 用の検証プロンプトテンプレート |

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
  "rects": [{"x0", "top", "x1", "bottom"}],
  "h_edges": [{"x0", "x1", "y"}],
  "v_edges": [{"x", "y0", "y1"}]
}
```

---

## レイアウト JSON の構造

`_auto_generate_layout()` が生成する要素：

### text 要素

```json
{
  "type": "text",
  "content": "テキスト内容",
  "row": 5,
  "col": 3,
  "end_col": 10,
  "font_color": "#000000",
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
  "borders": {"top": true, "bottom": true, "left": true, "right": true}
}
```

`borders` の各辺は独立して `true` / `false` を指定可能。水平罫線なら `top` のみ `true`、垂直罫線なら `left` のみ `true` になります。
