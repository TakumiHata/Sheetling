"""
Sheetling 用の LLM プロンプト定義集。

  VISUAL_REVIEW_PROMPT — ビジョンLLMによる視覚的検証用プロンプト
  GRID_SIZES           — 対応グリッドサイズ定義 (1pt / 2pt)
"""

GRID_SIZES = {
    # =========================================================================
    # A4 (595pt × 842pt)
    # =========================================================================
    # Sheetling "1pt": A4縦 54列×46行 / A4横 79列×31行（列幅1.00表示）
    "1pt": {
        "col_width_mm": "3.48",
        "row_height_mm": "6.44",
        # A4縦
        "max_cols": 54,
        "max_rows": 46,
        # A4横
        "max_cols_landscape": 79,
        "max_rows_landscape": 31,
        "excel_col_width": 1.74,  # W = 表示値 + 0.74 (2pt と同じオフセット則; 1.69 では離散ステップを越えず 0.94 のまま)
        "excel_row_height": 18.25,
        "margin_left": 0.43,
        "margin_right": 0.43,
        "margin_top": 0.41,
        "margin_bottom": 0.41,
        "default_font_size": 7,
        "font_name": "MS Gothic",
        "position_tolerance_cells": "1〜2",
    },
    # Sheetling "2pt": A4縦 34列×46行 / A4横 50列×31行（列幅2.00表示）
    "2pt": {
        "col_width_mm": "6.18",
        "row_height_mm": "6.44",
        # A4縦
        "max_cols": 34,
        "max_rows": 46,
        # A4横
        "max_cols_landscape": 50,
        "max_rows_landscape": 31,
        "excel_col_width": 2.74,  # 経験的に列幅2.00表示となる値（MDW≈7.0環境）
        "excel_row_height": 18.25,
        "margin_left": 0.43,
        "margin_right": 0.43,
        "margin_top": 0.41,
        "margin_bottom": 0.41,
        "default_font_size": 6,
        "font_name": "MS Gothic",
        "position_tolerance_cells": "1",  # 4mm/セルと粗いため厳しく
    },
    # =========================================================================
    # A3 (842pt × 1190pt) — セルサイズは A4 と同一、用紙が大きい分だけ列数・行数が増える
    # =========================================================================
    # A3 1pt: A3縦 102列×65行 / A3横 144列×46行（列幅1.00表示）
    # ※A3縦は手元PDFがないため A3横 値からアスペクト比 (297/420, 420/297) で予測
    "1pt_a3": {
        "col_width_mm": "3.48",
        "row_height_mm": "6.44",
        # A3縦
        "max_cols": 102,
        "max_rows": 65,
        # A3横
        "max_cols_landscape": 144,
        "max_rows_landscape": 46,
        "excel_col_width": 1.74,
        "excel_row_height": 18.25,
        "margin_left": 0.43,
        "margin_right": 0.43,
        "margin_top": 0.41,
        "margin_bottom": 0.41,
        "default_font_size": 7,
        "font_name": "MS Gothic",
        "position_tolerance_cells": "1〜2",
    },
    # A3 2pt: A3縦 52列×65行 / A3横 73列×46行（列幅2.00表示）
    # ※A3縦は手元PDFがないため A3横 値からアスペクト比 (297/420, 420/297) で予測
    "2pt_a3": {
        "col_width_mm": "6.18",
        "row_height_mm": "6.44",
        # A3縦
        "max_cols": 52,
        "max_rows": 65,
        # A3横
        "max_cols_landscape": 73,
        "max_rows_landscape": 46,
        "excel_col_width": 2.74,
        "excel_row_height": 18.25,
        "margin_left": 0.43,
        "margin_right": 0.43,
        "margin_top": 0.41,
        "margin_bottom": 0.41,
        "default_font_size": 6,
        "font_name": "MS Gothic",
        "position_tolerance_cells": "1",
    },
}


# ビジョンLLM（画像入力対応AIチャット）向けの視覚的検証プロンプト。
# JSONは渡さず、PDFページ画像 + 自動生成した罫線プレビュー画像の2枚比較で罫線差分を検出する。
# ページごとに生成し、対応する2枚のPNG画像と一緒にLLMに投入して使用する。
VISUAL_REVIEW_PROMPT = """\
以下の2枚の画像を比較し、**罫線（枠線）の過不足**を報告してください。

- 画像1: PDFの原本（ページ {page_number}）
- 画像2: 自動生成した罫線プレビュー（ページ {page_number}）

## 画像2（罫線プレビュー）の見方

- 薄いグレーの縦横線はグリッド背景線です。**罫線ではありません。無視してください。**
- 太い黒線のみが罫線（枠線）です。
- グリッドサイズ: {max_rows} 行 × {max_cols} 列（1マス = {col_width_mm}mm × {row_height_mm}mm）

### 座標の読み方（重要）

赤いラベルは **セルの中央** に表示されており、**ラベルの数値 = JSON の `row`/`col` 値** に直接対応します。

- ラベル `1` のセル → `col=1`（または `row=1`）
- ラベル `6` のセル → `col=6`（または `row=6`）
- ラベルのないセルはその前後の数値から数えてください（例: ラベル `6` の3つ右 → `col=9`）
- 左上セルが `(row=1, col=1)`、右・下方向に増加します

## 判定基準

**報告してよいもの**
- PDF に存在するが、プレビューに描画されていない罫線 → `add_border`
- プレビューに描画されているが、PDF に存在しない罫線 → `remove_border`
- 罫線の範囲（開始・終了位置）がPDFと明らかに異なる場合 → `remove_border` + `add_border` で修正

**報告してはいけないもの（無視してください）**
- テキスト・文字の差異（フォント・配置・内容の違いはすべて無視）
- グリッド背景線（薄いグレー線）
- PDF の薄い罫線・飾り線・影など、Excel で表現不要な装飾的な線

## 座標の範囲制約（厳守）

コンテンツが配置されている有効範囲は **row: 1〜{content_max_row}、col: 1〜{content_max_col}** です。
この範囲外の座標を corrections に指定しないでください。
罫線の `end_row` は {content_max_row} 以下、`end_col` は {content_max_col} 以下としてください。

## 出力形式

差異がない場合は `{{"corrections": []}}` のみ出力してください。

### フィールド名の厳守
- 列の終端: **`end_col`**（`col_end` は誤り）
- 行の終端: **`end_row`**（`row_end` は誤り）

### フォーマット

```json
{{
  "corrections": [
    {{
      "action": "add_border",
      "page": {page_number},
      "row": <開始行>,
      "end_row": <終了行>,
      "col": <開始列>,
      "end_col": <終了列>,
      "borders": {{"top": true, "bottom": true, "left": true, "right": true}}
    }},
    {{
      "action": "remove_border",
      "page": {page_number},
      "row": <開始行>,
      "end_row": <終了行>,
      "col": <開始列>,
      "end_col": <終了列>
    }}
  ]
}}
```

### 記入例（下線のみ追加する場合）
```json
{{"corrections": [{{"action": "add_border", "page": {page_number}, "row": 5, "end_row": 6, "col": 3, "end_col": 12, "borders": {{"bottom": true, "top": false, "left": false, "right": false}}}}]}}
```

【最重要】出力はJSONのみ。説明文・前置き・コードブロック記号（```）は一切不要です。
【最重要】フィールド名は `end_col`・`end_row` を使用すること（`col_end`・`row_end` は誤り）。"""
