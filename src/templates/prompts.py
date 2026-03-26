"""
Sheetling 用の LLM プロンプト定義集。

  VISUAL_REVIEW_PROMPT — ビジョンLLMによる視覚的検証用プロンプト
  GRID_SIZES           — 対応グリッドサイズ定義 (1pt / 2pt)
"""

GRID_SIZES = {
    # Sheetling "1pt": A4縦 62列×42行 / A4横 96列×30行（列幅1.00表示）
    # 縦横のグリッド数は A4 ポイント寸法とセル幅から算出後に余白列を加算:
    #   縦: 595pt÷(595/57)pt ≈ 57列 + 5列 = 62列, 842pt÷(842/42)pt ≈ 42行
    #   横: 842pt÷(595/57)pt ≈ 81列 + 15列 = 96列, 595pt÷(842/42)pt ≈ 30行
    "1pt": {
        "col_width_mm": "3.48",
        "row_height_mm": "6.44",
        # A4縦
        "max_cols":            62,   # 算出値57 + 余白5列
        "max_rows":            42,
        # A4横（縦と同一セルサイズ、枚数が変わる）
        "max_cols_landscape":  96,   # 算出値81 + 余白15列
        "max_rows_landscape":  30,
        "excel_col_width": 1.625,    # (1*8+5)/8: デスクトップExcel(MDW=8)で列幅1.00表示
        "excel_row_height": 18.25,
        "margin_left": 0.43,
        "margin_right": 0.43,
        "margin_top": 0.41,
        "margin_bottom": 0.41,
        "default_font_size": 7,
        "font_name": "MS Gothic",
        "position_tolerance_cells": "1〜2",
    },
    # 粗めセル密度（A4縦: 37列×42行 / A4横: 58列×30行、列幅2.00表示）
    # 縦横のグリッド数は A4 ポイント寸法とセル幅から算出後に余白列を加算:
    #   縦: 595pt÷(595/34)pt ≈ 34列 + 3列 = 37列, 842pt÷(842/42)pt ≈ 42行
    #   横: 842pt÷(595/34)pt ≈ 48列 + 10列 = 58列, 595pt÷(842/42)pt ≈ 30行
    "2pt": {
        "col_width_mm": "6.18",
        "row_height_mm": "6.44",
        # A4縦
        "max_cols":            37,   # 算出値34 + 余白3列
        "max_rows":            42,
        # A4横（縦と同一セルサイズ、枚数が変わる）
        "max_cols_landscape":  58,   # 算出値48 + 余白10列
        "max_rows_landscape":  30,
        "excel_col_width": 2.625,    # (2*8+5)/8: デスクトップExcel(MDW=8)で列幅2.00表示
        "excel_row_height": 18.25,
        "margin_left": 0.43,
        "margin_right": 0.43,
        "margin_top": 0.41,
        "margin_bottom": 0.41,
        "default_font_size": 6,
        "font_name": "MS Gothic",
        "position_tolerance_cells": "1",  # 4mm/セルと粗いため厳しく
    },
}


# ビジョンLLM（画像入力対応AIチャット）向けの視覚的検証プロンプト。
# JSONは渡さず、PDFページ画像 + 自動生成した罫線プレビュー画像の2枚比較で罫線差分を検出する。
# ページごとに生成し、対応する2枚のPNG画像と一緒にLLMに投入して使用する。
VISUAL_REVIEW_PROMPT = """\
以下の2枚の画像を比較し、**罫線（枠線）の明確な過不足のみ**を報告してください。

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

## 判定基準（厳格に守ってください）

**報告してよいもの（明確な差異のみ）**
- PDF に明らかに存在するが、プレビューに全く描画されていない罫線 → `add_border`
- プレビューに描画されているが、PDF には明らかに存在しない罫線 → `remove_border`

**報告してはいけないもの（無視してください）**
- テキスト・文字の差異（フォント・配置・内容の違いはすべて無視）
- グリッド背景線（薄いグレー線）
- 罫線の位置が {position_tolerance_cells} セルずれている程度の微小なズレ
- PDF の薄い罫線・飾り線・影など、Excel で表現不要な装飾的な線
- すでにプレビューに描画されている罫線を「位置修正」するような操作
- 判断に迷う・曖昧な差異（確信が持てない場合は報告しない）

**`remove_border` は特に慎重に使ってください。**
プレビューに罫線があり、PDFにも類似した線がある場合は、削除を提案しないでください。
明らかに余分な罫線（PDFのどこにも対応する線がない）にのみ使用してください。

## 出力形式

差異がない、または軽微な場合は `{{"corrections": []}}` のみ出力してください。

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
【最重要】疑わしい差異は報告しない。確実な差異のみ報告する。
【最重要】フィールド名は `end_col`・`end_row` を使用すること（`col_end`・`row_end` は誤り）。"""
