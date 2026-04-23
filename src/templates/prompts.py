"""
Sheetling 用の LLM プロンプト定義集。

  VISUAL_REVIEW_PROMPT — ビジョンLLMによる視覚的検証用プロンプト
  GRID_SIZES           — 対応グリッドサイズ定義 (1pt / 2pt)
"""

#
# max_cols/max_rows はコンテンツが配置される実セル数。
# Excel 上の印刷総セル数は左1列・上1行のパディングを足した値となる。
#   印刷総列数 = max_cols + 1, 印刷総行数 = max_rows + 1
# 例: A4横 1pt → 印刷総 79列×31行 → max_cols=78, max_rows=30
#
GRID_SIZES = {
    # =========================================================================
    # A4 (595pt × 842pt)
    # =========================================================================
    # Sheetling "1pt": 印刷総 A4縦 54列×46行 / A4横 79列×31行（列幅1.00表示）
    "1pt": {
        "col_width_mm": "3.48",
        "row_height_mm": "6.44",
        # A4縦 (印刷総 54×46 → コンテンツ 53×45)
        "max_cols": 53,
        "max_rows": 45,
        # A4横 (印刷総 79×31 → コンテンツ 78×30)
        "max_cols_landscape": 78,
        "max_rows_landscape": 30,
        "excel_col_width": 1.74,  # W = 表示値 + 0.74 (2pt と同じオフセット則; 1.69 では離散ステップを越えず 0.94 のまま)
        "excel_row_height": 18.25,
        "margin_left": 0.43,
        "margin_right": 0.43,
        "margin_top": 0.41,
        "margin_bottom": 0.41,
        "default_font_size": 7,
        "font_name": "MSゴシック",
        "position_tolerance_cells": "1〜2",
    },
    # Sheetling "2pt": 印刷総 A4縦 34列×46行 / A4横 50列×31行（列幅2.00表示）
    "2pt": {
        "col_width_mm": "6.18",
        "row_height_mm": "6.44",
        # A4縦 (印刷総 34×46 → コンテンツ 33×45)
        "max_cols": 33,
        "max_rows": 45,
        # A4横 (印刷総 50×31 → コンテンツ 49×30)
        "max_cols_landscape": 49,
        "max_rows_landscape": 30,
        "excel_col_width": 2.74,  # 経験的に列幅2.00表示となる値（MDW≈7.0環境）
        "excel_row_height": 18.25,
        "margin_left": 0.43,
        "margin_right": 0.43,
        "margin_top": 0.41,
        "margin_bottom": 0.41,
        "default_font_size": 6,
        "font_name": "MSゴシック",
        "position_tolerance_cells": "1",  # 4mm/セルと粗いため厳しく
    },
    # =========================================================================
    # A3 (842pt × 1190pt) — セルサイズは A4 と同一、用紙が大きい分だけ列数・行数が増える
    # =========================================================================
    # A3 1pt: 印刷総 A3縦 81列×65行 / A3横 114列×46行（列幅1.00表示）
    # ※A3縦は手元PDFがないため A3横 値からアスペクト比 (297/420, 420/297) で予測
    "1pt_a3": {
        "col_width_mm": "3.48",
        "row_height_mm": "6.44",
        # A3縦 (印刷総 81×65 → コンテンツ 80×64)
        "max_cols": 80,
        "max_rows": 64,
        # A3横 (印刷総 114×46 → コンテンツ 113×45)
        "max_cols_landscape": 113,
        "max_rows_landscape": 45,
        "excel_col_width": 1.74,
        "excel_row_height": 18.25,
        "margin_left": 0.43,
        "margin_right": 0.43,
        "margin_top": 0.41,
        "margin_bottom": 0.41,
        "default_font_size": 7,
        "font_name": "MSゴシック",
        "position_tolerance_cells": "1〜2",
    },
    # A3 2pt: 印刷総 A3縦 52列×65行 / A3横 73列×46行（列幅2.00表示）
    # ※A3縦は手元PDFがないため A3横 値からアスペクト比 (297/420, 420/297) で予測
    "2pt_a3": {
        "col_width_mm": "6.18",
        "row_height_mm": "6.44",
        # A3縦 (印刷総 52×65 → コンテンツ 51×64)
        "max_cols": 51,
        "max_rows": 64,
        # A3横 (印刷総 73×46 → コンテンツ 72×45)
        "max_cols_landscape": 72,
        "max_rows_landscape": 45,
        "excel_col_width": 2.74,
        "excel_row_height": 18.25,
        "margin_left": 0.43,
        "margin_right": 0.43,
        "margin_top": 0.41,
        "margin_bottom": 0.41,
        "default_font_size": 6,
        "font_name": "MSゴシック",
        "position_tolerance_cells": "1",
    },
}


# ビジョンLLM（画像入力対応AIチャット）向けの視覚的検証プロンプト。
# 入力: PDF原本に現プレビュー罫線を半透明赤で重ねた差分画像 1枚 + ID 付き罫線リスト JSON。
# LLM は赤線が PDF の黒線と一致するか判定し、不要な赤線の ID と不足分の罫線を返す。
VISUAL_REVIEW_PROMPT = """\
PDF原本に現プレビューの罫線を **半透明の赤線** で重ねた **差分画像** と、
画像中の各赤線に振られた **ID 付きの罫線リスト (JSON)** を渡します。

赤線が PDF の黒い罫線と一致するか確認し、過不足を報告してください。

## 入力

- 差分画像（ページ {page_number}）: PDF 原本 + 赤色の現プレビュー罫線（数字ラベルは罫線 ID）
- 罫線リスト JSON: 画像内の全赤線が `id`, `type` (H/V), 位置情報付きで列挙されている
- グリッド: {max_rows} 行 × {max_cols} 列（1マス = {col_width_mm}mm × {row_height_mm}mm）

## 座標の規約（重要）

- `H, row=N` = **N 行目の上辺** にある水平線（= N-1 行目の下辺）。最上端は `row=1`
- `V, col=N` = **N 列目の左辺** にある垂直線。最左端は `col=1`
- `col_end` / `row_end` は **排他的境界**（= 最終セル + 1）。col 3〜12 の H 線なら `col_start=3, col_end=13`
- 左上セルが `(row=1, col=1)`、右・下方向に増加

## 判定基準

各赤線（ID 付き）を順に確認:

| 状況 | アクション |
|------|----------|
| 赤線の位置に対応する黒線が PDF にある | 何もしない |
| 赤線の位置に対応する黒線が **無い** | `remove_edges` で ID を削除 |
| PDF に黒線があるが対応する赤線が **無い** | `add_edge` で新規エッジを追加 |

**無視するもの**: テキスト・文字、PDF の薄い飾り線・影など。

## 座標の範囲制約

有効範囲は **row: 1〜{content_max_row}、col: 1〜{content_max_col}** です。範囲外を指定しないでください。

## 出力形式

差異がない場合は `{{"corrections": []}}` のみ出力。

```json
{{
  "corrections": [
    {{"action": "remove_edges", "page": {page_number}, "ids": [3, 17, 42]}},
    {{"action": "add_edge", "page": {page_number}, "type": "H", "row": 5, "col_start": 3, "col_end": 13}},
    {{"action": "add_edge", "page": {page_number}, "type": "V", "col": 3, "row_start": 1, "row_end": 9}}
  ]
}}
```

### フィールド名の厳守

- H エッジ: `type="H"`, `row`, `col_start`, `col_end`
- V エッジ: `type="V"`, `col`, `row_start`, `row_end`
- 削除: `ids` は配列（単一でも `[7]` の形式）

【最重要】出力は JSON のみ。説明・前置き・コードブロック記号（```）は不要。
【最重要】`col_end` / `row_end` は **排他的境界**（最終セル + 1）であること。"""
