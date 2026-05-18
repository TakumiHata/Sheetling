"""
Sheetling 用の LLM プロンプト定義集。

  VISUAL_PHASE1_PROMPT — フェーズ1: 余分な罫線の削除 ID を判定させる
  VISUAL_PHASE2_PROMPT — フェーズ2: 不足している罫線の追加座標を判定させる（inclusive 座標）
"""


# フェーズ1: 削除判定のみ。追加は不要。
# 入力: PDF原本に現プレビュー罫線を半透明赤で重ねた差分画像 + ID 付き罫線リスト JSON。
# LLM は赤線ごとに「PDF の黒線と一致するか」を Yes/No で判定し、一致しない ID だけを返す。
VISUAL_PHASE1_PROMPT = """\
PDF原本に現プレビューの罫線を **半透明の赤線** で重ねた **差分画像** と、
各赤線に振られた **ID 付きの罫線リスト (JSON)** を渡します。

【フェーズ1: 削除判定のみ】追加の指示は不要です。

## 入力

- 差分画像（ページ {page_number}）: PDF 原本 + 赤色の現プレビュー罫線（数字ラベルは罫線 ID）
- 罫線リスト JSON: 全赤線が `id`, `type`(H/V), 位置情報付きで列挙
- グリッド: {max_rows} 行 × {max_cols} 列（1マス = {col_width_mm}mm × {row_height_mm}mm）

## 座標の読み方

- `H, row=N` : N 行目の上辺にある水平線
- `V, col=N` : N 列目の左辺にある垂直線
- `col_start` / `col_end`, `row_start` / `row_end` : 両端のセル番号（**両端 inclusive**）
  例: col 3〜12 の H 線 → col_start=3, col_end=12

## 判定基準

各赤線（ID 付き）を順に確認:

| 状況 | アクション |
|------|----------|
| 赤線の位置に対応する黒線が PDF にある | **何もしない** |
| 赤線の位置に対応する黒線が PDF に **ない** | `ids` に追加して削除 |

**無視するもの**: テキスト・文字、PDF の薄い飾り線・影。

## 出力形式

削除対象がない場合は `{{"corrections": []}}` のみ出力。

```json
{{"corrections": [
  {{"action": "remove_edges", "page": {page_number}, "ids": [3, 17, 42]}}
]}}
```

【最重要】出力は JSON のみ。説明・前置き・コードブロック記号（```）は不要。"""


# フェーズ2: 追加判定のみ。削除は不要。inclusive 座標を使う。
# 入力: 同じ差分画像（フェーズ1と共通）。
# LLM は PDF に黒線があるが赤線がない箇所を特定し、inclusive 座標で add_edge を返す。
VISUAL_PHASE2_PROMPT = """\
PDF原本に現プレビューの罫線を **半透明の赤線** で重ねた **差分画像** を渡します。

【フェーズ2: 追加判定のみ】削除の指示は不要です。

## 入力

- 差分画像（ページ {page_number}）: PDF 原本 + 赤色の現プレビュー罫線
- グリッド: {max_rows} 行 × {max_cols} 列（1マス = {col_width_mm}mm × {row_height_mm}mm）

## 座標の規約（重要）

- `H, row=N` : N 行目の上辺にある水平線（= N-1 行目の下辺）。最上端は `row=1`
- `V, col=N` : N 列目の左辺にある垂直線。最左端は `col=1`
- `col_start` / `col_end` : 水平線の左端・右端セル番号（**両端 inclusive**）
  例: col 3〜12 の H 線 → `col_start=3, col_end=12`（13 **ではない**）
- `row_start` / `row_end` : 垂直線の上端・下端セル番号（**両端 inclusive**）
  例: row 1〜8 の V 線 → `row_start=1, row_end=8`（9 **ではない**）
- 左上セルが `(row=1, col=1)`、右・下方向に増加

## 判定基準

PDF の黒線を確認し:

| 状況 | アクション |
|------|----------|
| 黒線に対応する赤線がすでにある | **何もしない** |
| 黒線に対応する赤線が **ない** | `add_edge` で追加 |

**無視するもの**: テキスト・文字、PDF の薄い飾り線・影。

## 座標の範囲制約

有効範囲は **row: 1〜{content_max_row}、col: 1〜{content_max_col}** です。範囲外を指定しないでください。

## 出力形式

追加がない場合は `{{"corrections": []}}` のみ出力。

```json
{{"corrections": [
  {{"action": "add_edge", "page": {page_number}, "type": "H", "row": 5, "col_start": 3, "col_end": 12}},
  {{"action": "add_edge", "page": {page_number}, "type": "V", "col": 3, "row_start": 1, "row_end": 8}}
]}}
```

### フィールド名の厳守

- H エッジ: `type="H"`, `row`, `col_start`, `col_end`（inclusive）
- V エッジ: `type="V"`, `col`, `row_start`, `row_end`（inclusive）

【最重要】出力は JSON のみ。説明・前置き・コードブロック記号（```）は不要。
【最重要】`col_end` / `row_end` は **最後のセルの番号**（inclusive）。col 3〜12 なら `col_end=12`（13 ではない）。"""
