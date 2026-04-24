# グリッドシステム

コンテンツ境界ベースのグリッド計算の仕組み、座標変換ロジック、GRID_SIZES 定数について説明します。

---

## 概要

Sheetling は PDF 上のコンテンツが存在する領域（コンテンツ境界）を基準にグリッドを計算します。ページ全体ではなくコンテンツ領域だけを `max_cols × max_rows` のグリッドに分割するため、余白の多い PDF でも Excel セルを効率的に使えます。

---

## コンテンツ境界の検出

`compute_grid_coords()`（`src/core/grid.py`） はまず全要素の座標からコンテンツ境界を検出します：

1. 全 words の `x0`, `x1`, `top`, `bottom` を収集
2. 全 rects の座標を収集
3. 全 table_cells の座標を収集
4. これらの最小値・最大値から `min_x`, `max_x`, `min_y`, `max_y` を決定

```
┌─────────────────────────── PDF ページ ───────────────────────────┐
│                                                                  │
│     (min_x, min_y)                                               │
│         ┌──────────── コンテンツ領域 ────────────┐                │
│         │  テキスト・テーブル・矩形が存在する範囲  │                │
│         │                                       │                │
│         └───────────────────────────────────────┘                │
│                                        (max_x, max_y)            │
│                                                                  │
└──────────────────────────────────────────────────────────────────┘
```

---

## グリッドサイズの計算

コンテンツ領域の幅・高さを `max_cols` / `max_rows` で均等分割します：

```
content_w = max_x - min_x
content_h = max_y - min_y
grid_w = content_w / max_cols   （1列あたりの幅 / PDF pt単位）
grid_h = content_h / max_rows   （1行あたりの高さ / PDF pt単位）
```

---

## 座標変換関数 `to_row()` / `to_col()`

PDF 座標をグリッドの行番号・列番号に変換します。左上が `(row=1, col=1)` です。

```python
def to_row(y):
    return max(1, min(max_rows, 1 + int((y - min_y) / grid_h)))

def to_col(x):
    return max(1, min(max_cols, 1 + int((x - min_x) / grid_w)))
```

- `y - min_y`: コンテンツ境界からの相対位置
- `/ grid_h`: 何行目に相当するか計算
- `1 +`: 1始まりのインデックス
- `max(1, min(max_rows, ...))`: 範囲内にクランプ

`to_col()` も同様のロジックで列番号に変換します。

### 変換例

```
コンテンツ領域: min_x=50, min_y=80, content_w=500, content_h=700
max_cols=61, max_rows=42 の場合:
  grid_w = 500 / 61 ≈ 8.20pt
  grid_h = 700 / 42 ≈ 16.67pt

PDF座標 (x=130, y=180) の場合:
  col = 1 + int((130 - 50) / 8.20) = 1 + int(9.76) = 10
  row = 1 + int((180 - 80) / 16.67) = 1 + int(6.00) = 7
  → グリッド座標: (row=7, col=10)
```

---

## 要素への座標付与

`compute_grid_coords()`（`src/core/grid.py`） は変換後の座標を各要素に直接付与します：

| 要素 | 付与されるフィールド |
|------|-------------------|
| words | `_row`, `_col`（+ 縦書きの場合 `_end_row`） |
| rects | `_row`, `_end_row`, `_col`, `_end_col` |
| table_cells → table_border_rects | `_row`, `_end_row`, `_col`, `_end_col`, `_borders` |

### table_border_rects の生成

テーブルの `table_cells` 2D配列からセルごとに border_rect を生成します：

- `None` のセル（結合セルの延長部分）は罫線を生成しない → 結合セル内部の不要な罫線が自動除去
- 各セルの `borders` はデフォルトで全辺 `true`

---

## GRID_SIZES 定数

`src/core/grid_config.py` で定義。用紙サイズ × グリッド密度の組み合わせ：

### 定数一覧

| キー | 用紙 | 密度 | 縦 (cols × rows) | 横 (cols × rows) |
|------|------|------|-------------------|-------------------|
| `1pt` | A4 | 高密度 | 47 × 39 | 70 × 25 |
| `2pt` | A4 | 中密度 | 29 × 39 | 44 × 25 |
| `1pt_a3` | A3 | 高密度 | 70 × 57 | 104 × 39 |
| `2pt_a3` | A3 | 中密度 | 44 × 57 | 65 × 39 |

### 各定数のフィールド

```python
{
    "col_width_mm": "3.48",        # Excel 列幅（mm）
    "row_height_mm": "6.44",       # Excel 行高（mm）
    "max_cols": 61,                # 縦向き最大列数
    "max_rows": 42,                # 縦向き最大行数
    "max_cols_landscape": 89,      # 横向き最大列数
    "max_rows_landscape": 30,      # 横向き最大行数
    "excel_col_width": 1.00,       # openpyxl 列幅値
    "excel_row_height": 18.25,     # openpyxl 行高値（pt）
    "margin_left": 0.2,            # 印刷マージン（インチ）
    "margin_right": 0.2,
    "margin_top": 0.3,
    "margin_bottom": 0.3,
    "default_font_size": 7,        # デフォルトフォントサイズ（pt）
    "font_name": "MS Gothic",      # フォント名
    "position_tolerance_cells": "1",  # 許容誤差（セル単位）
}
```

---

## 用紙サイズ・向きの自動検出

`setup_grid_params()`（`src/core/grid.py`） がページの寸法から判定します：

| 判定 | 条件 |
|------|------|
| A3 | `max(page_w, page_h) > 1000pt` |
| A4 | 上記以外 |
| 横向き | `page_w > page_h` |
| 縦向き | 上記以外 |

判定結果に応じて `GRID_SIZES` のキーを選択（例: A3横向き × 1pt → `1pt_a3` の `max_cols_landscape` / `max_rows_landscape` を使用）。

---

## コンテンツ境界情報の保持

`compute_grid_coords()`（`src/core/grid.py`） は計算したコンテンツ境界情報をページオブジェクトに保存します：

```python
page['_content_min_x'] = min_x
page['_content_min_y'] = min_y
page['_content_grid_w'] = grid_w
page['_content_grid_h'] = grid_h
```

これらは後段の `generate_border_preview()`（`src/renderer/preview.py`） で PDF 画像上の正しい位置に罫線を描画するために使用されます。
