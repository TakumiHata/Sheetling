"""グリッドサイズ別の共有設定値。

用紙サイズ（A4/A3）・グリッド密度（1pt/2pt）ごとのセル寸法・Excel 列幅・
マージン・フォント等をまとめた辞書。レンダリング（`renderer/excel.py`）、
グリッド計算（`core/grid.py`）、プロンプト整形（`core/auto_layout_service.py`）
から共通に参照される。
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
        "font_name": "MS 明朝",
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
        "font_name": "MS 明朝",
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
        "font_name": "MS 明朝",
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
        "font_name": "MS 明朝",
        "position_tolerance_cells": "1",
    },
}
