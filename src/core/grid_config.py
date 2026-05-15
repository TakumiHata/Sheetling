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
# 例: A4横 1pt → 印刷総 87列×32行 → max_cols=86, max_rows=31
#
GRID_SIZES = {
    # =========================================================================
    # A4 (595pt × 842pt)
    # =========================================================================
    # Sheetling "1pt": 印刷総 A4縦 59列×45行 / A4横 86列×31行（列幅1.00表示）
    "1pt": {
        "col_width_mm": "3.48",
        "row_height_mm": "6.44",
        # A4縦 (印刷総 59×45 → コンテンツ 58×44)
        "max_cols": 58,
        "max_rows": 44,
        # A4横 (印刷総 86×31 → コンテンツ 85×30)
        "max_cols_landscape": 85,
        "max_rows_landscape": 30,
        "excel_col_width": 1.65,  # W = 表示値 + 0.65 (実測: 1.74→1.09表示のためオフセット0.65に修正)
        "excel_row_height": 18.25,
        "margin_left": 0.43,
        "margin_right": 0.43,
        "margin_top": 0.41,
        "margin_bottom": 0.41,
        "default_font_size": 7,
        "font_name": "MS 明朝",
        "position_tolerance_cells": "1〜2",
    },
    # Sheetling "2pt": 印刷総 A4縦 36列×45行 / A4横 53列×31行（列幅2.00表示）
    "2pt": {
        "col_width_mm": "6.18",
        "row_height_mm": "6.44",
        # A4縦 (印刷総 36×45 → コンテンツ 35×44)
        "max_cols": 35,
        "max_rows": 44,
        # A4横 (印刷総 53×31 → コンテンツ 52×30)
        "max_cols_landscape": 52,
        "max_rows_landscape": 30,
        "excel_col_width": 2.65,  # W = 表示値 + 0.65 (実測: 2.74→2.09表示のためオフセット0.65に修正)
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
    # A3 1pt: 印刷総 A3縦 89列×65行 / A3横 126列×46行（列幅1.00表示）
    # A3横は実測値。A3縦は A3横の列密度から比例換算: 126*(842/1190)=89.2→89列。
    # ※A3縦 PDF での最終確認は未実施。実機確認後に要調整。
    "1pt_a3": {
        "col_width_mm": "3.48",
        "row_height_mm": "6.44",
        # A3縦 (印刷総 89×65 → コンテンツ 88×64)
        "max_cols": 88,
        "max_rows": 64,
        # A3横 (印刷総 126×46 → コンテンツ 125×45)
        "max_cols_landscape": 125,
        "max_rows_landscape": 45,
        "excel_col_width": 1.65,
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
    # A3横は実測値。A3縦は A3横の列密度から比例換算: 73*(842/1190)=51.6→52列。
    # ※A3縦 PDF での最終確認は未実施。実機確認後に要調整。
    "2pt_a3": {
        "col_width_mm": "6.18",
        "row_height_mm": "6.44",
        # A3縦 (印刷総 52×65 → コンテンツ 51×64)
        "max_cols": 51,
        "max_rows": 64,
        # A3横 (印刷総 73×46 → コンテンツ 72×45)
        "max_cols_landscape": 72,
        "max_rows_landscape": 45,
        "excel_col_width": 2.65,
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
