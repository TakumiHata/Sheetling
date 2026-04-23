"""Sheetling 全体で共有する数値定数。

複数モジュールで重複していた、または文脈が分かりにくい数値を集約する。
tolerance（座標マッチ許容誤差）は PDF point 単位。
"""

# Excel レンダリング
COL_OFFSET = 1    # 印刷範囲の左に空ける列数
ROW_PADDING = 1   # ページ間に挿入する空行数

# 座標マッチ許容誤差（PDF point）
RECT_CONTAINMENT_TOL = 1.0   # pdf_extractor: 重複 rect の包含判定
WORD_IN_BBOX_TOL = 2.0       # layout: bbox 内の単語判定
WORD_IN_TABLE_TOL = 2.0      # layout: テーブル内外の単語判定
RECT_IN_TABLE_TOL = 3.0      # grid: テーブル bbox 内の rect 判定
EDGE_MERGE_GAP_TOL = 2.0     # pdf_extractor: エッジ連結許容

# 罫線・ライン抽出
THIN_LINE_THICKNESS = 3.0    # grid: この値未満の厚みは「線」扱い
LINE_RECT_MERGE_TOL = 5.0    # grid: 4線→矩形統合の誤差
LINE_HV_MIN_LENGTH = 2.0     # pdf_extractor: 縦横判定の最小長

# 線幅 → 罫線スタイル閾値
BORDER_THIN_MAX_LW = 1.0
BORDER_MEDIUM_MAX_LW = 2.0

# テキスト分割
VISUAL_LINE_GAP = 3.0         # layout: 視覚行の top 差分閾値
HORIZONTAL_GAP_FACTOR = 2.0   # text: 水平ギャップの font_size 倍率

# ページ全体を占有する rect/table の除外閾値（面積比）
MAX_TABLE_AREA_RATIO = 0.80
MAX_RECT_AREA_RATIO = 0.80
MAX_EDGE_RECT_AREA_RATIO = 0.85
