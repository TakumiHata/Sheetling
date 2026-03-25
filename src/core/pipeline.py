"""
Sheetling パイプライン。

【全自動モード】
  auto: PDF解析 → レイアウトJSON自動生成 → _gen.py 生成 → Excel描画
  correct: ビジョンLLM修正指示を適用して Excel を再生成

【LLM協業モード（高精度）】
  extract: PDF解析 → STEP1/STEP1.5 プロンプト生成 → PDFページ画像出力
  fill:    STEP1.5 LLM出力 → テキスト補完 → layout.json 保存
  review:  layout.json → 罫線プレビュー画像 + VISUAL_REVIEW_PROMPT 生成
  generate: layout.json → Excel 直接描画（corrections 適用後）
"""

import json
import re
from collections import defaultdict
from pathlib import Path

from src.parser.pdf_extractor import extract_pdf_data
from src.templates.prompts import (
    CODE_ERROR_FIXING_PROMPT, GEN_CODE_TEMPLATE,
    TABLE_ANCHOR_PROMPT, LAYOUT_REVIEW_PROMPT,
    VISUAL_REVIEW_PROMPT, GRID_SIZES,
)
from src.utils.logger import get_logger


# MS Access / Windows 環境でよく使われるフォント名の正規化テーブル。
# PDF 埋め込み名はハイフン・スペース・大小文字が揺れるため、
# Excel が認識できる表記に統一する。
_FONT_ALIASES: dict = {
    'MS Gothic':     'MS Gothic',
    'MSGothic':      'MS Gothic',
    'MS PGothic':    'MS PGothic',
    'MSPGothic':     'MS PGothic',
    'MS Mincho':     'MS Mincho',
    'MSMincho':      'MS Mincho',
    'MS PMincho':    'MS PMincho',
    'MSPMincho':     'MS PMincho',
    'MS UI Gothic':  'MS UI Gothic',
    'MSUIGothic':    'MS UI Gothic',
    'Meiryo':        'Meiryo',
    'Meiryo UI':     'Meiryo UI',
    'MeiryoUI':      'Meiryo UI',
    'Yu Gothic':     'Yu Gothic',
    'YuGothic':      'Yu Gothic',
    'Yu Mincho':     'Yu Mincho',
    'YuMincho':      'Yu Mincho',
}


def _normalize_font_name(raw_name):
    """PDF フォント名を Excel に渡せる形式に整形する。
    1. bytes / bytes repr 文字列（pdfminer が返す非ASCII名）を除外
    2. サブセットプレフィックスを除去 (例: "ABCDEF+MS-Gothic" → "MS-Gothic")
    3. ハイフン区切りをスペースに変換 (例: "MS-Gothic" → "MS Gothic")
    4. _FONT_ALIASES でエイリアス解決（揺れ表記を正規表記に統一）
    Excel が認識できないフォント名はデフォルトフォントにフォールバックされる。"""
    if not raw_name:
        return None
    # bytes オブジェクトは使用不可
    if isinstance(raw_name, bytes):
        return None
    # pdfminer が bytes を str() した "b'...'" 形式の文字列も使用不可
    if isinstance(raw_name, str) and raw_name.startswith("b'"):
        return None
    # サブセットプレフィックスを除去 (例: "ABCDEF+MS-Gothic" → "MS-Gothic")
    name = re.sub(r'^[A-Z]{6}\+', '', raw_name)
    # ハイフン区切りをスペースに変換 (例: "MS-Gothic" → "MS Gothic")
    name = name.replace('-', ' ').strip()
    # エイリアステーブルで正規化（大文字小文字も考慮して前方一致）
    return _FONT_ALIASES.get(name, name) or None


def _sanitize_generated_code(code: str) -> tuple[str, list[str]]:
    """生成コードの既知の問題パターンを検出・自動修正する。"""
    fixes = []

    # 修正: ws.page_margins = {...} → 属性代入形式に変換
    margins_dict_pattern = re.compile(r"ws\.page_margins\s*=\s*\{([^}]*)\}", re.DOTALL)
    match = margins_dict_pattern.search(code)
    if match:
        kv_pattern = re.compile(r"['\"](\w+)['\"]\s*:\s*([\d.]+)")
        pairs = kv_pattern.findall(match.group(1))
        if pairs:
            replacement = "\n".join(f"ws.page_margins.{k} = {v}" for k, v in pairs)
            code = margins_dict_pattern.sub(replacement, code)
            fixes.append("ws.page_margins への dict 代入を属性代入形式に自動修正しました")

    return code, fixes


def _compute_grid_coords(page: dict, max_rows: int, max_cols: int) -> None:
    """
    PDF座標をExcel行・列番号に変換し、各要素にインプレースで付与する。
    Y・X座標ともにクラスタリングを行い、近接する座標を同一行・列に統一する。
    """
    page_height = page['height']
    page_width = page['width']
    grid_h = page_height / max_rows
    grid_w = page_width / max_cols

    def snap(v: float) -> float:
        return round(float(v), 2)

    def build_cluster_map(raw_vals: set, grid_size: float, max_idx: int, anchor_vals: set = None) -> dict:
        """
        近接する座標値をクラスタリングしてグリッドインデックスに変換する。
        anchor_vals に含まれる値は直前のクラスタと近接していても必ず独立したクラスタを開始する。
        これによりテーブル列境界が隣接列と合算されるのを防ぐ。

        比較基準: clusters[-1][-1]（クラスタ末尾値）を使用する。
        旧実装の clusters[-1][0]（先頭値）では、クラスタが拡張するにつれて
        先頭から遠い値まで取り込み続けてしまい、密集した行が誤合算される原因になっていた。
        しきい値も 0.5 → 0.35 に絞り、近接しすぎる行の分離精度を向上させる。
        """
        anchor_vals = anchor_vals or set()
        sorted_vals = sorted(raw_vals)
        clusters: list = []
        for v in sorted_vals:
            # [修正] clusters[-1][0](先頭) → clusters[-1][-1](末尾) との距離で判定し、
            # しきい値を 0.5 → 0.35 に縮小。
            if not clusters or v - clusters[-1][-1] > grid_size * 0.35 or v in anchor_vals:
                clusters.append([v])
            else:
                clusters[-1].append(v)
        val_map = {}
        for cluster in clusters:
            centroid = sum(cluster) / len(cluster)
            idx = max(1, min(max_idx, 1 + int(centroid / grid_size)))
            for v in cluster:
                val_map[v] = idx
        return val_map

    # 全Y・X座標を収集
    y_vals: set = set()
    x_vals: set = set()
    # テーブル列境界X座標（クラスタリング時に独立扱いにするため別途保持）
    table_col_x_anchors: set = set()
    # テーブル行境界Y座標（クラスタリング時に独立扱いにするため別途保持）
    table_row_y_anchors: set = set()

    for w in page['words']:
        y_vals.add(snap(w['top']))
        x_vals.add(snap(w['x0']))
        x_vals.add(snap(w['x1']))
        if 'bottom' in w:
            # [修正] 縦文字に限らず全ワードの bottom を Y 座標セットに追加。
            # これにより行間ギャップがクラスタ境界として機能し、
            # 密集した行が誤合算されるのを防ぐ。
            y_vals.add(snap(w['bottom']))

    # [修正追加] 隣接する水平ワード間に有意な垂直ギャップがある場合、
    # 次の行の top をアンカーとして登録し、クラスタリングで行が合算されないよう強制分離する。
    _h_words_with_bottom = [
        w for w in page['words']
        if not w.get('is_vertical') and 'bottom' in w
    ]
    _h_words_sorted = sorted(_h_words_with_bottom, key=lambda w: float(w['top']))
    _GAP_ANCHOR_THRESHOLD = 2.0  # pt: これを超える bottom→top ギャップで次行をアンカー化
    for _i in range(len(_h_words_sorted) - 1):
        _cur  = _h_words_sorted[_i]
        _next = _h_words_sorted[_i + 1]
        _cur_bottom  = float(_cur['bottom'])
        _next_top    = float(_next['top'])
        if _next_top - _cur_bottom > _GAP_ANCHOR_THRESHOLD:
            _sy = snap(_next_top)
            y_vals.add(_sy)
            table_row_y_anchors.add(_sy)
    for r in page['rects']:
        y_vals.add(snap(r['top']))
        y_vals.add(snap(r['bottom']))
        x_vals.add(snap(r['x0']))
        x_vals.add(snap(r['x1']))
    for bbox in page['table_bboxes']:
        y_vals.add(snap(bbox[1]))  # top
        y_vals.add(snap(bbox[3]))  # bottom
    for col_xs in page['table_col_x_positions']:
        for x in col_xs:
            sx = snap(x)
            x_vals.add(sx)
            table_col_x_anchors.add(sx)
    for row_ys in page.get('table_row_y_positions', []):
        for y in row_ys:
            sy = snap(y)
            y_vals.add(sy)
            table_row_y_anchors.add(sy)
    # テーブル内の水平エッジをアンカーとして補完。
    # pdfplumber がテーブル行境界を見逃した場合（セル幅が edge_min_length 未満等）、
    # h_edges の水平線から内部行境界を復元する。
    _h_edge_anchor_tol_x = 5.0  # エッジがテーブル幅に対してどれだけ短くてもよいか(pt)
    _h_edge_anchor_tol_y = 1.0  # テーブル上下端からの除外マージン(pt)
    for bbox in page.get('table_bboxes', []):
        bx0, by0, bx1, by1 = bbox
        for edge in page.get('h_edges', []):
            ey = snap(edge['y'])
            # テーブル内部のy座標（上下端は除外）
            if not (by0 + _h_edge_anchor_tol_y < ey < by1 - _h_edge_anchor_tol_y):
                continue
            # エッジのx範囲がテーブルのx範囲と有意に重複している
            if edge['x0'] > bx1 - _h_edge_anchor_tol_x or edge['x1'] < bx0 + _h_edge_anchor_tol_x:
                continue
            y_vals.add(ey)
            table_row_y_anchors.add(ey)
    for cells in page.get('table_cells', []):
        for c in cells:
            y_vals.add(snap(c['top']))
            y_vals.add(snap(c['bottom']))
            x_vals.add(snap(c['x0']))
            x_vals.add(snap(c['x1']))
    # エッジ座標もクラスタリングに含める（罫線位置をグリッドに正確に反映）
    for edge in page.get('h_edges', []):
        y_vals.add(snap(edge['y']))
        x_vals.add(snap(edge['x0']))
        x_vals.add(snap(edge['x1']))
    for edge in page.get('v_edges', []):
        x_vals.add(snap(edge['x']))
        y_vals.add(snap(edge['y0']))
        y_vals.add(snap(edge['y1']))

    y_map = build_cluster_map(y_vals, grid_h, max_rows, anchor_vals=table_row_y_anchors)
    x_map = build_cluster_map(x_vals, grid_w, max_cols, anchor_vals=table_col_x_anchors)

    # テーブル列境界が同一グリッド列に潰れた場合の後処理:
    # 各テーブルの列X座標を左から順に走査し、前の列と同じグリッド列になっていたら +1 する。
    for col_xs in page['table_col_x_positions']:
        snapped_xs = sorted(set(snap(x) for x in col_xs))
        prev_idx = 0
        for x in snapped_xs:
            idx = x_map[x]
            if idx <= prev_idx:
                idx = prev_idx + 1
            idx = min(idx, max_cols)
            x_map[x] = idx
            prev_idx = idx

    # テーブル行境界が同一グリッド行に潰れた場合の後処理:
    # 各テーブルの行Y座標を上から順に走査し、前の行と同じグリッド行になっていたら +1 する。
    for row_ys in page.get('table_row_y_positions', []):
        snapped_ys = sorted(set(snap(y) for y in row_ys))
        prev_idx = 0
        for y in snapped_ys:
            idx = y_map[y]
            if idx <= prev_idx:
                idx = prev_idx + 1
            idx = min(idx, max_rows)
            y_map[y] = idx
            prev_idx = idx

    # テーブル底辺直下の注釈要素がテーブル底辺と同一グリッド行に落ちる場合の補正。
    # グリッド解像度が低い（行高＞底辺〜注釈上端の間隔）ときに発生する位置重複を解消する。
    # テーブル境界 y 値（table_row_y_anchors）は除外し、注釈等のコンテンツ y 値のみ対象とする。
    for row_ys_list in page.get('table_row_y_positions', []):
        if not row_ys_list:
            continue
        table_bottom_y = snap(max(row_ys_list))
        if table_bottom_y not in y_map:
            continue
        table_end_row = y_map[table_bottom_y]
        # テーブル底辺直下（1グリッド行以内）に非テーブル境界の y 値が同一行に落ちていれば衝突
        has_collision = any(
            table_bottom_y < yv <= table_bottom_y + grid_h
            and y_map[yv] == table_end_row
            and yv not in table_row_y_anchors
            for yv in y_map
        )
        if not has_collision:
            continue
        # 衝突あり: テーブル底辺より下の全 y 値を 1 行ずらして空行を挿入
        for yv in list(y_map.keys()):
            if yv > table_bottom_y:
                y_map[yv] = min(y_map[yv] + 1, max_rows)

    # words に付与
    for w in page['words']:
        w['_row'] = y_map[snap(w['top'])]
        w['_col'] = x_map[snap(w['x0'])]
        if w.get('is_vertical') and 'bottom' in w:
            sv = snap(w['bottom'])
            w['_end_row'] = y_map.get(sv, w['_row'])

    # rects に付与
    for r in page['rects']:
        r['_row'] = y_map[snap(r['top'])]
        r['_end_row'] = y_map[snap(r['bottom'])]
        r['_col'] = x_map[snap(r['x0'])]
        r['_end_col'] = x_map[snap(r['x1'])]

    # テーブル内に含まれる rects を除外（table_border_rects で代替するため）
    # [修正] tol 1.0 → 3.0pt: テーブル外周ぎりぎりに配置された rect が
    # 残存して table_border_rects と重複描画される問題を防ぐ。
    tol = 3.0
    table_bboxes = page.get('table_bboxes', [])

    def is_inside_table(r: dict) -> bool:
        for bbox in table_bboxes:
            if (r['x0'] >= bbox[0] - tol and r['x1'] <= bbox[2] + tol and
                    r['top'] >= bbox[1] - tol and r['bottom'] <= bbox[3] + tol):
                return True
        return False

    page['rects'] = [r for r in page['rects'] if not is_inside_table(r)]

    # テーブルの列・行グリッドから border_rect を生成（pdfplumber が検出した列数×行数）
    table_border_rects = []
    for col_xs, row_ys in zip(page.get('table_col_x_positions', []),
                               page.get('table_row_y_positions', [])):
        col_xs_s = sorted(set(snap(x) for x in col_xs))
        row_ys_s = sorted(set(snap(y) for y in row_ys))
        n_cols = len(col_xs_s) - 1
        n_rows = len(row_ys_s) - 1
        for ri in range(n_rows):
            for ci in range(n_cols):
                table_border_rects.append({
                    '_row':        y_map.get(row_ys_s[ri], 1),
                    '_end_row':    y_map.get(row_ys_s[ri + 1], 1),
                    '_col':        x_map.get(col_xs_s[ci], 1),
                    '_end_col':    x_map.get(col_xs_s[ci + 1], 1),
                    '_pdf_x0':     col_xs_s[ci],
                    '_pdf_top':    row_ys_s[ri],
                    '_pdf_x1':     col_xs_s[ci + 1],
                    '_pdf_bottom': row_ys_s[ri + 1],
                    # テーブル外周フラグ（描画時の垂直罫線抑制で使用）
                    '_outer_left':   ci == 0,
                    '_outer_right':  ci == n_cols - 1,
                })
    page['table_border_rects'] = table_border_rects

    # ---- エッジから辺ごとの罫線有無を判定 ----------------------------------------

    def _nearest_idx(val: float, coord_map: dict) -> int:
        """valに最も近いcoord_mapのキーに対応するグリッドインデックスを返す。"""
        if not coord_map:
            return 1
        sv = snap(val)
        if sv in coord_map:
            return coord_map[sv]
        return coord_map[min(coord_map.keys(), key=lambda k: abs(k - sv))]

    # エッジをグリッド座標に変換してマップ化
    # h_edge_map: row_idx -> [(col_start, col_end), ...]
    # v_edge_map: col_idx -> [(row_start, row_end), ...]
    h_edge_map: dict = {}
    for edge in page.get('h_edges', []):
        ri = _nearest_idx(edge['y'], y_map)
        cs = _nearest_idx(edge['x0'], x_map)
        ce = _nearest_idx(edge['x1'], x_map)
        h_edge_map.setdefault(ri, []).append((min(cs, ce), max(cs, ce)))

    v_edge_map: dict = {}
    v_edge_max_span: dict = {}  # col_idx -> その列で検出された垂直エッジの最大スパン(pt)
    for edge in page.get('v_edges', []):
        ci = _nearest_idx(edge['x'], x_map)
        rs = _nearest_idx(edge['y0'], y_map)
        re = _nearest_idx(edge['y1'], y_map)
        v_edge_map.setdefault(ci, []).append((min(rs, re), max(rs, re)))
        span = edge.get('span', abs(edge['y1'] - edge['y0']))
        if span > v_edge_max_span.get(ci, 0):
            v_edge_max_span[ci] = span

    def _overlaps_h(edges: list, col_s: int, col_e: int) -> bool:
        """エッジリストの中に col_s〜col_e スパンと 30% 以上重複するものがあるか。"""
        span = col_e - col_s
        if span <= 0:
            return any(cs <= col_s <= ce for cs, ce in edges)
        for cs, ce in edges:
            overlap = min(ce, col_e) - max(cs, col_s)
            if overlap >= span * 0.3:
                return True
        return False

    def _overlaps_v(edges: list, row_s: int, row_e: int) -> bool:
        """エッジリストの中に row_s〜row_e スパンと 30% 以上重複するものがあるか。"""
        span = row_e - row_s
        if span <= 0:
            return any(rs <= row_s <= re for rs, re in edges)
        for rs, re in edges:
            overlap = min(re, row_e) - max(rs, row_s)
            if overlap >= span * 0.3:
                return True
        return False

    def _has_h(row: int, col_s: int, col_e: int) -> bool:
        """指定行に col_s〜col_e をカバーする水平エッジがあるか。
        グリッド量子化誤差で ±1 行ずれる場合があるため近傍行も検索する。"""
        for r in (row - 1, row, row + 1):
            if _overlaps_h(h_edge_map.get(r, []), col_s, col_e):
                return True
        return False

    def _has_v(col: int, row_s: int, row_e: int) -> bool:
        """指定列に row_s〜row_e をカバーする垂直エッジがあるか。
        グリッド量子化誤差で ±1 列ずれる場合があるため近傍列も検索する。"""
        for c in (col - 1, col, col + 1):
            if _overlaps_v(v_edge_map.get(c, []), row_s, row_e):
                return True
        return False

    # 主要垂直境界の閾値(pt): この高さ以上のエッジがある列は月区切り等の主要線とみなす。
    # 短いセル側辺（≈1行高≈20pt）は除外し、表高の30%超を占める線だけを採用する。
    _MAJOR_V_SPAN_THRESHOLD = page['height'] * 0.30

    # table_border_rects に _borders を付与
    for tbr in page['table_border_rects']:
        r, er, c, ec = tbr['_row'], tbr['_end_row'], tbr['_col'], tbr['_end_col']
        tbr['_borders'] = {
            'top':    _has_h(r,  c, ec),
            'bottom': _has_h(er, c, ec),
            'left':   _has_v(c,  r, er),
            'right':  _has_v(ec, r, er),
        }
        # 主要垂直境界フラグ: 長い縦線（月区切り等）が検出された列か
        tbr['_major_left']  = v_edge_max_span.get(c,  0) >= _MAJOR_V_SPAN_THRESHOLD
        tbr['_major_right'] = v_edge_max_span.get(ec, 0) >= _MAJOR_V_SPAN_THRESHOLD

    # rects にも _borders を付与（矩形枠の各辺）
    for rect in page['rects']:
        r, er = rect['_row'], rect['_end_row']
        c, ec = rect['_col'], rect['_end_col']
        if r == er and c != ec:
            # 水平分割線（高さ1グリッド未満の細長い矩形）:
            # top のみ描画。bottom は同じ行なので重複ライン、left/right は端キャップになるため除外。
            rect['_borders'] = {
                'top':    _has_h(r, c, ec),
                'bottom': False,
                'left':   False,
                'right':  False,
            }
        elif c == ec and r != er:
            # 垂直分割線（幅1グリッド未満の細長い矩形）:
            # left のみ描画。right は同じ列なので重複ライン、top/bottom は端キャップになるため除外。
            rect['_borders'] = {
                'top':    False,
                'bottom': False,
                'left':   _has_v(c, r, er),
                'right':  False,
            }
        else:
            rect['_borders'] = {
                'top':    _has_h(r,  c, ec),
                'bottom': _has_h(er, c, ec),
                'left':   _has_v(c,  r, er),
                'right':  _has_v(ec, r, er),
            }

    # ---- 隣接セル間の _borders 整合性パス ---------------------------------------
    # 量子化誤差で A.right と B.left（共有辺）が不一致になる場合を OR で統一する。
    tbrs = page['table_border_rects']
    # 水平方向: (row, end_row) -> {col -> tbr}
    h_band: dict = {}
    # 垂直方向: (col, end_col) -> {row -> tbr}
    v_band: dict = {}
    for tbr in tbrs:
        h_band.setdefault((tbr['_row'], tbr['_end_row']), {})[tbr['_col']] = tbr
        v_band.setdefault((tbr['_col'], tbr['_end_col']), {})[tbr['_row']] = tbr

    for tbr in tbrs:
        # 右隣: 同じ row/end_row 帯で _col == self._end_col
        right_neighbor = h_band.get((tbr['_row'], tbr['_end_row']), {}).get(tbr['_end_col'])
        if right_neighbor:
            merged = tbr['_borders']['right'] or right_neighbor['_borders']['left']
            tbr['_borders']['right'] = merged
            right_neighbor['_borders']['left'] = merged
        # 下隣: 同じ col/end_col 帯で _row == self._end_row
        bottom_neighbor = v_band.get((tbr['_col'], tbr['_end_col']), {}).get(tbr['_end_row'])
        if bottom_neighbor:
            merged = tbr['_borders']['bottom'] or bottom_neighbor['_borders']['top']
            tbr['_borders']['bottom'] = merged
            bottom_neighbor['_borders']['top'] = merged

    # ---------------------------------------------------------------------------------

    # ---- PDF 余白分の空き行を除去するため _row を正規化 --------------------------------
    # PDF の上部余白（margin_top）がグリッド行として現れ、コンテンツ前に数行の空白が生じる。
    # 全要素の最小 _row を求め、1 になるようにシフトして空き行を除去する。
    all_rows = (
        [w['_row'] for w in page['words'] if '_row' in w]
        + [r['_row'] for r in page['rects'] if '_row' in r]
        + [tbr['_row'] for tbr in page['table_border_rects']]
    )
    if all_rows:
        row_shift = min(all_rows) - 1
        if row_shift > 0:
            for w in page['words']:
                if '_row' in w:
                    w['_row'] -= row_shift
                    if '_end_row' in w:
                        w['_end_row'] -= row_shift
            for r in page['rects']:
                if '_row' in r:
                    r['_row'] -= row_shift
                    r['_end_row'] -= row_shift
            for tbr in page['table_border_rects']:
                tbr['_row'] -= row_shift
                tbr['_end_row'] -= row_shift

    # ---- 以下はグリッド座標計算には使用したが LLM には渡さない -------------------------
    page.pop('table_cells', None)
    page.pop('table_data', None)
    page.pop('table_row_y_positions', None)
    page.pop('h_edges', None)
    page.pop('v_edges', None)

logger = get_logger(__name__)


# ===========================================================================
# pre版から移植: LLM協業モード用ユーティリティ関数
# ===========================================================================

def _merge_table_border_rects(tbrs: list) -> list:
    """
    隣接セル間に境界線がない table_border_rects を統合し、結合セルを 1 つの
    border_rect として表現する。隣接整合性パスの後に呼ぶこと（共有辺の値が一致済み）。
    """
    # --- 縦方向マージ ---
    col_groups: dict = defaultdict(list)
    for tbr in tbrs:
        col_groups[(tbr['_col'], tbr['_end_col'])].append(tbr)

    v_merged: list = []
    for cells in col_groups.values():
        cells = sorted(cells, key=lambda c: c['_row'])
        stack = [dict(cells[0])]
        for cell in cells[1:]:
            prev = stack[-1]
            if prev['_end_row'] == cell['_row'] and not prev['_borders']['bottom']:
                prev['_end_row']           = cell['_end_row']
                prev['_pdf_bottom']        = cell['_pdf_bottom']
                prev['_borders']['bottom'] = cell['_borders']['bottom']
                prev['_borders']['left']   = prev['_borders']['left']  or cell['_borders']['left']
                prev['_borders']['right']  = prev['_borders']['right'] or cell['_borders']['right']
                prev['_outer_right']       = prev.get('_outer_right', False) or cell.get('_outer_right', False)
                prev['_major_right']       = prev.get('_major_right', False) or cell.get('_major_right', False)
            else:
                stack.append(dict(cell))
        v_merged.extend(stack)

    # --- 横方向マージ ---
    row_groups: dict = defaultdict(list)
    for tbr in v_merged:
        row_groups[(tbr['_row'], tbr['_end_row'])].append(tbr)

    h_merged: list = []
    for cells in row_groups.values():
        cells = sorted(cells, key=lambda c: c['_col'])
        stack = [dict(cells[0])]
        for cell in cells[1:]:
            prev = stack[-1]
            if prev['_end_col'] == cell['_col'] and not prev['_borders']['right']:
                prev['_end_col']           = cell['_end_col']
                prev['_pdf_x1']            = cell['_pdf_x1']
                prev['_borders']['right']  = cell['_borders']['right']
                prev['_borders']['top']    = prev['_borders']['top']    or cell['_borders']['top']
                prev['_borders']['bottom'] = prev['_borders']['bottom'] or cell['_borders']['bottom']
                prev['_outer_right']       = cell.get('_outer_right', False)
                prev['_major_right']       = cell.get('_major_right', False)
            else:
                stack.append(dict(cell))
        h_merged.extend(stack)

    return h_merged


def _fix_empty_cell_type_attr(xlsx_path: str) -> None:
    """
    openpyxl 3.1.x は値なしでスタイル（罫線）のみ設定されたセルに t="n" 属性を付与する。
    Excel Online はこれを不正な属性として修復処理（ブックが修復されました）を行い、
    罫線スタイルを除去してしまう。
    保存後に xlsx の ZIP 内 sheet XML を走査し、空セルの t="n" 属性を除去することで回避する。
    対象: <c r="..." s="数字" t="n" /> 形式の空セル（子要素なし・値なし）
    """
    import zipfile, shutil, tempfile
    pat = re.compile(r'(<c\s+r="[^"]+"\s+s="\d+"\s+)t="n"\s*(/>)')
    tmp = xlsx_path + '.tmp_fix'
    with zipfile.ZipFile(xlsx_path, 'r') as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith('xl/worksheets/') and item.filename.endswith('.xml'):
                text = data.decode('utf-8')
                text = pat.sub(r'\1\2', text)
                data = text.encode('utf-8')
            zout.writestr(item, data)
    shutil.move(tmp, xlsx_path)


def _render_layout_to_xlsx(layout: list, grid_params: dict, output_path: str) -> None:
    """
    レイアウト JSON を openpyxl で直接 Excel ファイルに描画する。
    LLM 生成コードを使わずにプログラム的に Excel を生成する。

    COL_OFFSET  = 1  （左1マス余白）
    ROW_PADDING = 1  （ページ上部1マス余白）
    row_offset for page N = (N-1) * max_rows + ROW_PADDING
    page break at page N  = N * (max_rows + ROW_PADDING)
    """
    from openpyxl import Workbook
    from openpyxl.styles import Border, Side, Alignment, Font
    from openpyxl.worksheet.pagebreak import Break
    from openpyxl.utils import get_column_letter

    COL_OFFSET = 1
    ROW_PADDING = 1
    max_rows = grid_params['max_rows']
    col_width = grid_params.get('excel_col_width', 1.45)
    row_height = grid_params.get('excel_row_height', 11.34)
    paper_size = grid_params.get('paper_size', 9)
    orientation = grid_params.get('orientation', 'portrait')
    default_font_size = grid_params.get('default_font_size', 7)
    font_name     = grid_params.get('font_name', 'MS Gothic')
    margin_left   = grid_params.get('margin_left',   0.43)
    margin_right  = grid_params.get('margin_right',  0.43)
    margin_top    = grid_params.get('margin_top',    0.41)
    margin_bottom = grid_params.get('margin_bottom', 0.41)

    total_pages = len(layout)

    wb = Workbook()
    ws = wb.active
    ws.sheet_format.defaultColWidth = col_width
    ws.sheet_format.defaultRowHeight = row_height
    ws.sheet_format.customHeight = True
    ws.sheet_view.showGridLines = True

    thin = Side(style='thin')

    def _set_border_side(row: int, col: int, **sides) -> None:
        if row < 1:
            return
        try:
            cell = ws.cell(row=row, column=col)
            ex = cell.border
            cell.border = Border(
                top=sides.get('top', ex.top),
                bottom=sides.get('bottom', ex.bottom),
                left=sides.get('left', ex.left),
                right=sides.get('right', ex.right),
            )
        except AttributeError:
            pass

    def _draw_border(s_row: int, e_row: int, s_col: int, e_col: int, borders: dict) -> None:
        has_top    = borders.get('top',    True)
        has_bottom = borders.get('bottom', True)
        has_left   = borders.get('left',   True)
        has_right  = borders.get('right',  True)
        if has_top:
            for c in range(s_col, e_col):
                _set_border_side(s_row - 1, c, bottom=thin)
        if has_bottom:
            for c in range(s_col, e_col):
                _set_border_side(e_row, c, top=thin)
        if has_left:
            for r in range(s_row, e_row):
                _set_border_side(r, s_col, left=thin)
        if has_right:
            for r in range(s_row, e_row):
                _set_border_side(r, e_col - 1, right=thin)

    max_used_row = 0
    max_used_col = 0

    for page_layout in layout:
        page_num = page_layout.get('page_number', 1)
        row_offset = (page_num - 1) * max_rows + ROW_PADDING

        for elem in page_layout.get('elements', []):
            etype = elem.get('type')

            if etype == 'text':
                r = elem.get('row', 1) + row_offset
                c = elem.get('col', 1) + COL_OFFSET
                try:
                    cell = ws.cell(row=r, column=c)
                    cell.value = elem.get('content', '')
                    if elem.get('is_vertical'):
                        cell.alignment = Alignment(text_rotation=255, vertical='top', wrap_text=False)
                    else:
                        # [修正] 水平テキストは PDF 座標準拠で left/top を明示。
                        # Excel のデフォルト挙動（数値の自動右揃え等）を上書きし、
                        # PDF の配置位置をそのまま再現する。
                        # セル内改行（\n）を含む場合は wrap_text=True で折り返しを有効化する。
                        cell.alignment = Alignment(
                            horizontal='left', vertical='top',
                            wrap_text=bool(elem.get('multiline')),
                        )
                    # [修正] フォント名: エイリアス解決済みの値、なければグリッドデフォルト
                    resolved_font_name = elem.get('font_name') or font_name
                    # [修正] フォントサイズ: PDF 値を優先しつつ、セル高さに収まる上限でクランプ
                    raw_font_size = elem.get('font_size') or default_font_size
                    max_font_size = row_height * 0.72  # row_height(pt) の約72%を上限とする
                    resolved_font_size = min(float(raw_font_size), max_font_size)
                    font_kwargs = {
                        'name': resolved_font_name,
                        'size': resolved_font_size,
                    }
                    if elem.get('font_color'):
                        font_kwargs['color'] = elem['font_color']
                    cell.font = Font(**font_kwargs)
                except AttributeError:
                    pass
                max_used_row = max(max_used_row, r)
                max_used_col = max(max_used_col, c)

            elif etype == 'border_rect':
                s_row = elem.get('row', 1) + row_offset
                e_row_v = elem.get('end_row', 1) + row_offset
                s_col = elem.get('col', 1) + COL_OFFSET
                e_col_v = elem.get('end_col', 1) + COL_OFFSET
                borders = elem.get('borders', {'top': True, 'bottom': True, 'left': True, 'right': True})
                _draw_border(s_row, e_row_v, s_col, e_col_v, borders)
                max_used_row = max(max_used_row, e_row_v)
                max_used_col = max(max_used_col, e_col_v)

    for pn in range(1, total_pages):
        ws.row_breaks.append(Break(id=pn * (max_rows + ROW_PADDING)))

    ws.page_setup.paperSize = paper_size
    ws.page_setup.orientation = orientation
    ws.page_margins.left   = margin_left
    ws.page_margins.right  = margin_right
    ws.page_margins.top    = margin_top
    ws.page_margins.bottom = margin_bottom

    # 右端は max_cols + COL_OFFSET まで広げてグリッド幅を最大限利用する
    max_cols = grid_params.get('max_cols', 54)
    max_used_col = max(max_used_col, max_cols + COL_OFFSET)

    if max_used_row > 0 and max_used_col > 0:
        ws.print_area = f"A1:{get_column_letter(max_used_col)}{max_used_row}"

    wb.save(output_path)
    _fix_empty_cell_type_attr(output_path)
    logger.info(f"[render_layout] Excel生成完了: {output_path} ({total_pages} ページ)")


def _generate_border_preview(page_layout: dict, grid_params: dict, output_path: str, pdf_image_path: str | None = None) -> None:
    """
    layout の border_rect 要素を PIL キャンバスに描画し、罫線プレビュー画像を生成する。
    pdf_image_path が指定された場合、その画像と同じ解像度・アスペクト比で生成する。
    """
    from PIL import Image, ImageDraw, ImageFont

    max_c = int(grid_params.get('max_cols', 54))
    max_r = int(grid_params.get('max_rows', 42))

    if pdf_image_path and Path(pdf_image_path).exists():
        with Image.open(pdf_image_path) as ref:
            img_w, img_h = ref.size
        cell_w = img_w / max_c
        cell_h = img_h / max_r
    else:
        cell_w = 20.0
        cell_h = 14.0
        img_w = int(cell_w * max_c) + 1
        img_h = int(cell_h * max_r) + 1

    img = Image.new('RGB', (img_w, img_h), 'white')
    draw = ImageDraw.Draw(img)

    def cx(col: float) -> int: return int(col * cell_w)
    def cy(row: float) -> int: return int(row * cell_h)

    for c in range(max_c + 1):
        x = cx(c)
        draw.line([(x, 0), (x, img_h)], fill='#E0E0E0', width=1)
    for r in range(max_r + 1):
        y = cy(r)
        draw.line([(0, y), (img_w, y)], fill='#E0E0E0', width=1)

    border_width = max(2, int(min(cell_w, cell_h) / 7))
    for elem in page_layout.get('elements', []):
        if elem.get('type') != 'border_rect':
            continue
        r1 = cy(elem['row'] - 1)
        r2 = cy(elem['end_row'])
        c1 = cx(elem['col'] - 1)
        c2 = cx(elem['end_col'])
        borders = elem.get('borders', {'top': True, 'bottom': True, 'left': True, 'right': True})
        if borders.get('top',    True): draw.line([(c1, r1), (c2, r1)], fill='black', width=border_width)
        if borders.get('bottom', True): draw.line([(c1, r2), (c2, r2)], fill='black', width=border_width)
        if borders.get('left',   True): draw.line([(c1, r1), (c1, r2)], fill='black', width=border_width)
        if borders.get('right',  True): draw.line([(c2, r1), (c2, r2)], fill='black', width=border_width)

    try:
        font = ImageFont.load_default(size=max(8, int(cell_h * 0.8)))
    except TypeError:
        font = ImageFont.load_default()
    label_color = (200, 0, 0)
    for c in range(0, max_c + 1, 5):
        draw.text((cx(c) + 1, 1), str(c), fill=label_color, font=font)
    for r in range(0, max_r + 1, 5):
        draw.text((1, cy(r) + 1), str(r), fill=label_color, font=font)

    img.save(output_path)


def _apply_visual_corrections(layout: list, corrections: list) -> list:
    """
    visual_corrections の add_border / remove_border アクションを layout に適用する。
    元の layout は変更せず、ディープコピーに対して操作した結果を返す。
    """
    import copy
    layout = copy.deepcopy(layout)

    for corr in corrections:
        action   = corr.get('action')
        page_num = corr.get('page')
        row      = corr.get('row')
        end_row  = corr.get('end_row')
        col      = corr.get('col')
        end_col  = corr.get('end_col')

        page_layout = next((p for p in layout if p.get('page_number') == page_num), None)
        if page_layout is None:
            logger.warning(f"[corrections] ページ {page_num} が layout に存在しません")
            continue

        if action == 'add_border':
            borders = corr.get('borders', {'top': True, 'bottom': True, 'left': True, 'right': True})
            page_layout['elements'].append({
                'type': 'border_rect',
                'row': row, 'end_row': end_row,
                'col': col, 'end_col': end_col,
                'borders': borders,
            })
        elif action == 'remove_border':
            page_layout['elements'] = [
                e for e in page_layout['elements']
                if not (
                    e.get('type') == 'border_rect' and
                    e.get('row') == row and e.get('end_row') == end_row and
                    e.get('col') == col and e.get('end_col') == end_col
                )
            ]

    return layout


def _apply_borders_to_xlsx(xlsx_path: str, extracted_data: dict, max_rows: int) -> None:
    """
    extracted_data の table_border_rects / rects の _borders を
    openpyxl で直接 XLSX ファイルに適用する（LLM 非依存の確定的罫線描画）。

    GEN_CODE_TEMPLATE と同じオフセット定数を使用:
      COL_OFFSET  = 1  （左1マス余白）
      ROW_PADDING = 1  （ページ上部1マス余白）
      row_offset for page N = (N-1) * max_rows + ROW_PADDING
    """
    from openpyxl import load_workbook
    from openpyxl.styles import Border, Side

    COL_OFFSET = 1
    ROW_PADDING = 1

    wb = load_workbook(xlsx_path)
    ws = wb.active
    thin = Side(style='thin')

    def _draw(s_row: int, e_row: int, s_col: int, e_col: int, borders: dict) -> None:
        has_top    = borders.get('top',    True)
        has_bottom = borders.get('bottom', True)
        has_left   = borders.get('left',   True)
        has_right  = borders.get('right',  True)
        for r in range(s_row, e_row):
            for c in range(s_col, e_col):
                try:
                    cell = ws.cell(row=r, column=c)
                    # 既存の Border オブジェクトを読み取り、各辺を個別にマージ書き込み
                    # （完全上書きすると他のセルが書いた辺を消してしまうため）
                    existing = cell.border
                    new_top    = thin if (r == s_row and has_top)    else existing.top
                    new_bottom = thin if (r == e_row - 1 and has_bottom) else existing.bottom
                    new_left   = thin if (c == s_col and has_left)   else existing.left
                    new_right  = thin if (c == e_col - 1 and has_right)  else existing.right
                    cell.border = Border(top=new_top, bottom=new_bottom,
                                         left=new_left, right=new_right)
                except AttributeError:
                    pass

    total = 0
    for page in extracted_data.get('pages', []):
        page_number = page.get('page_number', 1)
        row_offset = (page_number - 1) * max_rows + ROW_PADDING

        for tbr in page.get('table_border_rects', []):
            borders = tbr.get('_borders', {'top': True, 'bottom': True, 'left': True, 'right': True})
            # ガントチャート等の細幅セル（Excel 2列以下）は内側の垂直罫線を抑制する。
            # ただし以下は例外として縦線を保持する:
            #   - テーブル外周（_outer_left/_outer_right）
            #   - 主要境界（_major_left/_major_right）: 月区切り等の長いエッジが検出された列
            col_span = tbr['_end_col'] - tbr['_col']
            if col_span <= 2:
                borders = dict(borders)
                if not tbr.get('_outer_left', False) and not tbr.get('_major_left', False):
                    borders['left'] = False
                if not tbr.get('_outer_right', False) and not tbr.get('_major_right', False):
                    borders['right'] = False
            _draw(
                tbr['_row'] + row_offset, tbr['_end_row'] + row_offset,
                tbr['_col'] + COL_OFFSET, tbr['_end_col'] + COL_OFFSET,
                borders,
            )
            total += 1

        for rect in page.get('rects', []):
            if '_row' not in rect:
                continue
            r, er = rect['_row'], rect['_end_row']
            c, ec = rect['_col'], rect['_end_col']
            raw_borders = rect.get('_borders', {'top': True, 'bottom': True, 'left': True, 'right': True})
            # 水平・垂直分割線（1行/1列に収まる細長い矩形）の端キャップを除去する
            if r == er and c != ec:
                borders = {'top': raw_borders.get('top', True), 'bottom': False, 'left': False, 'right': False}
            elif c == ec and r != er:
                borders = {'top': False, 'bottom': False, 'left': raw_borders.get('left', True), 'right': False}
            else:
                borders = raw_borders
            _draw(
                r + row_offset, er + row_offset,
                c + COL_OFFSET, ec + COL_OFFSET,
                borders,
            )
            total += 1

    wb.save(xlsx_path)
    _fix_empty_cell_type_attr(xlsx_path)
    logger.info(f"[apply_borders] {total} 個の罫線要素を適用しました")


def _has_japanese(text: str) -> bool:
    """文字列に日本語文字（漢字・ひらがな・カタカナ・全角記号）が含まれるか判定する。"""
    return any(
        '\u3040' <= c <= '\u30ff'  # ひらがな・カタカナ
        or '\u4e00' <= c <= '\u9fff'  # CJK 統合漢字
        or '\uff00' <= c <= '\uffef'  # 全角英数・記号
        for c in text
    )


def _join_word_texts(texts: list) -> str:
    """
    word テキストのリストを結合する。
    テキスト結合ルール:
      - 日本語文字を含む場合はスペースなし
      - 英数字のみの場合は半角スペースで結合
    """
    combined = ''.join(texts)
    if _has_japanese(combined):
        return combined
    return ' '.join(t for t in texts if t.strip())


def _fill_missing_text(layout_json_str: str, extracted_data: dict) -> str:
    """
    LLMが生成したレイアウトJSONに対し、extracted_dataのwordsと照合して
    欠落しているテキスト要素をプログラム的に補完する。

    Step 1 / Step 1.5 の LLM が見落とした word を確実に補う。
    既に text 要素が存在する (row, col) には追加しない（上書き禁止）。
    """
    try:
        layout = json.loads(layout_json_str)
    except (json.JSONDecodeError, ValueError):
        return layout_json_str  # パース失敗時はそのまま返す

    total_added = 0
    for page_layout in layout:
        page_num = page_layout.get('page_number', 1)
        page_data = next(
            (p for p in extracted_data['pages'] if p['page_number'] == page_num),
            None,
        )
        if not page_data:
            continue

        # 既存 text 要素の (row, col) を収集
        existing: set = set()
        for elem in page_layout.get('elements', []):
            if elem.get('type') == 'text' and 'row' in elem and 'col' in elem:
                existing.add((elem['row'], elem['col']))

        # words を (_row, _col) でグループ化
        groups: dict = {}
        for w in page_data.get('words', []):
            if '_row' not in w or '_col' not in w:
                continue
            key = (w['_row'], w['_col'])
            groups.setdefault(key, []).append(w)

        added = []
        for (row, col), words in sorted(groups.items()):
            if (row, col) in existing:
                continue
            content = _join_word_texts([w.get('text', '') for w in words])
            stripped = content.strip()
            # 空白・純粋な区切り記号（ASCII句読点の1文字）はスキップ
            # ただし △▼○● 等の図形記号・日本語1文字は意味があるため残す
            if not stripped or (len(stripped) == 1 and stripped in '!"#$%&\'()*+,-./:;<=>?@[\\]^_`{|}~'):
                continue
            elem: dict = {
                'type': 'text',
                'content': content,
                'row': row,
                'col': col,
                'end_col': col + len(content),
            }
            first = words[0]
            if first.get('font_color') and first['font_color'] != '000000':
                elem['font_color'] = first['font_color']
            if first.get('font_size'):
                elem['font_size'] = first['font_size']
            font_name = _normalize_font_name(first.get('fontname', ''))
            if font_name:
                elem['font_name'] = font_name
            added.append(elem)

        if added:
            page_layout['elements'].extend(added)
            total_added += len(added)

    if total_added:
        logger.info(f"[fill_missing_text] {total_added} 個の欠落テキスト要素を補完しました")

    return json.dumps(layout, ensure_ascii=False)


def _auto_generate_layout(extracted_data: dict, grid_params: dict) -> str:
    """
    extracted_data から直接レイアウトJSONを生成する（step1 + step1.5 のスクリプト代替）。

    - table_border_rects / rects → border_rect 要素
    - words を (_row, _col) でグループ化 → text 要素
    - 座標整合性チェック・重複除去・クランプを適用
    """
    max_rows = grid_params['max_rows']
    max_cols = grid_params['max_cols']

    def _is_near_duplicate(seen: list, r: int, er: int, c: int, ec: int, tol: int = 1) -> bool:
        """グリッド座標が tol 以内の既登録要素があれば重複とみなす。
        完全一致チェックだけでは grid 量子化誤差（±1行/列ずれ）による
        実質同一矩形の二重登録を取りこぼすため、近似一致も除外する。"""
        for (sr, ser, sc, sec) in seen:
            if (abs(sr - r) <= tol and abs(ser - er) <= tol and
                    abs(sc - c) <= tol and abs(sec - ec) <= tol):
                return True
        return False

    layout = []
    for page in extracted_data.get('pages', []):
        elements = []
        seen_border_rects: list = []  # [修正] set → list（近似一致検索のため）

        # table_border_rects → border_rect 要素
        for tbr in page.get('table_border_rects', []):
            r  = min(tbr['_row'],     max_rows)
            er = min(tbr['_end_row'], max_rows)
            c  = min(tbr['_col'],     max_cols)
            ec = min(tbr['_end_col'], max_cols)
            if r > er: r, er = er, r
            if c > ec: c, ec = ec, c
            if r == er and c == ec:
                continue
            # [修正] 完全一致 → ±1グリッド近似一致で重複を除外
            if _is_near_duplicate(seen_border_rects, r, er, c, ec):
                continue
            seen_border_rects.append((r, er, c, ec))
            elements.append({
                'type': 'border_rect',
                'row': r, 'end_row': er,
                'col': c, 'end_col': ec,
                'borders': tbr.get('_borders', {'top': True, 'bottom': True, 'left': True, 'right': True}),
            })

        # rects → border_rect 要素
        for rect in page.get('rects', []):
            if '_row' not in rect:
                continue
            r  = min(rect['_row'],     max_rows)
            er = min(rect['_end_row'], max_rows)
            c  = min(rect['_col'],     max_cols)
            ec = min(rect['_end_col'], max_cols)
            if r > er: r, er = er, r
            if c > ec: c, ec = ec, c
            if r == er and c == ec:
                continue
            # [修正] rects も同じ近似一致チェックで重複を除外
            if _is_near_duplicate(seen_border_rects, r, er, c, ec):
                continue
            seen_border_rects.append((r, er, c, ec))
            # 水平・垂直分割線の場合、_borders が旧データでも端キャップ除去を適用する
            raw_borders = rect.get('_borders', {'top': True, 'bottom': True, 'left': True, 'right': True})
            if r == er and c != ec:
                borders = {'top': raw_borders.get('top', True), 'bottom': False, 'left': False, 'right': False}
            elif c == ec and r != er:
                borders = {'top': False, 'bottom': False, 'left': raw_borders.get('left', True), 'right': False}
            else:
                borders = raw_borders
            elements.append({
                'type': 'border_rect',
                'row': r, 'end_row': er,
                'col': c, 'end_col': ec,
                'borders': borders,
            })

        # words → text 要素（_row, _col でグループ化）
        groups: dict = {}
        for w in page.get('words', []):
            if '_row' not in w or '_col' not in w:
                continue
            groups.setdefault((w['_row'], w['_col']), []).append(w)

        # [修正追加] 同一 (row, col) グループ内でワードの垂直スパンが重ならない場合、
        # 行を分割して別 row に振り直す。
        # build_cluster_map をすり抜けた密集行の誤合算をここで最終補正する。
        _SPLIT_GAP = 3.0  # pt: bottom → next_top のギャップがこれを超えたら別行
        split_groups: dict = {}
        for (row, col), gwords in groups.items():
            if len(gwords) <= 1 or gwords[0].get('is_vertical'):
                split_groups[(row, col)] = gwords
                continue
            # top 順にソートして垂直ギャップを検査
            sw = sorted(gwords, key=lambda w: float(w['top']))
            current: list = [sw[0]]
            sub_row: int = row
            for w in sw[1:]:
                prev_bottom = float(current[-1].get('bottom', current[-1]['top']))
                this_top    = float(w['top'])
                if this_top - prev_bottom > _SPLIT_GAP:
                    # ギャップあり → 現グループを確定し次サブグループを開始
                    split_groups[(sub_row, col)] = current
                    sub_row += 1
                    current = [w]
                else:
                    current.append(w)
            split_groups[(sub_row, col)] = current
        groups = split_groups

        seen_text: set = set()
        for (row, col), words in sorted(groups.items()):
            # 同一グループ内に複数の視覚行が含まれる場合（_SPLIT_GAP 以内の小さなギャップ）、
            # 行ごとに \n で結合してセル内改行として表現する。
            _INLINE_LINE_GAP = 1.0  # pt: この値を超えるギャップを行区切りとみなす
            sw = sorted(words, key=lambda w: float(w.get('top', 0)))
            vis_lines: list = [[sw[0]]]
            for _w in sw[1:]:
                prev_b = float(vis_lines[-1][-1].get('bottom', vis_lines[-1][-1]['top']))
                this_t = float(_w.get('top', 0))
                if this_t - prev_b > _INLINE_LINE_GAP:
                    vis_lines.append([_w])
                else:
                    vis_lines[-1].append(_w)
            if len(vis_lines) > 1:
                content = '\n'.join(
                    _join_word_texts([_w.get('text', '') for _w in _line])
                    for _line in vis_lines
                )
            else:
                content = _join_word_texts([w.get('text', '') for w in words])
            stripped = content.strip()
            if not stripped or (len(stripped) == 1 and stripped in '!"#$%&\'()*+,-./:;<=>?@[\\]^_`{|}~'):
                continue
            row_c = min(row, max_rows)
            col_c = min(col, max_cols)
            pos = (row_c, col_c)
            if pos in seen_text:
                continue
            seen_text.add(pos)
            first = words[0]
            elem: dict = {
                'type': 'text',
                'content': content,
                'row': row_c,
                'col': col_c,
                'end_col': min(col_c + len(content), max_cols),
            }
            if '\n' in content:
                elem['multiline'] = True
            if first.get('font_color') and first['font_color'] != '000000':
                elem['font_color'] = first['font_color']
            if first.get('font_size'):
                elem['font_size'] = first['font_size']
            font_name = _normalize_font_name(first.get('fontname', ''))
            if font_name:
                elem['font_name'] = font_name
            if first.get('is_vertical'):
                elem['is_vertical'] = True
                if '_end_row' in first:
                    elem['end_row'] = min(first['_end_row'], max_rows)
                elem['end_col'] = min(col_c + 1, max_cols)
            elements.append(elem)

        layout.append({'page_number': page['page_number'], 'elements': elements})

    return json.dumps(layout, ensure_ascii=False)


# A4 縦の基準サイズ (pt) — GRID_SIZES のセル密度はこのサイズを基準に調整されている
_A4_W_PT: float = 595.28
_A4_H_PT: float = 841.89


def _setup_grid_params(first_page: dict, grid_size: str) -> dict:
    """
    ページ寸法に基づいてグリッドパラメータを設定する（Sheetling-pre方式）。

    GRID_SIZES の max_cols / max_rows は A4縦(595.28×841.89pt)を基準とする。
    _A4_W_PT / max_cols = pt/列 を基準に、実際のPDFページ寸法から動的にスケーリングし、
    アスペクト比を維持したまま任意の用紙サイズに対応する。
    """
    ref = GRID_SIZES.get(grid_size, GRID_SIZES["small"])
    grid_params = dict(ref)
    grid_params['grid_size'] = grid_size

    # 用紙サイズ・向き検出
    max_dim_pt = max(first_page['width'], first_page['height'])
    grid_params['paper_size'] = 8 if max_dim_pt > 1000 else 9  # 8=A3, 9=A4
    is_landscape = first_page['width'] > first_page['height']
    grid_params['orientation'] = 'landscape' if is_landscape else 'portrait'

    # PDFページ寸法から max_cols / max_rows を動的計算（Sheetling-pre方式）
    pt_per_col = _A4_W_PT / ref['max_cols']
    pt_per_row = _A4_H_PT / ref['max_rows']
    max_cols = max(1, round(first_page['width']  / pt_per_col))
    max_rows = max(1, round(first_page['height'] / pt_per_row))
    grid_params['max_cols'] = max_cols
    grid_params['max_rows'] = max_rows

    # 列幅をページ幅に比例スケール（A4縦基準から横・A3等への対応）
    grid_params['excel_col_width'] = round(ref['excel_col_width'] * ref['max_cols'] / max_cols, 4)

    logger.debug(
        f"[grid] {grid_size} ({grid_params['orientation']}): "
        f"page={first_page['width']:.1f}×{first_page['height']:.1f}pt "
        f"→ max_cols={max_cols}, max_rows={max_rows}, "
        f"excel_col_width={grid_params['excel_col_width']}"
    )

    return grid_params


class SheetlingPipeline:
    """PDF から Excel 方眼紙を自動生成するパイプライン。"""

    def __init__(self, output_base_dir: str):
        self.output_base_dir = Path(output_base_dir)

    def auto_layout(self, pdf_path: str, in_base_dir: str = "data/in", grid_size: str = "small") -> dict:
        """
        [全自動] PDF → Excel 高精度解析パイプライン (Sheetling-pre 方式)。

        1. extract_pdf_data() でテキスト・罫線を抽出
        2. _setup_grid_params() で A4 ポイント基準の方眼密度を動的計算
        3. _compute_grid_coords() で全要素にグリッド座標を付与
        4. _merge_table_border_rects() で結合セルを復元 (Pre版ロジック)
        5. 行シフト補正で上部余白を詰め、コンテンツを上詰めに配置
        6. _auto_generate_layout() でプログラム的にレイアウト JSON を生成
        7. _fill_missing_text() で欠落したテキスト要素を最終補完
        8. _render_layout_to_xlsx() で Excel を直接レンダリング
        """
        logger.info(f"--- [auto] PDF → Excel 高精度自動生成: {Path(pdf_path).name} ---")
        path_obj = Path(pdf_path)
        pdf_name = path_obj.stem

        try:
            rel_path = path_obj.parent.relative_to(Path(in_base_dir))
            out_dir = self.output_base_dir / rel_path
        except ValueError:
            out_dir = self.output_base_dir / pdf_name

        out_dir.mkdir(parents=True, exist_ok=True)
        prompts_dir = out_dir / "prompts" / grid_size
        prompts_dir.mkdir(parents=True, exist_ok=True)

        layout_json_name = f"{pdf_name}_{grid_size}_layout.json"
        grid_params_name = f"{pdf_name}_{grid_size}_grid_params.json"

        # PDF抽出 & グリッド座標付与
        extracted_data = extract_pdf_data(pdf_path)
        first_page = extracted_data['pages'][0]
        grid_params = _setup_grid_params(first_page, grid_size)
        for page in extracted_data['pages']:
            _compute_grid_coords(page, grid_params['max_rows'], grid_params['max_cols'])

        # ---- [Sheetling-pre 移植ロジック] -----------------------------------------
        # 1. 結合セルのマージ
        for page in extracted_data['pages']:
            page['table_border_rects'] = _merge_table_border_rects(page['table_border_rects'])

        # 2. 行シフト補正（上部余白の除去）
        for page in extracted_data['pages']:
            all_rows = (
                [w['_row'] for w in page['words'] if '_row' in w]
                + [r['_row'] for r in page['rects'] if '_row' in r]
                + [tbr['_row'] for tbr in page['table_border_rects']]
            )
            if all_rows:
                row_shift = min(all_rows) - 1
                if row_shift > 0:
                    for w in page['words']:
                        if '_row' in w:
                            w['_row'] -= row_shift
                            if '_end_row' in w: w['_end_row'] -= row_shift
                    for r in page['rects']:
                        if '_row' in r:
                            r['_row'] -= row_shift
                            r['_end_row'] -= row_shift
                    for tbr in page['table_border_rects']:
                        tbr['_row'] -= row_shift
                        tbr['_end_row'] -= row_shift

        # 3. 列シフト補正（左余白を1列に統一）
        # 行シフト補正と同様に、コンテンツ最左列を1にそろえ、
        # COL_OFFSET=1 で左側ちょうど1列の余白になるようにする。
        for page in extracted_data['pages']:
            all_cols = (
                [w['_col'] for w in page['words'] if '_col' in w]
                + [r['_col'] for r in page['rects'] if '_col' in r]
                + [tbr['_col'] for tbr in page['table_border_rects']]
            )
            if all_cols:
                col_shift = min(all_cols) - 1
                if col_shift > 0:
                    for w in page['words']:
                        if '_col' in w:
                            w['_col'] -= col_shift
                    for r in page['rects']:
                        if '_col' in r:
                            r['_col'] -= col_shift
                            r['_end_col'] -= col_shift
                    for tbr in page['table_border_rects']:
                        tbr['_col'] -= col_shift
                        tbr['_end_col'] -= col_shift
        # ---------------------------------------------------------------------------

        # デバッグ用に中間データを保存
        with open(out_dir / f"{pdf_name}_extracted.json", "w", encoding="utf-8") as f:
            json.dump(extracted_data, f, indent=2, ensure_ascii=False)
        with open(out_dir / grid_params_name, "w", encoding="utf-8") as f:
            json.dump(grid_params, f, ensure_ascii=False)

        # レイアウトJSON生成 & 欠落テキスト補完
        layout_json_str = _auto_generate_layout(extracted_data, grid_params)
        filled_json_str = _fill_missing_text(layout_json_str, extracted_data)
        layout_data = json.loads(filled_json_str)

        # アーカイブ・correct コマンド用保存
        output_json_path = out_dir / layout_json_name
        with open(output_json_path, "w", encoding="utf-8") as f:
            f.write(filled_json_str)

        # 直接 Excel レンダリング (Pre版方式)
        # _1pt / _2pt のみサフィックスを付与。それ以外は従来どおり。
        _xlsx_suffix = f"_{grid_size}" if grid_size in ("1pt", "2pt") else ""
        xlsx_path = out_dir / f"{pdf_name}_Python版{_xlsx_suffix}.xlsx"
        try:
            _render_layout_to_xlsx(layout_data, grid_params, str(xlsx_path))
            logger.info(f"✅ Excel 生成完了: {xlsx_path.name}")
        except Exception as e:
            logger.error(f"❌ Excel 生成に失敗しました: {e}")
            raise

        # ---- ビジョンレビュー素材の自動生成 ----------------------------------------
        # correct コマンドで AI 視覚比較できるよう、PDF 画像・罫線プレビュー・
        # VISUAL_REVIEW_PROMPT・corrections テンプレートをまとめて出力する。
        # prompts_dir はすでに auto_layout() 冒頭で作成済み。
        try:
            import pdfplumber as _pdfplumber
            with _pdfplumber.open(pdf_path) as _pdf:
                for _pg in _pdf.pages:
                    _pn = _pg.page_number
                    _pdir = prompts_dir / f"page_{_pn}"
                    _pdir.mkdir(parents=True, exist_ok=True)
                    _img = _pg.to_image(resolution=144)
                    _img.save(str(_pdir / f"{pdf_name}_page{_pn}.png"))
            logger.info(f"  PDF ページ画像を生成しました: {prompts_dir}/page_N/")
        except Exception as _e:
            logger.warning(f"PDF ページ画像の生成に失敗しました（correct フェーズは利用不可）: {_e}")

        for _page_layout in layout_data:
            _pn = _page_layout.get('page_number', 1)
            _pdir = prompts_dir / f"page_{_pn}"
            _pdir.mkdir(parents=True, exist_ok=True)

            _pdf_img = _pdir / f"{pdf_name}_page{_pn}.png"
            _preview  = _pdir / f"{pdf_name}_excel_page{_pn}.png"
            try:
                _generate_border_preview(_page_layout, grid_params, str(_preview),
                                         pdf_image_path=str(_pdf_img))
            except Exception as _e:
                logger.warning(f"  ページ {_pn}: 罫線プレビュー生成に失敗しました: {_e}")

            _gp_for_prompt = dict(grid_params)
            _gp_for_prompt.setdefault('position_tolerance_cells', '1〜2')
            _prompt_text = VISUAL_REVIEW_PROMPT.format(page_number=_pn, **_gp_for_prompt)
            (_pdir / f"{pdf_name}_visual_review_page{_pn}.txt").write_text(_prompt_text, encoding="utf-8")

            _corr_path = _pdir / f"{pdf_name}_visual_corrections_page{_pn}.json"
            if not _corr_path.exists():
                _corr_path.write_text('{"corrections": []}', encoding="utf-8")

        logger.info(
            f"  [review 素材] prompts/{grid_size}/page_N/ に PDF 画像・罫線プレビュー・プロンプトを出力しました\n"
            f"  次のステップ:\n"
            f"    1. PDF 画像と罫線プレビューを AI に渡し、visual_review プロンプトで比較させる\n"
            f"    2. AI の出力 JSON を visual_corrections_page{{N}}.json に保存する\n"
            f"    3. python -m src.main correct --pdf {pdf_name} --grid-size {grid_size} を実行する"
        )
        # -------------------------------------------------------------------------

        return {
            "xlsx_path": str(xlsx_path),
            "layout_json": str(output_json_path),
            "grid_params": grid_params
        }


    def _auto_generate_code(self, pdf_name: str, grid_params: dict, layout_json_name: str = None) -> str:
        """GEN_CODE_TEMPLATE からランタイムにJSONを読み込む _gen.py コードを生成する。"""
        if layout_json_name is None:
            layout_json_name = f"{pdf_name}_layout.json"
        return GEN_CODE_TEMPLATE.substitute(layout_json_name=layout_json_name, **grid_params)

    def render_excel(self, pdf_name: str, specific_out_dir: str = None, apply_border_post_process: bool = True,
                     gen_py_name: str = None, grid_params_name: str = None) -> str:
        """
        Phase 3: AI出力の生成コードを読み込み、Excel方眼紙を描画する。
        apply_border_post_process=False のとき _apply_borders_to_xlsx をスキップする（correct 実行時）。
        gen_py_name / grid_params_name を指定するとパターン別ファイルを使用する。
        """
        logger.info(f"--- [Phase 3] Excel生成: {pdf_name} ---")
        if specific_out_dir:
            out_dir = Path(specific_out_dir)
        else:
            out_dir = self.output_base_dir / pdf_name

        # grid_params_name が指定されていればそれを使用、なければ旧来のパスにフォールバック
        if grid_params_name:
            _grid_params_path = out_dir / grid_params_name
        else:
            _grid_params_path = out_dir / f"{pdf_name}_grid_params.json"

        _grid_size_suffix = ""
        if _grid_params_path.exists():
            try:
                with open(_grid_params_path, "r", encoding="utf-8") as _f:
                    _gp = json.load(_f)
                _gs = _gp.get("grid_size", "")
                if _gs in ("pattern_1", "pattern_2"):
                    _grid_size_suffix = f"_{_gs}"
            except Exception:
                pass
        output_xlsx_path = out_dir / f"{pdf_name}_Python版{_grid_size_suffix}.xlsx"

        # gen_py_name が指定されていればそれを使用、なければ旧来のパスにフォールバック
        if gen_py_name:
            generated_code_path = out_dir / gen_py_name
        else:
            generated_code_path = out_dir / f"{pdf_name}_gen.py"

        if generated_code_path.exists():
            with open(generated_code_path, "r", encoding="utf-8") as f:
                content = f.read().strip()

            code_lines = [line for line in content.splitlines() if not line.strip().startswith("#")]
            actual_code = "\n".join(code_lines).strip()
            is_placeholder = len(actual_code) < 50

            if content and not is_placeholder:
                # 既知の問題パターンを静的チェック・自動修正
                sanitized_content, fixes = _sanitize_generated_code(content)
                if fixes:
                    for fix in fixes:
                        logger.warning(f"🔧 静的修正: {fix}")
                    with open(generated_code_path, "w", encoding="utf-8") as f:
                        f.write(sanitized_content)
                    content = sanitized_content

                logger.info(f"✨ 生成されたコードを実行します: {generated_code_path.name}")
                import subprocess
                import os
                import sys

                try:
                    env = os.environ.copy()
                    env["PYTHONPATH"] = os.getcwd()

                    result = subprocess.run(
                        [sys.executable, generated_code_path.name],
                        cwd=str(out_dir),
                        env=env,
                        capture_output=True,
                        text=True
                    )

                    if result.returncode == 0:
                        temp_xlsx = out_dir / "output.xlsx"
                        if temp_xlsx.exists():
                            temp_xlsx.replace(output_xlsx_path)
                            # 罫線を Python で直接適用（correct 実行時はスキップして corrections を優先）
                            if apply_border_post_process:
                                extracted_json_path = out_dir / f"{pdf_name}_extracted.json"
                                grid_params_path = out_dir / f"{pdf_name}_grid_params.json"
                                if extracted_json_path.exists() and grid_params_path.exists():
                                    try:
                                        with open(extracted_json_path, "r", encoding="utf-8") as f:
                                            extracted_data = json.load(f)
                                        with open(grid_params_path, "r", encoding="utf-8") as f:
                                            grid_params = json.load(f)
                                        _apply_borders_to_xlsx(
                                            str(output_xlsx_path), extracted_data, grid_params['max_rows']
                                        )
                                    except Exception as e:
                                        logger.warning(f"罫線後処理に失敗しました（Excel は生成済み）: {e}")
                            logger.info(f"✅ Phase 3 完了: {output_xlsx_path}")
                            return str(output_xlsx_path)
                        else:
                            error_msg = "生成コードは正常終了しましたが、output.xlsx が生成されませんでした。"
                            logger.error(f"❌ {error_msg}")
                            self._generate_error_prompt(out_dir, pdf_name, error_msg, content)
                    else:
                        error_msg = f"生成コードの実行に失敗しました:\n{result.stderr}"
                        logger.error(f"❌ {error_msg}")
                        self._generate_error_prompt(out_dir, pdf_name, error_msg, content)
                except Exception as e:
                    error_msg = f"生成コード実行中に例外が発生しました: {e}"
                    logger.error(f"❌ {error_msg}")
                    self._generate_error_prompt(out_dir, pdf_name, error_msg, content)
            else:
                logger.warning(f"⚠️ 生成コードファイル {generated_code_path.name} が空、または未編集です。")
        else:
            logger.error(f"❌ 生成コードファイル {generated_code_path.name} が見つかりません。STEP 2 の結果を保存してください。")

        raise RuntimeError(f"Excelの生成に失敗しました ({pdf_name})")

    def apply_corrections(self, pdf_name: str, corrections_json: str, specific_out_dir: str = None,
                          layout_json_name: str = None, gen_py_name: str = None, grid_params_name: str = None) -> None:
        """
        ビジョンLLMが出力した修正指示を _layout.json に適用し、_gen.py を再生成する。
        layout_json_name / gen_py_name / grid_params_name を指定するとパターン別ファイルを使用する。

        corrections_json の形式:
        {
          "corrections": [
            {"action": "add_text",    "page": 1, "row": 3, "col": 5, "content": "テキスト"},
            {"action": "fix_text",    "page": 1, "row": 3, "col": 5, "new_row": 4, "new_col": 6},
            {"action": "add_border",  "page": 1, "row": 3, "end_row": 8, "col": 2, "end_col": 15,
                                      "borders": {"top": true, ...}},
            {"action": "remove_border","page": 1, "row": 3, "end_row": 8, "col": 2, "end_col": 15}
          ]
        }
        """
        if specific_out_dir:
            out_dir = Path(specific_out_dir)
        else:
            out_dir = self.output_base_dir / pdf_name

        _layout_json_name = layout_json_name or f"{pdf_name}_layout.json"
        _gen_py_name = gen_py_name or f"{pdf_name}_gen.py"
        _grid_params_name = grid_params_name or f"{pdf_name}_grid_params.json"

        output_json_path  = out_dir / _layout_json_name
        grid_params_path  = out_dir / _grid_params_name

        if not output_json_path.exists():
            raise FileNotFoundError(f"_layout.json が見つかりません: {output_json_path}")
        if not grid_params_path.exists():
            raise FileNotFoundError(f"_grid_params.json が見つかりません: {grid_params_path}")

        layout = json.loads(output_json_path.read_text(encoding="utf-8"))
        grid_params = json.loads(grid_params_path.read_text(encoding="utf-8"))

        try:
            corrections_data = json.loads(corrections_json)
            corrections = corrections_data.get("corrections", [])
        except json.JSONDecodeError as e:
            raise ValueError(f"corrections JSON のパースに失敗しました: {e}")

        # ページ番号 → elements のマップを構築
        page_map: dict = {p["page_number"]: p["elements"] for p in layout}

        applied = 0
        for c in corrections:
            action  = c.get("action")
            page_no = c.get("page", 1)
            elements = page_map.get(page_no)
            if elements is None:
                logger.warning(f"[correct] ページ {page_no} が見つかりません。スキップします。")
                continue

            if action == "add_text":
                elements.append({
                    "type": "text",
                    "content": c["content"],
                    "row": c["row"],
                    "col": c["col"],
                    "end_col": c["col"] + len(c["content"]),
                })
                applied += 1

            elif action == "fix_text":
                for elem in elements:
                    if elem.get("type") == "text" and elem["row"] == c["row"] and elem["col"] == c["col"]:
                        elem["row"] = c["new_row"]
                        elem["col"] = c["new_col"]
                        applied += 1
                        break

            elif action == "add_border":
                elements.append({
                    "type": "border_rect",
                    "row": c["row"], "end_row": c["end_row"],
                    "col": c["col"], "end_col": c["end_col"],
                    "borders": c.get("borders", {"top": True, "bottom": True, "left": True, "right": True}),
                })
                applied += 1

            elif action == "remove_border":
                # 指定範囲と重複する border_rect をすべて削除（完全一致ではなく重複判定）
                before = len(elements)
                def _overlaps(e: dict, c: dict) -> bool:
                    return (e.get("type") == "border_rect"
                            and e["row"]     <= c["end_row"] and e["end_row"] >= c["row"]
                            and e["col"]     <= c["end_col"] and e["end_col"] >= c["col"])
                # layout オブジェクトが参照する同一リストをインプレースで変更する
                elements[:] = [e for e in elements if not _overlaps(e, c)]
                applied += before - len(elements)

        # 修正済みレイアウトを保存
        updated_json = json.dumps(layout, ensure_ascii=False)
        output_json_path.write_text(updated_json, encoding="utf-8")
        logger.info(f"[correct] {applied} 件の修正を適用しました: {output_json_path}")

        # _gen.py を再生成（パターン別ファイル名を使用）
        gen_code = self._auto_generate_code(pdf_name, grid_params, layout_json_name=_layout_json_name)
        gen_path = out_dir / _gen_py_name
        gen_path.write_text(gen_code, encoding="utf-8")
        logger.info(f"[correct] _gen.py を再生成しました: {gen_path}")

    def switch_pattern(self, pdf_name: str, grid_size: str, specific_out_dir: str = None) -> None:
        """
        _grid_params.json のパターン固有の値を上書きして _gen.py を再生成する。
        correct コマンドで複数パターンを出力する際に使用。
        """
        out_dir = Path(specific_out_dir) if specific_out_dir else self.output_base_dir / pdf_name
        grid_params_path = out_dir / f"{pdf_name}_grid_params.json"

        grid_params = json.loads(grid_params_path.read_text(encoding="utf-8"))
        ref = GRID_SIZES.get(grid_size, GRID_SIZES["small"])
        for key in ["col_width_mm", "row_height_mm", "excel_col_width", "excel_row_height",
                    "margin_left", "margin_right", "margin_top", "margin_bottom", "default_font_size"]:
            if key in ref:
                grid_params[key] = ref[key]
        grid_params["grid_size"] = grid_size

        # PDFページ寸法から max_cols / max_rows を再計算（Sheetling-pre方式）
        _PAPER_DIMS_PT = {8: (841.89, 1190.55), 9: (595.28, 841.89)}
        page_w_pt, page_h_pt = _PAPER_DIMS_PT.get(grid_params.get('paper_size', 9), (595.28, 841.89))
        if grid_params.get('orientation') == 'landscape':
            page_w_pt, page_h_pt = page_h_pt, page_w_pt
        pt_per_col = _A4_W_PT / ref['max_cols']
        pt_per_row = _A4_H_PT / ref['max_rows']
        new_max_cols = max(1, round(page_w_pt / pt_per_col))
        new_max_rows = max(1, round(page_h_pt / pt_per_row))
        grid_params['max_cols'] = new_max_cols
        grid_params['max_rows'] = new_max_rows
        grid_params['excel_col_width'] = round(ref['excel_col_width'] * ref['max_cols'] / new_max_cols, 4)

        grid_params_path.write_text(json.dumps(grid_params, ensure_ascii=False), encoding="utf-8")

        gen_code = self._auto_generate_code(pdf_name, grid_params)
        gen_path = out_dir / f"{pdf_name}_gen.py"
        gen_path.write_text(gen_code, encoding="utf-8")
        logger.info(f"[correct] パターン切り替え完了: {grid_size}")

    def _generate_error_prompt(self, out_dir: Path, pdf_name: str, error_msg: str, current_code: str):
        prompt_text = CODE_ERROR_FIXING_PROMPT.format(error_msg=error_msg, code=current_code)
        prompts_dir = out_dir / "prompts"
        prompts_dir.mkdir(parents=True, exist_ok=True)
        error_prompt_path = prompts_dir / f"{pdf_name}_prompt_error_fix.txt"
        with open(error_prompt_path, "w", encoding="utf-8") as f:
            f.write(prompt_text)
        logger.info(f"💡 エラー修正用プロンプトを出力しました: {error_prompt_path}")

    # ===========================================================================
    # pre版から移植: LLM協業モード（高精度フロー）
    # extract → fill → review → generate の4ステップ
    # ===========================================================================

    def generate_prompts(self, pdf_path: str, in_base_dir: str = "data/in") -> dict:
        """
        LLM協業 Step 1: PDF を解析し、ページごとに STEP1・STEP1.5 プロンプトを生成する。

        出力:
          - {pdf_name}_extracted.json  （グリッド座標付き）
          - {pdf_name}_grid_params.json
          - prompts/page_{N}/{pdf_name}_prompt_step1_page{N}.txt
          - prompts/page_{N}/{pdf_name}_prompt_step1_5_page{N}.txt
          - prompts/page_{N}/{pdf_name}_step1_5_input_page{N}.json  （LLM出力貼付用）
          - prompts/page_{N}/{pdf_name}_page{N}.png  （PDFページ画像）
        """
        logger.info(f"--- [extract] PDF解析 & プロンプト生成: {Path(pdf_path).name} ---")
        path_obj = Path(pdf_path)
        pdf_name = path_obj.stem

        try:
            rel_path = path_obj.parent.relative_to(Path(in_base_dir))
            out_dir = self.output_base_dir / rel_path
        except ValueError:
            out_dir = self.output_base_dir / pdf_name

        out_dir.mkdir(parents=True, exist_ok=True)

        extracted_data = extract_pdf_data(pdf_path)
        first_page = extracted_data['pages'][0]

        # small グリッドを base に用紙サイズ動的計算（pre版方式）
        _A4_W_PT: float = 595.28
        _A4_H_PT: float = 841.89
        ref = GRID_SIZES["small"]
        grid_params = dict(ref)
        pt_per_col = _A4_W_PT / ref['max_cols']
        pt_per_row = _A4_H_PT / ref['max_rows']
        grid_params['max_cols'] = max(1, round(first_page['width']  / pt_per_col))
        grid_params['max_rows'] = max(1, round(first_page['height'] / pt_per_row))
        max_dim_pt = max(first_page['width'], first_page['height'])
        grid_params['paper_size'] = 8 if max_dim_pt > 1000 else 9
        is_landscape = first_page['width'] > first_page['height']
        grid_params['orientation'] = 'landscape' if is_landscape else 'portrait'

        with open(out_dir / f"{pdf_name}_grid_params.json", "w", encoding="utf-8") as f:
            json.dump(grid_params, f, ensure_ascii=False)

        for page in extracted_data['pages']:
            _compute_grid_coords(page, grid_params['max_rows'], grid_params['max_cols'])

        # _merge_table_border_rects を適用（結合セルの統合）
        for page in extracted_data['pages']:
            page['table_border_rects'] = _merge_table_border_rects(page['table_border_rects'])

        extracted_json_path = out_dir / f"{pdf_name}_extracted.json"
        with open(extracted_json_path, "w", encoding="utf-8") as f:
            json.dump(extracted_data, f, indent=2, ensure_ascii=False)

        _WORD_KEEP = {"text", "_row", "_col", "font_color", "font_size", "is_vertical", "_end_row"}
        _TBR_KEEP  = {"_row", "_end_row", "_col", "_end_col", "_borders"}
        _RECT_KEEP = {"_row", "_col", "_end_row", "_end_col", "_borders"}

        prompts_dir = out_dir / "prompts"
        prompts_dir.mkdir(parents=True, exist_ok=True)
        total_pages = len(extracted_data['pages'])

        for page in extracted_data['pages']:
            page_num = page['page_number']
            page_dir = prompts_dir / f"page_{page_num}"
            page_dir.mkdir(parents=True, exist_ok=True)

            page_step1_data = {"pages": [{
                "page_number": page_num,
                "words": [
                    {k: v for k, v in w.items() if k in _WORD_KEEP}
                    for w in page.get("words", [])
                ],
                "table_border_rects": [
                    {k: v for k, v in r.items() if k in _TBR_KEEP}
                    for r in page.get("table_border_rects", [])
                ],
                "rects": [
                    {k: v for k, v in r.items() if k in _RECT_KEEP}
                    for r in page.get("rects", [])
                ],
            }]}

            page_slim_data = {"pages": [{
                "page_number": page_num,
                "words": [
                    {"text": w.get("text", ""), "_row": w["_row"], "_col": w["_col"]}
                    for w in page.get("words", [])
                ]
            }]}

            prompt_1 = TABLE_ANCHOR_PROMPT.format(
                input_data=json.dumps(page_step1_data, indent=2, ensure_ascii=False),
                **grid_params
            )
            with open(page_dir / f"{pdf_name}_prompt_step1_page{page_num}.txt", "w", encoding="utf-8") as f:
                f.write(prompt_1)

            prompt_1_5 = LAYOUT_REVIEW_PROMPT.format(
                input_data=json.dumps(page_slim_data, indent=2, ensure_ascii=False),
                step1_output="[ここにSTEP 1の出力（JSON部分のみ）を貼り付けてください]",
                **grid_params
            )
            with open(page_dir / f"{pdf_name}_prompt_step1_5_page{page_num}.txt", "w", encoding="utf-8") as f:
                f.write(prompt_1_5)

            step1_5_input_path = page_dir / f"{pdf_name}_step1_5_input_page{page_num}.json"
            if not step1_5_input_path.exists():
                with open(step1_5_input_path, "w", encoding="utf-8") as f:
                    f.write(
                        f"// ページ {page_num}/{total_pages}: STEP 1.5 の LLM 出力 JSON をここに貼り付けてください。\n"
                        f"// 全ページ貼り付け後に python -m src.main fill を実行してください。\n"
                    )

            logger.info(f"  ページ {page_num}/{total_pages}: STEP1 / STEP1.5 プロンプト生成完了")

        # PDFページ画像生成
        try:
            import pdfplumber
            with pdfplumber.open(pdf_path) as pdf_img:
                for idx, page in enumerate(extracted_data['pages']):
                    page_num = page['page_number']
                    page_dir = prompts_dir / f"page_{page_num}"
                    img = pdf_img.pages[idx].to_image(resolution=144)
                    img_path = page_dir / f"{pdf_name}_page{page_num}.png"
                    img.save(str(img_path))
            logger.info(f"  PDF ページ画像を生成しました: {prompts_dir}/page_N/")
        except Exception as e:
            logger.warning(f"PDF ページ画像の生成に失敗しました（review フェーズは利用不可）: {e}")

        logger.info(f"✅ extract 完了: {pdf_name} ({total_pages} ページ)")
        logger.info(f"  抽出データ: {extracted_json_path}")
        logger.info(f"  プロンプト: {prompts_dir}/page_N/")
        logger.info(f"  ※ 各ページ STEP1 → STEP1.5 → step1_5_input_pageN.json に貼付 → fill → generate")

        return {
            "json_path": str(extracted_json_path),
            "prompts_dir": str(prompts_dir),
            "total_pages": total_pages,
        }

    def fill_layout(self, pdf_name: str, step1_5_json: str, page_num: int,
                    specific_out_dir: str = None) -> str:
        """
        LLM協業 Step 2: 1ページ分の STEP1.5 LLM 出力に対してプログラム的ギャップ補完を適用する。

        LLM が見落とした text 要素を extracted.json と照合して補完する。
        """
        if specific_out_dir:
            out_dir = Path(specific_out_dir)
        else:
            out_dir = self.output_base_dir / pdf_name

        extracted_json_path = out_dir / f"{pdf_name}_extracted.json"
        if not extracted_json_path.exists():
            raise FileNotFoundError(
                f"extracted.json が見つかりません: {extracted_json_path}. "
                "generate_prompts() を先に実行してください。"
            )
        with open(extracted_json_path, "r", encoding="utf-8") as f:
            extracted_data = json.load(f)

        single_page_data = {
            "pages": [p for p in extracted_data["pages"] if p["page_number"] == page_num]
        }

        filled_json = _fill_missing_text(step1_5_json, single_page_data)

        page_dir = out_dir / "prompts" / f"page_{page_num}"
        page_dir.mkdir(parents=True, exist_ok=True)
        filled_json_path = page_dir / f"{pdf_name}_step1_5_output_page{page_num}.json"
        with open(filled_json_path, "w", encoding="utf-8") as f:
            f.write(filled_json)
        logger.info(f"[fill] ページ {page_num}: 補完済みJSON保存 → {filled_json_path.name}")

        return filled_json

    def _save_merged_layout(self, pdf_name: str, out_dir: Path, grid_params: dict,
                            force: bool = False) -> bool:
        """
        全ページの fill が完了しているか確認し、完了していれば全ページをマージして
        {pdf_name}_layout.json を保存する。

        Returns:
            True: layout.json を保存した / False: まだ未完了のページがある（force=False 時のみ）
        """
        prompts_dir = out_dir / "prompts"

        output_files = sorted(prompts_dir.rglob(f"{pdf_name}_step1_5_output_page*.json"))
        if not output_files:
            return False

        input_files = sorted(prompts_dir.rglob(f"{pdf_name}_step1_5_input_page*.json"))
        if len(output_files) < len(input_files):
            remaining = len(input_files) - len(output_files)
            if not force:
                logger.info(
                    f"[fill] あと {remaining} ページが未入力です。"
                    f"（--force で完了済み {len(output_files)} ページのみで更新できます）"
                )
                return False
            logger.info(
                f"[fill --force] {remaining} ページをスキップし、"
                f"完了済み {len(output_files)} ページのみでレイアウトを保存します。"
            )

        merged_layout: list = []
        for output_file in output_files:
            try:
                data = json.loads(output_file.read_text(encoding="utf-8"))
                if isinstance(data, list):
                    merged_layout.extend(data)
                elif isinstance(data, dict):
                    merged_layout.append(data)
            except (json.JSONDecodeError, ValueError) as e:
                logger.warning(f"[fill] {output_file.name} のパースに失敗しました: {e}")
                return False

        layout_path = out_dir / f"{pdf_name}_layout.json"
        with open(layout_path, "w", encoding="utf-8") as f:
            json.dump(merged_layout, f, ensure_ascii=False, indent=2)

        logger.info(f"✅ 全ページ完了: レイアウト JSON 保存 → {layout_path}")
        return True

    def render_excel_from_layout(self, pdf_name: str, specific_out_dir: str = None) -> str:
        """
        LLM協業 Step 4: layout.json を読み込み、openpyxl で直接 Excel 方眼紙を描画する。
        visual_corrections が存在すれば適用してから描画する。
        """
        logger.info(f"--- [generate] Excel生成: {pdf_name} ---")
        if specific_out_dir:
            out_dir = Path(specific_out_dir)
        else:
            out_dir = self.output_base_dir / pdf_name

        layout_json_path = out_dir / f"{pdf_name}_layout.json"
        grid_params_path = out_dir / f"{pdf_name}_grid_params.json"
        output_xlsx_path = out_dir / f"{pdf_name}.xlsx"

        if not layout_json_path.exists():
            raise FileNotFoundError(
                f"layout.json が見つかりません: {layout_json_path}. "
                "python -m src.main fill を先に実行してください。"
            )
        if not grid_params_path.exists():
            raise FileNotFoundError(
                f"grid_params.json が見つかりません: {grid_params_path}. "
                "python -m src.main extract を先に実行してください。"
            )

        with open(layout_json_path, "r", encoding="utf-8") as f:
            layout = json.load(f)
        with open(grid_params_path, "r", encoding="utf-8") as f:
            grid_params = json.load(f)

        # extracted.json の最新 table_border_rects / rects で border_rect を上書き
        extracted_json_path = out_dir / f"{pdf_name}_extracted.json"
        if extracted_json_path.exists():
            with open(extracted_json_path, "r", encoding="utf-8") as f:
                extracted_data = json.load(f)
            for page_layout in layout:
                page_num = page_layout.get('page_number', 1)
                ext_page = next(
                    (p for p in extracted_data['pages'] if p['page_number'] == page_num), None
                )
                if ext_page is None:
                    continue
                non_border = [e for e in page_layout.get('elements', []) if e.get('type') != 'border_rect']
                fresh_borders: list = []
                for tbr in ext_page.get('table_border_rects', []):
                    fresh_borders.append({
                        'type': 'border_rect',
                        'row': tbr['_row'],     'end_row': tbr['_end_row'],
                        'col': tbr['_col'],     'end_col': tbr['_end_col'],
                        'borders': tbr['_borders'],
                    })
                for r in ext_page.get('rects', []):
                    if '_row' not in r:
                        continue
                    fresh_borders.append({
                        'type': 'border_rect',
                        'row': r['_row'],     'end_row': r['_end_row'],
                        'col': r['_col'],     'end_col': r['_end_col'],
                        'borders': r.get('_borders', {'top': True, 'bottom': True, 'left': True, 'right': True}),
                    })
                page_layout['elements'] = non_border + fresh_borders
            logger.info(f"[generate] extracted.json の最新 border_rects を適用しました")

        # corrections ファイルがあれば適用
        prompts_dir = out_dir / "prompts"
        all_corrections: list = []
        for corr_file in sorted(prompts_dir.rglob(f"{pdf_name}_visual_corrections_page*.json")):
            try:
                with open(corr_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                corrections = data.get("corrections", [])
                if corrections:
                    all_corrections.extend(corrections)
            except Exception as e:
                logger.warning(f"corrections ファイルの読み込みに失敗しました: {corr_file.name}: {e}")

        if all_corrections:
            layout = _apply_visual_corrections(layout, all_corrections)
            logger.info(f"[generate] {len(all_corrections)} 件の罫線補正を適用しました")

        _render_layout_to_xlsx(layout, grid_params, str(output_xlsx_path))

        logger.info(f"✅ generate 完了: {output_xlsx_path}")
        return str(output_xlsx_path)

    def generate_visual_review(self, pdf_name: str, specific_out_dir: str = None) -> dict:
        """
        LLM協業 Step review: layout.json の border_rect から罫線プレビュー画像を生成し、
        VISUAL_REVIEW_PROMPT と corrections テンプレートファイルを出力する。

        出力ファイル（各ページ）:
          prompts/page_{N}/{pdf_name}_excel_page{N}.png   ← 罫線プレビュー画像
          prompts/page_{N}/{pdf_name}_visual_review_page{N}.txt ← LLM へ渡すプロンプト
          prompts/page_{N}/{pdf_name}_visual_corrections_page{N}.json ← LLM 出力を貼り付け
        """
        logger.info(f"--- [review] 罫線プレビュー生成: {pdf_name} ---")
        if specific_out_dir:
            out_dir = Path(specific_out_dir)
        else:
            out_dir = self.output_base_dir / pdf_name

        layout_json_path = out_dir / f"{pdf_name}_layout.json"
        grid_params_path = out_dir / f"{pdf_name}_grid_params.json"

        if not layout_json_path.exists():
            raise FileNotFoundError(
                f"layout.json が見つかりません: {layout_json_path}. "
                "python -m src.main fill を先に実行してください。"
            )
        if not grid_params_path.exists():
            raise FileNotFoundError(
                f"grid_params.json が見つかりません: {grid_params_path}. "
                "python -m src.main extract を先に実行してください。"
            )

        with open(layout_json_path, "r", encoding="utf-8") as f:
            layout = json.load(f)
        with open(grid_params_path, "r", encoding="utf-8") as f:
            grid_params = json.load(f)
        # 旧 grid_params.json には position_tolerance_cells が未保存の場合があるためフォールバック
        grid_params.setdefault('position_tolerance_cells', '1〜2')

        prompts_dir = out_dir / "prompts"
        generated = []

        for page_layout in layout:
            page_num = page_layout.get('page_number', 1)
            page_dir = prompts_dir / f"page_{page_num}"
            page_dir.mkdir(parents=True, exist_ok=True)

            preview_path = page_dir / f"{pdf_name}_excel_page{page_num}.png"
            pdf_img_path = page_dir / f"{pdf_name}_page{page_num}.png"
            try:
                _generate_border_preview(page_layout, grid_params, str(preview_path),
                                         pdf_image_path=str(pdf_img_path))
            except Exception as e:
                logger.warning(f"  ページ {page_num}: プレビュー画像の生成に失敗しました: {e}")

            prompt_text = VISUAL_REVIEW_PROMPT.format(
                page_number=page_num,
                **grid_params,
            )
            prompt_path = page_dir / f"{pdf_name}_visual_review_page{page_num}.txt"
            prompt_path.write_text(prompt_text, encoding="utf-8")

            corrections_path = page_dir / f"{pdf_name}_visual_corrections_page{page_num}.json"
            if not corrections_path.exists():
                corrections_path.write_text('{"corrections": []}', encoding="utf-8")

            pdf_image_path = page_dir / f"{pdf_name}_page{page_num}.png"

            generated.append({
                "page_num":    page_num,
                "pdf_image":   str(pdf_image_path),
                "preview":     str(preview_path),
                "prompt":      str(prompt_path),
                "corrections": str(corrections_path),
            })

            logger.info(f"  ページ {page_num}: プレビュー → {preview_path.name}")

        logger.info(f"✅ review 完了: {pdf_name} ({len(generated)} ページ)")
        logger.info("  次の手順:")
        logger.info("    1. 各ページの PDF 画像とプレビュー画像を LLM に渡す")
        logger.info("    2. プロンプトファイルの内容を使って LLM を実行する")
        logger.info("    3. LLM の出力 JSON を visual_corrections_page{N}.json に保存する")
        logger.info("    4. python -m src.main generate を実行する")
        return {"pages": generated}

    def rerender_after_corrections(
        self,
        pdf_name: str,
        grid_size: str,
        specific_out_dir: str = None,
    ) -> str:
        """
        correct コマンド用: 修正済み layout JSON + grid_params から Excel を再レンダリングする。

        auto_layout() が生成する grid_size サフィックス付きファイルを参照する。
          - {pdf_name}_{grid_size}_layout.json   （apply_corrections が更新済み）
          - {pdf_name}_{grid_size}_grid_params.json
        出力:
          - {pdf_name}_Python版.xlsx          （1pt/2pt 以外）
          - {pdf_name}_Python版_{grid_size}.xlsx  （1pt/2pt）
        """
        logger.info(f"--- [correct/rerender] Excel 再生成: {pdf_name} ({grid_size}) ---")
        out_dir = Path(specific_out_dir) if specific_out_dir else self.output_base_dir / pdf_name

        layout_path     = out_dir / f"{pdf_name}_{grid_size}_layout.json"
        grid_params_path = out_dir / f"{pdf_name}_{grid_size}_grid_params.json"
        _xlsx_suffix = f"_{grid_size}" if grid_size in ("1pt", "2pt") else ""
        xlsx_path    = out_dir / f"{pdf_name}_Python版{_xlsx_suffix}.xlsx"

        if not layout_path.exists():
            raise FileNotFoundError(f"layout JSON が見つかりません: {layout_path}")
        if not grid_params_path.exists():
            raise FileNotFoundError(f"grid_params JSON が見つかりません: {grid_params_path}")

        layout      = json.loads(layout_path.read_text(encoding="utf-8"))
        grid_params = json.loads(grid_params_path.read_text(encoding="utf-8"))

        _render_layout_to_xlsx(layout, grid_params, str(xlsx_path))
        logger.info(f"✅ correct/rerender 完了: {xlsx_path.name}")
        return str(xlsx_path)

