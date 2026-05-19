"""エッジ単位の罫線モデルとアグリゲーション。

罫線の最小単位は「セル境界」(cell-edge)。
  - ('H', row, col) = セル(row, col) の上辺
  - ('V', row, col) = セル(row, col) の左辺

連続するセル境界を「ラン (run)」にまとめて扱う。
  - H run: row=N, col_start=cs, col_end=ce  (col_end は排他的境界 = 最終列+1)
  - V run: col=N, row_start=rs, row_end=re  (row_end は排他的境界 = 最終行+1)
"""

from typing import Iterable


def decompose_to_cell_edges(elements: list) -> tuple[set, dict]:
    """border_rect 要素群をセル境界の集合に分解する。

    Returns:
        cell_edges: {('H'|'V', row, col)} の集合
        styles: 各 cell_edge → border_style のマップ (重複時は先勝ち)
    """
    cell_edges: set = set()
    styles: dict = {}
    for elem in elements:
        if elem.get('type') != 'border_rect':
            continue
        r, er = elem['row'], elem['end_row']
        c, ec = elem['col'], elem['end_col']
        borders = elem.get('borders', {'top': True, 'bottom': True, 'left': True, 'right': True})
        bs = elem.get('border_style', 'thin')
        if borders.get('top', True):
            for cc in range(c, ec):
                edge = ('H', r, cc)
                cell_edges.add(edge)
                styles.setdefault(edge, bs)
        if borders.get('bottom', True):
            for cc in range(c, ec):
                edge = ('H', er, cc)
                cell_edges.add(edge)
                styles.setdefault(edge, bs)
        if borders.get('left', True):
            for rr in range(r, er):
                edge = ('V', rr, c)
                cell_edges.add(edge)
                styles.setdefault(edge, bs)
        if borders.get('right', True):
            for rr in range(r, er):
                edge = ('V', rr, ec)
                cell_edges.add(edge)
                styles.setdefault(edge, bs)
    return cell_edges, styles


def _group_h_runs(h_by_row: dict, styles: dict) -> list:
    runs = []
    for row, items in sorted(h_by_row.items()):
        items.sort()
        cur_start = items[0]
        cur_end = cur_start
        cur_style = styles.get(('H', row, cur_start), 'thin')
        for col in items[1:]:
            style = styles.get(('H', row, col), 'thin')
            if col == cur_end + 1 and style == cur_style:
                cur_end = col
            else:
                runs.append({'type': 'H', 'row': row,
                             'col_start': cur_start, 'col_end': cur_end + 1,
                             'border_style': cur_style})
                cur_start, cur_end, cur_style = col, col, style
        runs.append({'type': 'H', 'row': row,
                     'col_start': cur_start, 'col_end': cur_end + 1,
                     'border_style': cur_style})
    return runs


def _group_v_runs(v_by_col: dict, styles: dict) -> list:
    runs = []
    for col, items in sorted(v_by_col.items()):
        items.sort()
        cur_start = items[0]
        cur_end = cur_start
        cur_style = styles.get(('V', cur_start, col), 'thin')
        for row in items[1:]:
            style = styles.get(('V', row, col), 'thin')
            if row == cur_end + 1 and style == cur_style:
                cur_end = row
            else:
                runs.append({'type': 'V', 'col': col,
                             'row_start': cur_start, 'row_end': cur_end + 1,
                             'border_style': cur_style})
                cur_start, cur_end, cur_style = row, row, style
        runs.append({'type': 'V', 'col': col,
                     'row_start': cur_start, 'row_end': cur_end + 1,
                     'border_style': cur_style})
    return runs


def group_into_runs(cell_edges: Iterable, styles: dict) -> list:
    """連続する同 style のセル境界を最大長のランに集約する。"""
    h_by_row: dict = {}
    v_by_col: dict = {}
    for t, r, c in cell_edges:
        if t == 'H':
            h_by_row.setdefault(r, []).append(c)
        else:
            v_by_col.setdefault(c, []).append(r)
    return _group_h_runs(h_by_row, styles) + _group_v_runs(v_by_col, styles)


def runs_to_border_rects(runs: list) -> list:
    """ラン群を border_rect 要素群に変換する (1ラン = 1 border_rect)。

    H ラン → end_row == row のゼロ高 rect (top のみ True)
    V ラン → end_col == col のゼロ幅 rect (left のみ True)
    レンダラ (excel.py / preview.py) はこの形式を正常に処理する。
    """
    elements = []
    for run in runs:
        bs = run.get('border_style', 'thin')
        if run['type'] == 'H':
            r = run['row']
            elements.append({
                'type': 'border_rect',
                'row': r, 'end_row': r,
                'col': run['col_start'], 'end_col': run['col_end'],
                'borders': {'top': True, 'bottom': False, 'left': False, 'right': False},
                'border_style': bs,
            })
        else:
            c = run['col']
            elements.append({
                'type': 'border_rect',
                'row': run['row_start'], 'end_row': run['row_end'],
                'col': c, 'end_col': c,
                'borders': {'top': False, 'bottom': False, 'left': True, 'right': False},
                'border_style': bs,
            })
    return elements


def filter_short_runs(elements: list, min_h_span: int, min_v_span: int) -> None:
    """短いスパンのランを layout 要素から除去する（in-place）。

    border_rect を一度セル境界に分解→ランに集約した後、
    スパンが閾値未満のランを除外して再構築する。
    rects / h_edges / v_edges の区別なく全ソースに適用される。

    Args:
        elements: layout のページ要素リスト
        min_h_span: H ランの最小列スパン（これ未満は除去）
        min_v_span: V ランの最小行スパン（これ未満は除去）
    """
    cell_edges, styles = decompose_to_cell_edges(elements)
    runs = group_into_runs(cell_edges, styles)
    # ランの境界は exclusive（col_end/row_end = 最終セル+1）なので
    # inclusive セル数 N = exclusive span - 1。
    # min_h_span=2 で「inclusive 2セル以上を保持」→ exclusive span > 2 が条件。
    filtered = [
        r for r in runs
        if (r['type'] == 'H' and r['col_end'] - r['col_start'] > min_h_span)
        or (r['type'] == 'V' and r['row_end'] - r['row_start'] > min_v_span)
    ]
    new_rects = runs_to_border_rects(filtered)
    non_border = [e for e in elements if e.get('type') != 'border_rect']
    elements.clear()
    elements.extend(non_border)
    elements.extend(new_rects)
