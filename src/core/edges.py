"""エッジ単位の罫線モデルとアグリゲーション。

罫線の最小単位は「セル境界」(cell-edge)。
  - ('H', row, col) = セル(row, col) の上辺
  - ('V', row, col) = セル(row, col) の左辺

LLM I/O 用には連続するセル境界を「ラン (run)」にまとめて扱う。
  - H run: row=N, col_start=cs, col_end=ce  (col_end は排他的境界 = 最終列+1)
  - V run: col=N, row_start=rs, row_end=re  (row_end は排他的境界 = 最終行+1)

apply ロジックは flatten → set 演算 → 再集約 の3段で実装する。
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


def run_to_cell_edges(run: dict) -> list:
    """ラン dict をセル境界タプルのリストに展開する。"""
    if run['type'] == 'H':
        return [('H', run['row'], c) for c in range(run['col_start'], run['col_end'])]
    return [('V', r, run['col']) for r in range(run['row_start'], run['row_end'])]


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


def enumerate_runs_with_ids(elements: list) -> list:
    """layout の border_rect 群を ID 付きランリストに変換する。

    LLM プロンプトに渡す比較対象の正本となる。
    ID は 1 から連番で振る (出力順は H→V、行/列昇順)。
    """
    cell_edges, styles = decompose_to_cell_edges(elements)
    runs = group_into_runs(cell_edges, styles)
    return [{'id': i + 1, **run} for i, run in enumerate(runs)]


def apply_edge_corrections(elements: list, removed_ids: list, added_runs: list,
                            id_map: dict) -> int:
    """エッジ単位の修正を elements に適用する。

    Args:
        elements: layout のページ要素リスト (in-place で書き換え)
        removed_ids: 削除対象のラン ID リスト
        added_runs: 追加するラン dict のリスト
        id_map: ID → ラン dict のマップ (auto 時生成)

    Returns:
        適用されたセル境界の総数 (削除 + 追加)
    """
    cell_edges, styles = decompose_to_cell_edges(elements)

    removed = 0
    for rid in removed_ids:
        run = id_map.get(rid) if isinstance(rid, int) else id_map.get(int(rid))
        if run is None:
            continue
        for edge in run_to_cell_edges(run):
            if edge in cell_edges:
                cell_edges.discard(edge)
                removed += 1

    added = 0
    for run in added_runs:
        bs = run.get('border_style', 'thin')
        for edge in run_to_cell_edges(run):
            if edge not in cell_edges:
                cell_edges.add(edge)
                styles[edge] = bs
                added += 1

    new_runs = group_into_runs(cell_edges, styles)
    new_rects = runs_to_border_rects(new_runs)

    non_border = [e for e in elements if e.get('type') != 'border_rect']
    elements.clear()
    elements.extend(non_border)
    elements.extend(new_rects)
    return removed + added
