"""PDF の rect/edge/table 枠から Excel 罫線要素を生成する。

seen_edges でセル境界単位の dedup を行い、同じ境界に罫線が
重複して発行されるのを防ぐ（例: table と rect が同一セルを共有する場合）。
"""

from src.utils.font import linewidth_to_border_style


def _edges_of_side(r: int, er: int, c: int, ec: int, side: str) -> set:
    if side == 'top':
        return {('H', r, cc) for cc in range(c, ec)}
    if side == 'bottom':
        return {('H', er, cc) for cc in range(c, ec)}
    if side == 'left':
        return {('V', rr, c) for rr in range(r, er)}
    if side == 'right':
        return {('V', rr, ec) for rr in range(r, er)}
    return set()


def _filter_sides_by_seen(r: int, er: int, c: int, ec: int,
                          borders: dict, seen_edges: set) -> dict | None:
    """Drop sides whose cell-edges are already all in seen_edges.

    Returns a new borders dict with redundant sides set to False, or None
    if all requested sides are fully redundant. seen_edges is updated
    in-place with any newly-covered edges.
    """
    new_sides = {'top': False, 'bottom': False, 'left': False, 'right': False}
    any_new = False
    for side in ('top', 'bottom', 'left', 'right'):
        if not borders.get(side):
            continue
        edges = _edges_of_side(r, er, c, ec, side)
        if edges and (edges - seen_edges):
            new_sides[side] = True
            seen_edges.update(edges)
            any_new = True
    return new_sides if any_new else None


def _collect_table_border_elements(page, max_rows, max_cols, seen_edges: set) -> list:
    elements = []
    for tbr in page.get('table_border_rects', []):
        r  = min(tbr['_row'],     max_rows)
        er = min(tbr['_end_row'], max_rows)
        c  = min(tbr['_col'],     max_cols)
        ec = min(tbr['_end_col'], max_cols)
        if r > er: r, er = er, r
        if c > ec: c, ec = ec, c
        if r == er and c == ec:
            continue
        borders = tbr.get('_borders', {'top': True, 'bottom': True, 'left': True, 'right': True})
        sides = _filter_sides_by_seen(r, er, c, ec, borders, seen_edges)
        if sides is None:
            continue
        elements.append({
            'type': 'border_rect',
            'row': r, 'end_row': er, 'col': c, 'end_col': ec,
            'borders': sides,
        })
    return elements


def _emit_rect_line(r: int, er: int, c: int, ec: int, borders: dict,
                    bs: str, seen_edges: set) -> dict | None:
    sides = _filter_sides_by_seen(r, er, c, ec, borders, seen_edges)
    if sides is None:
        return None
    return {'type': 'border_rect', 'row': r, 'end_row': er, 'col': c, 'end_col': ec,
            'borders': sides, 'border_style': bs}


def _collect_rect_border_elements(page, max_rows, max_cols, seen_edges: set) -> list:
    elements = []
    for rect in page.get('rects', []):
        if '_row' not in rect:
            continue
        r  = min(rect['_row'],     max_rows)
        er = min(rect['_end_row'], max_rows)
        c  = min(rect['_col'],     max_cols)
        ec = min(rect['_end_col'], max_cols)
        if r > er: r, er = er, r
        if c > ec: c, ec = ec, c
        bs = linewidth_to_border_style(rect.get('linewidth', 0.0))

        if r == er and c != ec:
            el = _emit_rect_line(r, r + 1, c, ec,
                                 {'top': True, 'bottom': False, 'left': False, 'right': False},
                                 bs, seen_edges)
        elif c == ec and r != er:
            el = _emit_rect_line(r, er, c, c + 1,
                                 {'top': False, 'bottom': False, 'left': True, 'right': False},
                                 bs, seen_edges)
        elif r == er and c == ec:
            el = None
        else:
            el = _emit_rect_line(r, er, c, ec,
                                 {'top': True, 'bottom': True, 'left': True, 'right': True},
                                 bs, seen_edges)
        if el is not None:
            elements.append(el)
    return elements


def _collect_edge_border_elements(page, max_rows, max_cols, seen_edges: set) -> list:
    elements = []
    for he in page.get('h_edges', []):
        if '_row' not in he:
            continue
        r = min(he['_row'], max_rows)
        c = min(he['_col'], max_cols)
        ec = min(he['_end_col'], max_cols)
        if c == ec:
            continue
        bs = linewidth_to_border_style(he.get('linewidth', 0.0))
        el = _emit_rect_line(r, r + 1, c, ec,
                             {'top': True, 'bottom': False, 'left': False, 'right': False},
                             bs, seen_edges)
        if el is not None:
            elements.append(el)

    for ve in page.get('v_edges', []):
        if '_col' not in ve:
            continue
        c = min(ve['_col'], max_cols)
        r = min(ve['_row'], max_rows)
        er = min(ve['_end_row'], max_rows)
        if r == er:
            continue
        bs = linewidth_to_border_style(ve.get('linewidth', 0.0))
        el = _emit_rect_line(r, er, c, c + 1,
                             {'top': False, 'bottom': False, 'left': True, 'right': False},
                             bs, seen_edges)
        if el is not None:
            elements.append(el)
    return elements
