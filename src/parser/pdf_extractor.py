import pdfplumber
from typing import Dict, Any, Optional


def _remove_containing_rects(rects: list) -> list:
    tol = 1.0
    to_remove = set()
    for i, a in enumerate(rects):
        if i in to_remove:
            continue
        for j, b in enumerate(rects):
            if i == j or j in to_remove:
                continue
            a_contains_b = (
                a['x0'] - tol <= b['x0'] and
                a['x1'] + tol >= b['x1'] and
                a['top'] - tol <= b['top'] and
                a['bottom'] + tol >= b['bottom']
            )
            is_same = (
                abs(a['x0'] - b['x0']) < tol and
                abs(a['x1'] - b['x1']) < tol and
                abs(a['top'] - b['top']) < tol and
                abs(a['bottom'] - b['bottom']) < tol
            )
            if a_contains_b and not is_same:
                to_remove.add(i)
                break
    return [r for i, r in enumerate(rects) if i not in to_remove]


def _to_hex_color(color) -> Optional[str]:
    if color is None:
        return None
    if isinstance(color, (int, float)):
        v = int(round(float(color) * 255))
        return f"{v:02X}{v:02X}{v:02X}"
    if isinstance(color, (list, tuple)):
        if len(color) == 3:
            r, g, b = [int(round(c * 255)) for c in color]
            return f"{r:02X}{g:02X}{b:02X}"
        if len(color) == 4:
            c, m, y, k = color
            r = int(round((1 - c) * (1 - k) * 255))
            g = int(round((1 - m) * (1 - k) * 255))
            b = int(round((1 - y) * (1 - k) * 255))
            return f"{r:02X}{g:02X}{b:02X}"
    return None


TABLE_SETTINGS = {
    "snap_tolerance": 5,
    "snap_y_tolerance": 5,
    "join_tolerance": 5,
    "join_y_tolerance": 5,
    "edge_min_length": 5,
    "intersection_tolerance": 5,
    "intersection_x_tolerance": 5,
    "intersection_y_tolerance": 5,
}


def _extract_words(page) -> list:
    all_words = page.extract_words(
        extra_attrs=['fontname', 'size', 'non_stroking_color']
    )
    words = []
    for w in all_words:
        word = dict(w)
        raw_color = word.pop('non_stroking_color', None)
        hex_color = _to_hex_color(raw_color)
        if hex_color is not None:
            word['font_color'] = hex_color
        raw_size = word.pop('size', None)
        if raw_size is not None:
            word['font_size'] = round(float(raw_size), 1)
        words.append(word)
    _append_vertical_chars(page, words)
    return words


def _append_vertical_chars(page, words: list) -> None:
    non_upright = [
        c for c in page.chars
        if not c.get('upright', True) and c.get('text', '').strip()
    ]
    if not non_upright:
        return
    groups: list[list] = []
    for ch in sorted(non_upright, key=lambda c: float(c['x0'])):
        cx = (float(ch['x0']) + float(ch['x1'])) / 2
        sz = float(ch.get('size', 10))
        placed = False
        for g in groups:
            gx = (float(g[0]['x0']) + float(g[0]['x1'])) / 2
            if abs(cx - gx) < sz:
                g.append(ch)
                placed = True
                break
        if not placed:
            groups.append([ch])
    for g in groups:
        g_sorted = sorted(g, key=lambda c: float(c.get('top', c.get('y0', 0))))
        content = ''.join(c.get('text', '') for c in g_sorted)
        if not content.strip():
            continue
        first, last = g_sorted[0], g_sorted[-1]
        v_entry: dict = {
            'x0':          float(first['x0']),
            'top':         float(first.get('top', first.get('y0', 0))),
            'x1':          float(last['x1']),
            'bottom':      float(last.get('bottom', last.get('y1', 0))),
            'text':        content,
            'is_vertical': True,
        }
        raw_color = first.get('non_stroking_color')
        hex_c = _to_hex_color(raw_color)
        if hex_c:
            v_entry['font_color'] = hex_c
        raw_sz = first.get('size')
        if raw_sz:
            v_entry['font_size'] = round(float(raw_sz), 1)
        if first.get('fontname'):
            v_entry['fontname'] = first['fontname']
        words.append(v_entry)


def _extract_tables(page):
    tables = page.find_tables(table_settings=TABLE_SETTINGS)
    table_bboxes = [table.bbox for table in tables]
    table_col_x_positions = []
    table_row_y_positions = []
    table_cells = []
    for table in tables:
        try:
            valid_cells = [c for c in table.cells if c is not None]
            col_xs = sorted(set(
                [float(c[0]) for c in valid_cells] +
                [float(c[2]) for c in valid_cells]
            ))
            row_ys = sorted(set(
                [float(c[1]) for c in valid_cells] + [float(table.bbox[3])]
            ))
            table_col_x_positions.append(col_xs)
            table_row_y_positions.append(row_ys)
            cells_2d = [
                [
                    {'x0': float(c[0]), 'top': float(c[1]),
                     'x1': float(c[2]), 'bottom': float(c[3])}
                    if c is not None else None
                    for c in row.cells
                ]
                for row in table.rows
            ]
            table_cells.append(cells_2d)
        except Exception:
            table_col_x_positions.append([])
            table_row_y_positions.append([])
            table_cells.append([])

    table_data = page.extract_tables(table_settings=TABLE_SETTINGS)
    cleaned_tables = []
    for table in table_data:
        cleaned_table = []
        for row in table:
            cleaned_row = [cell.replace('\n', ' ') if isinstance(cell, str) else cell for cell in row]
            cleaned_table.append(cleaned_row)
        cleaned_tables.append(cleaned_table)

    return tables, table_bboxes, table_col_x_positions, table_row_y_positions, table_cells, table_data, cleaned_tables


def _filter_page_boundary_tables(tables, table_bboxes, table_col_x_positions,
                                 table_row_y_positions, table_cells,
                                 table_data, cleaned_tables, page_area):
    valid_indices = []
    for ti, tbl in enumerate(tables):
        bbox = tbl.bbox
        tbl_area = (bbox[2] - bbox[0]) * (bbox[3] - bbox[1])
        cell_count = sum(1 for c in tbl.cells if c is not None)
        if tbl_area >= 0.80 * page_area and cell_count <= 4:
            continue
        valid_indices.append(ti)
    return (
        [tables[i] for i in valid_indices],
        [tables[i].bbox for i in range(len([tables[i] for i in valid_indices]))],
        [table_col_x_positions[i] for i in valid_indices],
        [table_row_y_positions[i] for i in valid_indices],
        [table_cells[i] for i in valid_indices],
        [table_data[i] for i in valid_indices],
        [cleaned_tables[i] for i in valid_indices],
    )


def _extract_rects(page, page_area: float) -> list:
    rects = []
    for r in page.rects:
        rect_area = (r['x1'] - r['x0']) * (r['bottom'] - r['top'])
        if rect_area < 0.80 * page_area:
            rects.append({
                'x0': float(r['x0']),
                'top': float(r['top']),
                'x1': float(r['x1']),
                'bottom': float(r['bottom']),
                'linewidth': float(r.get('linewidth', 0) or 0),
            })
    return _remove_containing_rects(rects)


def _r05(v: float) -> float:
    return round(v * 2) / 2


def _collect_raw_edges(page, page_area: float):
    h_seen: set = set()
    v_seen: set = set()
    h_edges: list = []
    v_edges: list = []

    def add_h(x0, x1, y, linewidth=0.0):
        key = (_r05(min(x0, x1)), _r05(max(x0, x1)), _r05(y))
        if key not in h_seen:
            h_seen.add(key)
            h_edges.append({'x0': key[0], 'x1': key[1], 'y': key[2], 'linewidth': linewidth})

    def add_v(x, y0, y1, linewidth=0.0):
        key = (_r05(x), _r05(min(y0, y1)), _r05(max(y0, y1)))
        if key not in v_seen:
            v_seen.add(key)
            v_edges.append({'x': key[0], 'y0': key[1], 'y1': key[2],
                            'span': key[2] - key[1], 'linewidth': linewidth})

    for line in page.lines:
        lx0, lx1 = float(line['x0']), float(line['x1'])
        lt = float(line.get('top', line.get('y0', 0)))
        lb = float(line.get('bottom', line.get('y1', lt)))
        lw = float(line.get('linewidth', 0) or 0)
        if abs(lb - lt) < 2.0 and abs(lx1 - lx0) > 2.0:
            add_h(lx0, lx1, (lt + lb) / 2, lw)
        elif abs(lx1 - lx0) < 2.0 and abs(lb - lt) > 2.0:
            add_v((lx0 + lx1) / 2, lt, lb, lw)

    for r in page.rects:
        rect_area = (r['x1'] - r['x0']) * (r['bottom'] - r['top'])
        if rect_area >= 0.85 * page_area:
            continue
        rw = float(r['x1']) - float(r['x0'])
        rh = float(r['bottom']) - float(r['top'])
        is_line_like = rh < 2.0 or rw < 2.0
        has_stroke = r.get('stroking_color') is not None
        if not has_stroke and not is_line_like:
            continue
        rx0, rx1 = float(r['x0']), float(r['x1'])
        rt, rb = float(r['top']), float(r['bottom'])
        rlw = float(r.get('linewidth', 0) or 0)
        add_h(rx0, rx1, rt, rlw)
        add_h(rx0, rx1, rb, rlw)
        add_v(rx0, rt, rb, rlw)
        add_v(rx1, rt, rb, rlw)

    return h_edges, v_edges


def _merge_edge_segments(edges: list, axis_key: str, start_key: str, end_key: str) -> list:
    gap_tol = 2.0
    by_axis: dict = {}
    for e in edges:
        by_axis.setdefault(e[axis_key], []).append(
            (e[start_key], e[end_key], e.get('linewidth', 0.0))
        )
    merged = []
    for axis_val, segs in by_axis.items():
        segs_sorted = sorted(segs)
        current = [list(segs_sorted[0])]
        for s0, s1, lw in segs_sorted[1:]:
            if s0 <= current[-1][1] + gap_tol:
                current[-1][1] = max(current[-1][1], s1)
                current[-1][2] = max(current[-1][2], lw)
            else:
                current.append([s0, s1, lw])
        for s0, s1, lw in current:
            entry = {axis_key: axis_val, start_key: s0, end_key: s1, 'linewidth': lw}
            if axis_key == 'x':
                entry['span'] = s1 - s0
            merged.append(entry)
    return merged


def _extract_edges(page, page_area: float):
    h_edges, v_edges = _collect_raw_edges(page, page_area)
    v_edges = _merge_edge_segments(v_edges, 'x', 'y0', 'y1')
    h_edges = _merge_edge_segments(h_edges, 'y', 'x0', 'x1')
    return h_edges, v_edges


def extract_pdf_data(pdf_path: str) -> Dict[str, Any]:
    extracted_pages = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            words = _extract_words(page)

            width = page.width
            height = page.height
            page_area = float(width) * float(height)

            (tables, table_bboxes, table_col_x_positions,
             table_row_y_positions, table_cells,
             table_data, cleaned_tables) = _extract_tables(page)

            (tables, table_bboxes, table_col_x_positions,
             table_row_y_positions, table_cells,
             table_data, cleaned_tables) = _filter_page_boundary_tables(
                tables, table_bboxes, table_col_x_positions,
                table_row_y_positions, table_cells,
                table_data, cleaned_tables, page_area)

            rects = _extract_rects(page, page_area)
            h_edges, v_edges = _extract_edges(page, page_area)

            extracted_pages.append({
                "page_number": page_number,
                "width": float(width),
                "height": float(height),
                "words": words,
                "table_bboxes": table_bboxes,
                "table_col_x_positions": table_col_x_positions,
                "table_row_y_positions": table_row_y_positions,
                "table_cells": table_cells,
                "table_data": cleaned_tables,
                "table_data_raw": table_data,
                "rects": rects,
                "h_edges": h_edges,
                "v_edges": v_edges,
            })

    return {"pages": extracted_pages}
