import json

from src.utils.font import normalize_font_name, linewidth_to_border_style
from src.utils.text import join_word_texts, split_by_horizontal_gap


def _make_text_element(word_group: list, row: int, col: int, end_col: int,
                       max_rows: int, is_vertical: bool = False) -> dict | None:
    content = join_word_texts([w.get('text', '') for w in word_group])
    stripped = content.strip()
    if not stripped or (len(stripped) == 1 and stripped in '!"#$%&\'()*+,-./:;<=>?@[\\]^_`{|}~'):
        return None
    elem: dict = {
        'type': 'text',
        'content': content,
        'row': min(row, max_rows),
        'col': col,
        'end_col': end_col,
    }
    first = word_group[0]
    if first.get('font_color') and first['font_color'] != '000000':
        elem['font_color'] = first['font_color']
    if first.get('font_size'):
        elem['font_size'] = first['font_size']
    fn = normalize_font_name(first.get('fontname', ''))
    if fn:
        elem['font_name'] = fn
    if is_vertical:
        elem['is_vertical'] = True
        if '_end_row' in first:
            elem['end_row'] = min(first['_end_row'], max_rows)
        elem['end_col'] = min(col + 1, end_col)
    return elem


def _split_into_visual_lines(words: list, gap: float = 3.0) -> list:
    sw = sorted(words, key=lambda w: float(w.get('top', 0)))
    vis_lines: list = [[sw[0]]]
    for w in sw[1:]:
        prev_b = float(vis_lines[-1][-1].get('bottom', vis_lines[-1][-1]['top']))
        this_t = float(w.get('top', 0))
        if this_t - prev_b > gap:
            vis_lines.append([w])
        else:
            vis_lines[-1].append(w)
    return vis_lines


def _calc_end_col(word_group: list, min_x: float, grid_w: float, col: int, max_cols: int) -> int:
    last = word_group[-1]
    x1 = float(last.get('x1', last.get('x0', 0)))
    return max(col + 1, min(max_cols, 1 + int((x1 - min_x) / grid_w)))


def _find_words_in_bbox(page_words, used_ids, x0, y0, x1, y1) -> list:
    tol = 2.0
    found = []
    for w in page_words:
        if '_row' not in w:
            continue
        wid = id(w)
        if wid in used_ids:
            continue
        wx0 = float(w.get('x0', 0))
        wy0 = float(w.get('top', 0))
        if x0 - tol <= wx0 <= x1 + tol and y0 - tol <= wy0 <= y1 + tol:
            found.append(w)
    return found


def _place_cell_words(cell_words, grid_row, to_row, to_col, y1,
                      max_rows, max_cols, min_x, grid_w) -> list:
    elements = []
    cell_max_row = max(grid_row, to_row(y1) - 1)
    cell_word_rows: dict = {}
    for w in cell_words:
        wr_clipped = min(w['_row'], cell_max_row)
        cell_word_rows.setdefault(wr_clipped, []).append(w)

    for wr, wds in sorted(cell_word_rows.items()):
        vis_lines = _split_into_visual_lines(wds)
        for vl_idx, vl_words in enumerate(vis_lines):
            vl_row = wr + vl_idx
            h_groups = split_by_horizontal_gap(
                sorted(vl_words, key=lambda x: float(x.get('x0', 0)))
            )
            for hg in h_groups:
                line_text = join_word_texts([w.get('text', '') for w in hg]).strip()
                if not line_text:
                    continue
                hg_col = max(1, to_col(float(hg[0].get('x0', 0))))
                hg_end_col = _calc_end_col(hg, min_x, grid_w, hg_col, max_cols)
                te = _make_text_element(hg, vl_row, hg_col, hg_end_col, max_rows)
                if te:
                    elements.append(te)
    return elements


def _process_table_cell(page_words, used_word_ids, cells_2d, r_idx, c_idx, trow,
                        num_cols, col_xs, row_ys, to_row, to_col,
                        max_rows, max_cols, min_x, grid_w) -> list:
    cell_content = trow[c_idx]
    if cell_content is None or c_idx >= len(col_xs) - 1:
        return []
    raw = cell_content if isinstance(cell_content, str) else ''
    lines = [ln.strip() for ln in raw.split('\n') if ln.strip()]
    if not lines:
        return []

    x0, y0, x1, y1 = _resolve_cell_bbox(cells_2d, r_idx, c_idx, trow, num_cols, col_xs, row_ys)
    grid_row = max(1, to_row(y0))
    grid_col = max(1, to_col(x0))
    grid_end_col = max(grid_col + 1, min(max_cols, to_col(x1)))

    cell_words = _find_words_in_bbox(page_words, used_word_ids, x0, y0, x1, y1)
    if cell_words:
        for w in cell_words:
            used_word_ids.add(id(w))
        return _place_cell_words(cell_words, grid_row, to_row, to_col, y1,
                                 max_rows, max_cols, min_x, grid_w)

    elements = []
    grid_end_row = max(grid_row, to_row(y1) - 1)
    for line_idx, line in enumerate(lines):
        line_row = grid_row + line_idx
        if line_row > grid_end_row:
            break
        elements.append({
            'type': 'text', 'content': line,
            'row': min(max_rows, line_row),
            'col': grid_col, 'end_col': grid_end_col,
        })
    return elements


def _table_text_elements_from_2d(page: dict, grid_params: dict) -> list:
    max_rows = grid_params['max_rows']
    max_cols = grid_params['max_cols']
    min_x = page.get('_content_min_x', 0.0)
    min_y = page.get('_content_min_y', 0.0)
    grid_h = page.get('_content_grid_h', float(page['height']) / max_rows)
    grid_w = page.get('_content_grid_w', float(page['width']) / max_cols)

    def to_row(y): return max(1, min(max_rows, 1 + int((float(y) - min_y) / grid_h)))
    def to_col(x): return max(1, min(max_cols, 1 + int((float(x) - min_x) / grid_w)))

    elements: list = []
    used_word_ids: set = set()
    table_data_src = page.get('table_data_raw') or page.get('table_data', [])
    page_words = page.get('words', [])

    for table_data, col_xs, row_ys, cells_2d in zip(
        table_data_src,
        page.get('table_col_x_positions', []),
        page.get('table_row_y_positions', []),
        page.get('table_cells', []),
    ):
        if not table_data or not col_xs or not row_ys:
            continue
        num_cols = len(table_data[0]) if table_data else 0
        for r_idx, trow in enumerate(table_data):
            if r_idx >= len(row_ys) - 1:
                continue
            for c_idx in range(len(trow)):
                elements.extend(_process_table_cell(
                    page_words, used_word_ids, cells_2d, r_idx, c_idx, trow,
                    num_cols, col_xs, row_ys, to_row, to_col,
                    max_rows, max_cols, min_x, grid_w))
    return elements


def _resolve_cell_bbox(cells_2d, r_idx, c_idx, trow, num_cols, col_xs, row_ys):
    cell_bbox = None
    if cells_2d and r_idx < len(cells_2d) and c_idx < len(cells_2d[r_idx]):
        cell_bbox = cells_2d[r_idx][c_idx]
    if cell_bbox is not None:
        return cell_bbox['x0'], cell_bbox['top'], cell_bbox['x1'], cell_bbox['bottom']
    col_end_idx = c_idx + 1
    while col_end_idx < num_cols and trow[col_end_idx] is None:
        col_end_idx += 1
    return (col_xs[c_idx], row_ys[r_idx],
            col_xs[min(col_end_idx, len(col_xs) - 1)],
            row_ys[min(r_idx + 1, len(row_ys) - 1)])


def _is_near_duplicate(seen: list, r: int, er: int, c: int, ec: int, tol: int = 0) -> bool:
    for (sr, ser, sc, sec) in seen:
        if (abs(sr - r) <= tol and abs(ser - er) <= tol and
                abs(sc - c) <= tol and abs(sec - ec) <= tol):
            return True
    return False


def _collect_table_border_elements(page, max_rows, max_cols, seen) -> list:
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
        if _is_near_duplicate(seen, r, er, c, ec):
            continue
        seen.append((r, er, c, ec))
        elements.append({
            'type': 'border_rect',
            'row': r, 'end_row': er, 'col': c, 'end_col': ec,
            'borders': tbr.get('_borders', {'top': True, 'bottom': True, 'left': True, 'right': True}),
        })
    return elements


def _collect_rect_border_elements(page, max_rows, max_cols, seen) -> list:
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
            key = (r, r + 1, c, ec)
            if not _is_near_duplicate(seen, *key):
                seen.append(key)
                elements.append({'type': 'border_rect', 'row': r, 'end_row': r + 1,
                                 'col': c, 'end_col': ec,
                                 'borders': {'top': True, 'bottom': False, 'left': False, 'right': False},
                                 'border_style': bs})
            continue
        if c == ec and r != er:
            key = (r, er, c, c + 1)
            if not _is_near_duplicate(seen, *key):
                seen.append(key)
                elements.append({'type': 'border_rect', 'row': r, 'end_row': er,
                                 'col': c, 'end_col': c + 1,
                                 'borders': {'top': False, 'bottom': False, 'left': True, 'right': False},
                                 'border_style': bs})
            continue
        if r == er and c == ec:
            continue
        if not _is_near_duplicate(seen, r, er, c, ec):
            seen.append((r, er, c, ec))
            elements.append({'type': 'border_rect', 'row': r, 'end_row': er,
                             'col': c, 'end_col': ec,
                             'borders': {'top': True, 'bottom': True, 'left': True, 'right': True},
                             'border_style': bs})
    return elements


def _collect_edge_border_elements(page, max_rows, max_cols, seen) -> list:
    elements = []
    for he in page.get('h_edges', []):
        if '_row' not in he:
            continue
        r = min(he['_row'], max_rows)
        c = min(he['_col'], max_cols)
        ec = min(he['_end_col'], max_cols)
        if c == ec:
            continue
        key = (r, r + 1, c, ec)
        if _is_near_duplicate(seen, *key):
            continue
        seen.append(key)
        bs = linewidth_to_border_style(he.get('linewidth', 0.0))
        elements.append({'type': 'border_rect', 'row': r, 'end_row': r + 1,
                         'col': c, 'end_col': ec,
                         'borders': {'top': True, 'bottom': False, 'left': False, 'right': False},
                         'border_style': bs})

    for ve in page.get('v_edges', []):
        if '_col' not in ve:
            continue
        c = min(ve['_col'], max_cols)
        r = min(ve['_row'], max_rows)
        er = min(ve['_end_row'], max_rows)
        if r == er:
            continue
        key = (r, er, c, c + 1)
        if _is_near_duplicate(seen, *key):
            continue
        seen.append(key)
        bs = linewidth_to_border_style(ve.get('linewidth', 0.0))
        elements.append({'type': 'border_rect', 'row': r, 'end_row': er,
                         'col': c, 'end_col': c + 1,
                         'borders': {'top': False, 'bottom': False, 'left': True, 'right': False},
                         'border_style': bs})
    return elements


def _build_table_cell_bboxes(page) -> list:
    bboxes = []
    table_data_src = page.get('table_data_raw') or page.get('table_data', [])
    for td, cells_2d in zip(table_data_src, page.get('table_cells', [])):
        if not td or not cells_2d:
            continue
        for ri, trow in enumerate(td):
            for ci, cell_content in enumerate(trow):
                if cell_content is None:
                    continue
                if (cells_2d and ri < len(cells_2d)
                        and ci < len(cells_2d[ri])
                        and cells_2d[ri][ci] is not None):
                    cb = cells_2d[ri][ci]
                    bboxes.append((float(cb['x0']), float(cb['top']),
                                   float(cb['x1']), float(cb['bottom'])))
    return bboxes


def _is_word_in_table(w, table_bboxes, table_cell_bboxes, tol=2.0) -> bool:
    wx = float(w.get('x0', 0))
    wy = float(w.get('top', 0))
    in_any_table = False
    for bbox in table_bboxes:
        if (bbox[0] - tol <= wx <= bbox[2] + tol and
                bbox[1] - tol <= wy <= bbox[3] + tol):
            in_any_table = True
            break
    if not in_any_table:
        return False
    for cb in table_cell_bboxes:
        if (cb[0] - tol <= wx <= cb[2] + tol and
                cb[1] - tol <= wy <= cb[3] + tol):
            return True
    return False


def _dedup_words(words: list) -> list:
    seen: set = set()
    deduped: list = []
    for w in words:
        key = (w.get('text', ''),
               round(float(w.get('top', 0)) * 2) / 2,
               round(float(w.get('x0', 0)) * 2) / 2)
        if key not in seen:
            seen.add(key)
            deduped.append(w)
    return deduped


def _collect_text_elements(page, max_rows, max_cols, min_x, grid_w) -> list:
    table_bboxes = page.get('table_bboxes', [])
    table_cell_bboxes = _build_table_cell_bboxes(page)

    groups: dict = {}
    for w in page.get('words', []):
        if '_row' not in w or '_col' not in w:
            continue
        if _is_word_in_table(w, table_bboxes, table_cell_bboxes):
            continue
        groups.setdefault((w['_row'], w['_col']), []).append(w)

    elements = []
    seen_text: set = set()

    for (row, col), words in sorted(groups.items()):
        words = _dedup_words(words)
        vis_lines = _split_into_visual_lines(words)
        row_c = min(row, max_rows)
        col_c = min(col, max_cols)

        if len(vis_lines) > 1:
            _process_multiline_group(vis_lines, row_c, col_c, max_rows, max_cols,
                                     min_x, grid_w, seen_text, elements)
            continue

        _process_single_line_group(words, row_c, col_c, max_rows, max_cols,
                                   min_x, grid_w, seen_text, elements)
    return elements


def _process_multiline_group(vis_lines, row_c, col_c, max_rows, max_cols,
                             min_x, grid_w, seen_text, elements):
    prev_row_c = row_c - 1
    for line in vis_lines:
        h_groups = split_by_horizontal_gap(line)
        word_row = line[0].get('_row', row_c)
        line_row_c = min(max_rows, max(prev_row_c + 1, word_row))
        for hg in h_groups:
            hg_col = hg[0].get('_col', col_c)
            hg_col_c = min(hg_col, max_cols)
            pos = (line_row_c, hg_col_c)
            if pos in seen_text:
                continue
            hg_end_col = _calc_end_col(hg, min_x, grid_w, hg_col_c, max_cols)
            elem = _make_text_element(hg, line_row_c, hg_col_c, hg_end_col, max_rows)
            if elem:
                seen_text.add(pos)
                elements.append(elem)
        prev_row_c = line_row_c


def _process_single_line_group(words, row_c, col_c, max_rows, max_cols,
                               min_x, grid_w, seen_text, elements):
    h_groups = split_by_horizontal_gap(words)
    for hg in h_groups:
        hg_col = hg[0].get('_col', col_c)
        hg_col_c = min(hg_col, max_cols)
        pos = (row_c, hg_col_c)
        if pos in seen_text:
            continue
        hg_end_col = _calc_end_col(hg, min_x, grid_w, hg_col_c, max_cols)
        is_vert = bool(hg[0].get('is_vertical'))
        elem = _make_text_element(hg, row_c, hg_col_c, hg_end_col, max_cols, is_vertical=is_vert)
        if elem:
            seen_text.add(pos)
            elements.append(elem)


def generate_layout(extracted_data: dict, grid_params: dict) -> str:
    max_rows = grid_params['max_rows']
    max_cols = grid_params['max_cols']

    layout = []
    for page in extracted_data.get('pages', []):
        seen_border_rects: list = []
        min_x = page.get('_content_min_x', 0.0)
        grid_w = page.get('_content_grid_w', float(page['width']) / max_cols)

        elements = []
        elements.extend(_collect_table_border_elements(page, max_rows, max_cols, seen_border_rects))
        elements.extend(_collect_rect_border_elements(page, max_rows, max_cols, seen_border_rects))
        elements.extend(_collect_edge_border_elements(page, max_rows, max_cols, seen_border_rects))

        text_elements = _collect_text_elements(page, max_rows, max_cols, min_x, grid_w)
        seen_text = {(e['row'], e['col']) for e in text_elements}
        elements.extend(text_elements)

        for tbl_elem in _table_text_elements_from_2d(page, grid_params):
            pos = (tbl_elem['row'], tbl_elem['col'])
            if pos not in seen_text:
                seen_text.add(pos)
                elements.append(tbl_elem)

        layout.append({'page_number': page['page_number'], 'elements': elements})

    return json.dumps(layout, ensure_ascii=False)
