from src.templates.prompts import GRID_SIZES
from src.utils.logger import get_logger

logger = get_logger(__name__)


def _detect_content_bounds(page: dict, page_h: float):
    all_x: list = []
    all_y: list = []
    for w in page['words']:
        t = float(w.get('top', 0))
        if t < 0 or t > page_h:
            continue
        all_x.append(float(w['x0']))
        if 'x1' in w:
            all_x.append(float(w['x1']))
        all_y.append(t)
        if 'bottom' in w:
            all_y.append(float(w['bottom']))
    for r in page['rects']:
        all_x.extend([float(r['x0']), float(r['x1'])])
        all_y.extend([float(r['top']), float(r['bottom'])])
    for cells_2d in page.get('table_cells', []):
        if not cells_2d:
            continue
        for row_cells in cells_2d:
            for cb in row_cells:
                if cb is None:
                    continue
                all_x.extend([float(cb['x0']), float(cb['x1'])])
                all_y.extend([float(cb['top']), float(cb['bottom'])])

    if all_x and all_y:
        return min(all_x), max(all_x), min(all_y), max(all_y)
    return 0.0, float(page['width']), 0.0, page_h


def _assign_word_grid_coords(page: dict, page_h: float, to_row, to_col) -> None:
    for w in page['words']:
        t = float(w.get('top', 0))
        if t < 0 or t > page_h:
            continue
        w['_row'] = to_row(t)
        w['_col'] = to_col(w['x0'])
        if w.get('is_vertical') and 'bottom' in w:
            w['_end_row'] = to_row(w['bottom'])


def _merge_thin_lines_to_rects(page: dict) -> None:
    _LINE_THICKNESS = 3.0
    _MERGE_TOL = 5.0

    h_lines = []
    v_lines = []
    for idx, r in enumerate(page['rects']):
        w = abs(r['x1'] - r['x0'])
        h = abs(r['bottom'] - r['top'])
        if w < _LINE_THICKNESS and h >= _LINE_THICKNESS:
            v_lines.append((idx, (r['x0'] + r['x1']) / 2, r['top'], r['bottom']))
        elif h < _LINE_THICKNESS and w >= _LINE_THICKNESS:
            h_lines.append((idx, r['x0'], r['x1'], (r['top'] + r['bottom']) / 2))

    used_indices, merged_rects = _find_line_rectangles(h_lines, v_lines, _MERGE_TOL)

    remaining = [page['rects'][i] for i in range(len(page['rects']))
                 if i not in used_indices]
    remaining.extend(merged_rects)
    page['rects'] = remaining


def _find_line_rectangles(h_lines, v_lines, merge_tol):
    used_indices: set = set()
    merged_rects: list = []
    for hi_top in range(len(h_lines)):
        if h_lines[hi_top][0] in used_indices:
            continue
        _, hx0_t, hx1_t, hy_t = h_lines[hi_top]
        for hi_bot in range(len(h_lines)):
            if hi_bot == hi_top or h_lines[hi_bot][0] in used_indices:
                continue
            _, hx0_b, hx1_b, hy_b = h_lines[hi_bot]
            if hy_b <= hy_t:
                continue
            if abs(hx0_t - hx0_b) > merge_tol or abs(hx1_t - hx1_b) > merge_tol:
                continue
            vl = _find_vertical_edge(v_lines, used_indices, min(hx0_t, hx0_b), hy_t, hy_b, merge_tol)
            vr = _find_vertical_edge(v_lines, used_indices, max(hx1_t, hx1_b), hy_t, hy_b, merge_tol)
            if vl is not None or vr is not None:
                used_indices.add(h_lines[hi_top][0])
                used_indices.add(h_lines[hi_bot][0])
                x0 = min(hx0_t, hx0_b)
                x1 = max(hx1_t, hx1_b)
                if vl is not None:
                    used_indices.add(v_lines[vl][0])
                    x0 = min(x0, v_lines[vl][1])
                if vr is not None:
                    used_indices.add(v_lines[vr][0])
                    x1 = max(x1, v_lines[vr][1])
                merged_rects.append({'x0': x0, 'top': hy_t, 'x1': x1, 'bottom': hy_b})
                break
    return used_indices, merged_rects


def _find_vertical_edge(v_lines, used_indices, target_x, y_top, y_bot, tol):
    for vi in range(len(v_lines)):
        if v_lines[vi][0] in used_indices:
            continue
        _, vx, vy0, vy1 = v_lines[vi]
        if (abs(vx - target_x) < tol and
                abs(vy0 - y_top) < tol and abs(vy1 - y_bot) < tol):
            return vi
    return None


def _assign_rect_grid_coords(page: dict, to_row, to_col, max_rows, max_cols) -> None:
    for r in page['rects']:
        r['_row']     = to_row(r['top'])
        r['_end_row'] = to_row(r['bottom']) + 1
        r['_col']     = to_col(r['x0'])
        r['_end_col'] = to_col(r['x1']) + 1

    tol = 3.0
    table_bboxes = page.get('table_bboxes', [])

    def is_inside_table(r: dict) -> bool:
        for bbox in table_bboxes:
            if (r['x0'] >= bbox[0] - tol and r['x1'] <= bbox[2] + tol and
                    r['top'] >= bbox[1] - tol and r['bottom'] <= bbox[3] + tol):
                return True
        return False

    page['rects'] = [r for r in page['rects'] if not is_inside_table(r)]
    for rect in page['rects']:
        rect['_borders'] = {'top': True, 'bottom': True, 'left': True, 'right': True}


def _build_table_border_rects(page: dict, to_row, to_col, max_rows, max_cols) -> None:
    table_border_rects = []
    for cells_2d in page.get('table_cells', []):
        if not cells_2d:
            continue
        for ri, row_cells in enumerate(cells_2d):
            for ci, cb in enumerate(row_cells):
                if cb is None:
                    continue
                r  = max(1, to_row(float(cb['top'])))
                er = max(r + 1, min(max_rows, to_row(float(cb['bottom']))))
                c  = max(1, to_col(float(cb['x0'])))
                ec = max(c + 1, min(max_cols, to_col(float(cb['x1']))))
                table_border_rects.append({
                    '_row': r, '_end_row': er,
                    '_col': c, '_end_col': ec,
                    '_pdf_x0': float(cb['x0']), '_pdf_top': float(cb['top']),
                    '_pdf_x1': float(cb['x1']), '_pdf_bottom': float(cb['bottom']),
                    '_borders': {'top': True, 'bottom': True, 'left': True, 'right': True},
                })
    page['table_border_rects'] = table_border_rects


def _assign_edge_grid_coords(page: dict, to_row, to_col) -> None:
    for he in page.get('h_edges', []):
        he['_row'] = to_row(float(he['y']))
        he['_col'] = to_col(float(he['x0']))
        he['_end_col'] = to_col(float(he['x1'])) + 1
    for ve in page.get('v_edges', []):
        ve['_col'] = to_col(float(ve['x']))
        ve['_row'] = to_row(float(ve['y0']))
        ve['_end_row'] = to_row(float(ve['y1'])) + 1


def compute_grid_coords(page: dict, max_rows: int, max_cols: int) -> None:
    page_h = float(page['height'])
    min_x, max_x, min_y, max_y = _detect_content_bounds(page, page_h)

    content_w = max_x - min_x
    content_h = max_y - min_y
    grid_w = content_w / max_cols if content_w > 0 else float(page['width']) / max_cols
    grid_h = content_h / max_rows if content_h > 0 else page_h / max_rows

    page['_content_min_x'] = min_x
    page['_content_min_y'] = min_y
    page['_content_grid_w'] = grid_w
    page['_content_grid_h'] = grid_h

    def to_row(y: float) -> int:
        return max(1, min(max_rows, 1 + int((float(y) - min_y) / grid_h)))

    def to_col(x: float) -> int:
        return max(1, min(max_cols, 1 + int((float(x) - min_x) / grid_w)))

    _assign_word_grid_coords(page, page_h, to_row, to_col)
    _merge_thin_lines_to_rects(page)
    _assign_rect_grid_coords(page, to_row, to_col, max_rows, max_cols)
    _build_table_border_rects(page, to_row, to_col, max_rows, max_cols)
    _assign_edge_grid_coords(page, to_row, to_col)


def setup_grid_params(first_page: dict, grid_size: str) -> dict:
    page_w = float(first_page['width'])
    page_h = float(first_page['height'])
    max_dim_pt = max(page_w, page_h)
    is_a3 = max_dim_pt > 1000
    is_landscape = page_w > page_h

    ref_key = f"{grid_size}_a3" if is_a3 else grid_size
    ref = GRID_SIZES.get(ref_key, GRID_SIZES.get(grid_size, GRID_SIZES["1pt"]))
    grid_params = dict(ref)
    grid_params['grid_size'] = grid_size
    grid_params['paper_size'] = 8 if is_a3 else 9
    grid_params['orientation'] = 'landscape' if is_landscape else 'portrait'

    if is_landscape:
        grid_params['max_cols'] = ref['max_cols_landscape']
        grid_params['max_rows'] = ref['max_rows_landscape']

    logger.debug(
        f"[grid] {grid_size} ({grid_params['orientation']}, "
        f"{'A3' if is_a3 else 'A4'}): "
        f"max_cols={grid_params['max_cols']}, max_rows={grid_params['max_rows']}, "
        f"excel_col_width={grid_params['excel_col_width']}"
    )

    return grid_params
