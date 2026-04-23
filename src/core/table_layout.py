"""テーブルセル内のテキスト要素を抽出する。

テーブル bbox と cells_2d (行×列の bbox 配列) を元に、
各セルに含まれる単語をグリッド座標に変換する。
"""

from src.core.constants import WORD_IN_BBOX_TOL
from src.core.text_layout import (
    _calc_end_col,
    _make_text_element,
    _split_into_visual_lines,
)
from src.utils.text import join_word_texts, split_by_horizontal_gap


def _find_words_in_bbox(page_words, used_ids, x0, y0, x1, y1) -> list:
    tol = WORD_IN_BBOX_TOL
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


