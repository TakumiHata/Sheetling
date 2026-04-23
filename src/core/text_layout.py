"""ページ内の単語をグリッドセルに配置してテキスト要素を生成する。

テーブル外の通常テキストを担当。テーブルセル内のテキスト抽出は
table_layout.py 側で行う（共通ヘルパーはここに置く）。
"""

from src.core.constants import VISUAL_LINE_GAP, WORD_IN_TABLE_TOL
from src.utils.font import normalize_font_name
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


def _split_into_visual_lines(words: list, gap: float = VISUAL_LINE_GAP) -> list:
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


def _is_word_in_table(w, table_bboxes, table_cell_bboxes, tol=WORD_IN_TABLE_TOL) -> bool:
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
