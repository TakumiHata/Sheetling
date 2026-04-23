"""レイアウト生成のオーケストレータ。

実処理は以下に分散：
  - text_layout: 通常テキスト配置
  - table_layout: テーブルセル内テキスト
  - border_layout: 罫線収集
"""

import json

from src.core.border_layout import (
    _collect_edge_border_elements,
    _collect_rect_border_elements,
    _collect_table_border_elements,
    _edges_of_side,
    _emit_rect_line,
    _filter_sides_by_seen,
)
from src.core.table_layout import (
    _find_words_in_bbox,
    _place_cell_words,
    _process_table_cell,
    _resolve_cell_bbox,
    _table_text_elements_from_2d,
)
from src.core.text_layout import (
    _build_table_cell_bboxes,
    _calc_end_col,
    _collect_text_elements,
    _dedup_words,
    _is_word_in_table,
    _make_text_element,
    _process_multiline_group,
    _process_single_line_group,
    _split_into_visual_lines,
)

__all__ = [
    '_build_table_cell_bboxes', '_calc_end_col', '_collect_edge_border_elements',
    '_collect_rect_border_elements', '_collect_table_border_elements',
    '_collect_text_elements', '_dedup_words', '_edges_of_side', '_emit_rect_line',
    '_filter_sides_by_seen', '_find_words_in_bbox', '_is_word_in_table',
    '_make_text_element', '_place_cell_words', '_process_multiline_group',
    '_process_single_line_group', '_process_table_cell', '_resolve_cell_bbox',
    '_split_into_visual_lines', '_table_text_elements_from_2d',
    'generate_layout',
]


def generate_layout(extracted_data: dict, grid_params: dict) -> str:
    max_rows = grid_params['max_rows']
    max_cols = grid_params['max_cols']

    layout = []
    for page in extracted_data.get('pages', []):
        seen_edges: set = set()
        min_x = page.get('_content_min_x', 0.0)
        grid_w = page.get('_content_grid_w', float(page['width']) / max_cols)

        elements = []
        elements.extend(_collect_table_border_elements(page, max_rows, max_cols, seen_edges))
        elements.extend(_collect_rect_border_elements(page, max_rows, max_cols, seen_edges))
        elements.extend(_collect_edge_border_elements(page, max_rows, max_cols, seen_edges))

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
