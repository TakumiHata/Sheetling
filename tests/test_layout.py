import json
import pytest
from src.core.layout import (
    _make_text_element,
    _split_into_visual_lines,
    _calc_end_col,
    _find_words_in_bbox,
    _edges_of_side,
    _filter_sides_by_seen,
    _dedup_words,
    _resolve_cell_bbox,
    generate_layout,
)


class TestMakeTextElement:
    def test_basic(self):
        words = [{'text': 'Hello', 'x0': 0}]
        elem = _make_text_element(words, row=5, col=3, end_col=10, max_rows=39)
        assert elem['type'] == 'text'
        assert elem['content'] == 'Hello'
        assert elem['row'] == 5
        assert elem['col'] == 3

    def test_with_font_info(self):
        words = [{'text': 'Test', 'font_color': 'FF0000', 'font_size': 12, 'fontname': 'MSGothic'}]
        elem = _make_text_element(words, row=1, col=1, end_col=5, max_rows=39)
        assert elem['font_color'] == 'FF0000'
        assert elem['font_size'] == 12
        assert elem['font_name'] == 'MS Gothic'

    def test_black_font_color_excluded(self):
        words = [{'text': 'Test', 'font_color': '000000'}]
        elem = _make_text_element(words, row=1, col=1, end_col=5, max_rows=39)
        assert 'font_color' not in elem

    def test_punctuation_only_returns_none(self):
        words = [{'text': '.'}]
        assert _make_text_element(words, 1, 1, 5, 39) is None

    def test_empty_returns_none(self):
        words = [{'text': ' '}]
        assert _make_text_element(words, 1, 1, 5, 39) is None

    def test_vertical(self):
        words = [{'text': '縦', 'is_vertical': True, '_end_row': 10}]
        elem = _make_text_element(words, 3, 5, 8, 39, is_vertical=True)
        assert elem['is_vertical'] is True
        assert elem['end_row'] == 10

    def test_row_clamped_to_max(self):
        words = [{'text': 'Test'}]
        elem = _make_text_element(words, row=100, col=1, end_col=5, max_rows=39)
        assert elem['row'] == 39


class TestSplitIntoVisualLines:
    def test_single_line(self):
        words = [
            {'top': 100, 'bottom': 110},
            {'top': 101, 'bottom': 111},
        ]
        result = _split_into_visual_lines(words)
        assert len(result) == 1

    def test_two_lines(self):
        words = [
            {'top': 100, 'bottom': 110},
            {'top': 120, 'bottom': 130},
        ]
        result = _split_into_visual_lines(words)
        assert len(result) == 2

    def test_custom_gap(self):
        words = [
            {'top': 100, 'bottom': 110},
            {'top': 114, 'bottom': 124},
        ]
        assert len(_split_into_visual_lines(words, gap=10.0)) == 1
        assert len(_split_into_visual_lines(words, gap=2.0)) == 2


class TestCalcEndCol:
    def test_basic(self):
        words = [{'x0': 0, 'x1': 100}]
        result = _calc_end_col(words, min_x=0, grid_w=10, col=1, max_cols=47)
        assert result == 11

    def test_clamps_to_max(self):
        words = [{'x0': 0, 'x1': 10000}]
        result = _calc_end_col(words, min_x=0, grid_w=10, col=1, max_cols=47)
        assert result == 47

    def test_minimum_one_more_than_col(self):
        words = [{'x0': 0, 'x1': 5}]
        result = _calc_end_col(words, min_x=0, grid_w=100, col=1, max_cols=47)
        assert result >= 2


class TestFindWordsInBbox:
    def test_finds_inside(self):
        words = [
            {'x0': 50, 'top': 50, '_row': 1},
            {'x0': 200, 'top': 200, '_row': 5},
        ]
        found = _find_words_in_bbox(words, set(), x0=40, y0=40, x1=100, y1=100)
        assert len(found) == 1
        assert found[0]['x0'] == 50

    def test_skips_used(self):
        w = {'x0': 50, 'top': 50, '_row': 1}
        words = [w]
        found = _find_words_in_bbox(words, {id(w)}, x0=40, y0=40, x1=100, y1=100)
        assert len(found) == 0

    def test_skips_no_row(self):
        words = [{'x0': 50, 'top': 50}]
        found = _find_words_in_bbox(words, set(), x0=40, y0=40, x1=100, y1=100)
        assert len(found) == 0


class TestEdgesOfSide:
    def test_top(self):
        assert _edges_of_side(2, 5, 3, 7, 'top') == {('H', 2, 3), ('H', 2, 4), ('H', 2, 5), ('H', 2, 6)}

    def test_bottom(self):
        assert _edges_of_side(2, 5, 3, 7, 'bottom') == {('H', 5, 3), ('H', 5, 4), ('H', 5, 5), ('H', 5, 6)}

    def test_left(self):
        assert _edges_of_side(2, 5, 3, 7, 'left') == {('V', 2, 3), ('V', 3, 3), ('V', 4, 3)}

    def test_right(self):
        assert _edges_of_side(2, 5, 3, 7, 'right') == {('V', 2, 7), ('V', 3, 7), ('V', 4, 7)}

    def test_unknown_side(self):
        assert _edges_of_side(2, 5, 3, 7, 'bogus') == set()


class TestFilterSidesBySeen:
    def test_all_new_keeps_all_sides(self):
        seen: set = set()
        sides = _filter_sides_by_seen(1, 3, 1, 4,
            {'top': True, 'bottom': True, 'left': True, 'right': True}, seen)
        assert sides == {'top': True, 'bottom': True, 'left': True, 'right': True}
        assert seen  # populated

    def test_exact_duplicate_drops_rect(self):
        seen: set = set()
        _filter_sides_by_seen(1, 3, 1, 4,
            {'top': True, 'bottom': True, 'left': True, 'right': True}, seen)
        # Second identical call is fully redundant
        result = _filter_sides_by_seen(1, 3, 1, 4,
            {'top': True, 'bottom': True, 'left': True, 'right': True}, seen)
        assert result is None

    def test_nested_rect_keeps_only_internal_divider(self):
        seen: set = set()
        # Outer rect row=2..3, col=1..40
        _filter_sides_by_seen(2, 3, 1, 40,
            {'top': True, 'bottom': True, 'left': True, 'right': True}, seen)
        # Inner rect row=2..3, col=1..9 — shares top/bottom/left with outer,
        # but its right (col=9) is a NEW internal divider.
        sides = _filter_sides_by_seen(2, 3, 1, 9,
            {'top': True, 'bottom': True, 'left': True, 'right': True}, seen)
        assert sides == {'top': False, 'bottom': False, 'left': False, 'right': True}

    def test_edge_after_table_border_skipped(self):
        seen: set = set()
        # table_border_rect covers all 4 sides
        _filter_sides_by_seen(5, 8, 2, 10,
            {'top': True, 'bottom': True, 'left': True, 'right': True}, seen)
        # An h_edge coming in later at the same top line is fully redundant
        sides = _filter_sides_by_seen(5, 6, 2, 10,
            {'top': True, 'bottom': False, 'left': False, 'right': False}, seen)
        assert sides is None

    def test_false_side_ignored(self):
        seen: set = set()
        sides = _filter_sides_by_seen(1, 3, 1, 4,
            {'top': False, 'bottom': True, 'left': False, 'right': False}, seen)
        assert sides == {'top': False, 'bottom': True, 'left': False, 'right': False}


class TestDedupWords:
    def test_removes_duplicates(self):
        words = [
            {'text': 'Hello', 'top': 100.0, 'x0': 50.0},
            {'text': 'Hello', 'top': 100.0, 'x0': 50.0},
        ]
        result = _dedup_words(words)
        assert len(result) == 1

    def test_keeps_different(self):
        words = [
            {'text': 'Hello', 'top': 100.0, 'x0': 50.0},
            {'text': 'World', 'top': 100.0, 'x0': 50.0},
        ]
        result = _dedup_words(words)
        assert len(result) == 2


class TestResolveCellBbox:
    def test_from_cells_2d(self):
        cells_2d = [[{'x0': 10, 'top': 20, 'x1': 100, 'bottom': 50}]]
        x0, y0, x1, y1 = _resolve_cell_bbox(cells_2d, 0, 0, ['text'], 1, [10, 100], [20, 50])
        assert (x0, y0, x1, y1) == (10, 20, 100, 50)

    def test_fallback_to_col_row_positions(self):
        x0, y0, x1, y1 = _resolve_cell_bbox([], 0, 0, ['text'], 1, [10, 100], [20, 50])
        assert x0 == 10
        assert y0 == 20
        assert x1 == 100
        assert y1 == 50


class TestGenerateLayout:
    def test_produces_valid_json(self):
        extracted = {
            'pages': [{
                'page_number': 1, 'width': 595, 'height': 842,
                'words': [{'x0': 50, 'x1': 100, 'top': 50, 'bottom': 70,
                           '_row': 3, '_col': 5, 'text': 'Test'}],
                'rects': [], 'table_border_rects': [],
                'table_bboxes': [], 'table_cells': [], 'table_data': [],
                'table_data_raw': [], 'table_col_x_positions': [],
                'table_row_y_positions': [],
                'h_edges': [], 'v_edges': [],
                '_content_min_x': 0, '_content_grid_w': 12.7,
            }]
        }
        result = generate_layout(extracted, {'max_rows': 39, 'max_cols': 47})
        layout = json.loads(result)
        assert len(layout) == 1
        assert layout[0]['page_number'] == 1
        text_elems = [e for e in layout[0]['elements'] if e['type'] == 'text']
        assert len(text_elems) >= 1
        assert text_elems[0]['content'] == 'Test'

    def test_border_rect_from_table(self):
        extracted = {
            'pages': [{
                'page_number': 1, 'width': 595, 'height': 842,
                'words': [], 'rects': [],
                'table_border_rects': [
                    {'_row': 2, '_end_row': 5, '_col': 3, '_end_col': 10,
                     '_borders': {'top': True, 'bottom': True, 'left': True, 'right': True}}
                ],
                'table_bboxes': [], 'table_cells': [], 'table_data': [],
                'table_data_raw': [], 'table_col_x_positions': [],
                'table_row_y_positions': [],
                'h_edges': [], 'v_edges': [],
                '_content_min_x': 0, '_content_grid_w': 12.7,
            }]
        }
        result = generate_layout(extracted, {'max_rows': 39, 'max_cols': 47})
        layout = json.loads(result)
        border_elems = [e for e in layout[0]['elements'] if e['type'] == 'border_rect']
        assert len(border_elems) == 1
        assert border_elems[0]['row'] == 2
        assert border_elems[0]['end_col'] == 10

    def test_border_dedup_across_sources(self):
        """table_border_rect + h_edge at same top line → only the table_border remains."""
        extracted = {
            'pages': [{
                'page_number': 1, 'width': 595, 'height': 842,
                'words': [], 'rects': [],
                'table_border_rects': [
                    {'_row': 2, '_end_row': 5, '_col': 3, '_end_col': 10,
                     '_borders': {'top': True, 'bottom': True, 'left': True, 'right': True}}
                ],
                'table_bboxes': [], 'table_cells': [], 'table_data': [],
                'table_data_raw': [], 'table_col_x_positions': [],
                'table_row_y_positions': [],
                'h_edges': [
                    {'_row': 2, '_col': 3, '_end_col': 10, 'linewidth': 0.5},
                ],
                'v_edges': [
                    {'_row': 2, '_end_row': 5, '_col': 3, 'linewidth': 0.5},
                ],
                '_content_min_x': 0, '_content_grid_w': 12.7,
            }]
        }
        result = generate_layout(extracted, {'max_rows': 39, 'max_cols': 47})
        layout = json.loads(result)
        border_elems = [e for e in layout[0]['elements'] if e['type'] == 'border_rect']
        # Only the table_border_rect should survive; the h_edge + v_edge are fully covered.
        assert len(border_elems) == 1
        assert border_elems[0]['borders'] == {
            'top': True, 'bottom': True, 'left': True, 'right': True,
        }

    def test_border_nested_emits_internal_divider_only(self):
        """Nested table_border: inner rect contributes only its non-shared side."""
        extracted = {
            'pages': [{
                'page_number': 1, 'width': 595, 'height': 842,
                'words': [], 'rects': [],
                'table_border_rects': [
                    {'_row': 2, '_end_row': 3, '_col': 1, '_end_col': 40,
                     '_borders': {'top': True, 'bottom': True, 'left': True, 'right': True}},
                    {'_row': 2, '_end_row': 3, '_col': 1, '_end_col': 9,
                     '_borders': {'top': True, 'bottom': True, 'left': True, 'right': True}},
                ],
                'table_bboxes': [], 'table_cells': [], 'table_data': [],
                'table_data_raw': [], 'table_col_x_positions': [],
                'table_row_y_positions': [],
                'h_edges': [], 'v_edges': [],
                '_content_min_x': 0, '_content_grid_w': 12.7,
            }]
        }
        result = generate_layout(extracted, {'max_rows': 39, 'max_cols': 47})
        layout = json.loads(result)
        border_elems = [e for e in layout[0]['elements'] if e['type'] == 'border_rect']
        # Outer keeps all 4 sides; inner keeps only its right edge (internal divider at col=9).
        assert len(border_elems) == 2
        assert border_elems[1]['borders'] == {
            'top': False, 'bottom': False, 'left': False, 'right': True,
        }
