import json
import pytest
from src.core.layout import (
    _make_text_element,
    _split_into_visual_lines,
    _calc_end_col,
    _find_words_in_bbox,
    _is_near_duplicate,
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


class TestIsNearDuplicate:
    def test_exact_match(self):
        seen = [(1, 5, 2, 10)]
        assert _is_near_duplicate(seen, 1, 5, 2, 10) is True

    def test_no_match(self):
        seen = [(1, 5, 2, 10)]
        assert _is_near_duplicate(seen, 10, 20, 15, 25) is False

    def test_empty_seen(self):
        assert _is_near_duplicate([], 1, 5, 2, 10) is False

    def test_tolerance(self):
        seen = [(1, 5, 2, 10)]
        assert _is_near_duplicate(seen, 2, 6, 3, 11, tol=1) is True
        assert _is_near_duplicate(seen, 2, 6, 3, 11, tol=0) is False


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
