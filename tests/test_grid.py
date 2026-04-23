import pytest
from src.core.grid import (
    _detect_content_bounds,
    _assign_word_grid_coords,
    _merge_thin_lines_to_rects,
    _find_vertical_edge,
    _build_table_border_rects,
    compute_grid_coords,
    setup_grid_params,
)


def _make_page(width=595, height=842, words=None, rects=None, table_cells=None):
    return {
        'width': width, 'height': height,
        'words': words or [],
        'rects': rects or [],
        'table_bboxes': [],
        'table_cells': table_cells or [],
        'table_col_x_positions': [],
        'table_row_y_positions': [],
        'table_data': [],
        'table_data_raw': [],
        'h_edges': [],
        'v_edges': [],
    }


class TestDetectContentBounds:
    def test_from_words(self):
        page = _make_page(words=[
            {'x0': 50, 'x1': 100, 'top': 30, 'bottom': 50},
            {'x0': 200, 'x1': 300, 'top': 100, 'bottom': 120},
        ])
        min_x, max_x, min_y, max_y = _detect_content_bounds(page, 842.0)
        assert min_x == 50
        assert max_x == 300
        assert min_y == 30
        assert max_y == 120

    def test_from_rects(self):
        page = _make_page(rects=[
            {'x0': 10, 'x1': 500, 'top': 20, 'bottom': 800},
        ])
        min_x, max_x, min_y, max_y = _detect_content_bounds(page, 842.0)
        assert min_x == 10
        assert max_x == 500

    def test_empty_falls_back_to_page(self):
        page = _make_page()
        min_x, max_x, min_y, max_y = _detect_content_bounds(page, 842.0)
        assert min_x == 0.0
        assert max_x == 595.0

    def test_skips_out_of_bounds_words(self):
        page = _make_page(words=[
            {'x0': 50, 'top': -10, 'bottom': 5},
            {'x0': 100, 'x1': 200, 'top': 50, 'bottom': 70},
        ])
        min_x, max_x, min_y, max_y = _detect_content_bounds(page, 842.0)
        assert min_x == 100
        assert min_y == 50


class TestAssignWordGridCoords:
    def test_assigns_row_col(self):
        page = _make_page(words=[
            {'x0': 50, 'top': 100, 'bottom': 120},
        ])
        to_row = lambda y: max(1, min(39, 1 + int(float(y) / 21.6)))
        to_col = lambda x: max(1, min(47, 1 + int(float(x) / 12.7)))
        _assign_word_grid_coords(page, 842.0, to_row, to_col)
        assert '_row' in page['words'][0]
        assert '_col' in page['words'][0]
        assert page['words'][0]['_row'] >= 1
        assert page['words'][0]['_col'] >= 1

    def test_vertical_word_gets_end_row(self):
        page = _make_page(words=[
            {'x0': 50, 'top': 100, 'bottom': 300, 'is_vertical': True},
        ])
        to_row = lambda y: max(1, min(39, 1 + int(float(y) / 21.6)))
        to_col = lambda x: max(1, min(47, 1 + int(float(x) / 12.7)))
        _assign_word_grid_coords(page, 842.0, to_row, to_col)
        assert '_end_row' in page['words'][0]
        assert page['words'][0]['_end_row'] > page['words'][0]['_row']


class TestMergeThinLinesToRects:
    def test_four_lines_become_rect(self):
        page = _make_page(rects=[
            {'x0': 100, 'x1': 300, 'top': 99, 'bottom': 101},   # top h-line
            {'x0': 100, 'x1': 300, 'top': 399, 'bottom': 401},  # bottom h-line
            {'x0': 99, 'x1': 101, 'top': 100, 'bottom': 400},   # left v-line
            {'x0': 299, 'x1': 301, 'top': 100, 'bottom': 400},  # right v-line
        ])
        original_count = len(page['rects'])
        _merge_thin_lines_to_rects(page)
        assert len(page['rects']) < original_count

    def test_normal_rect_unchanged(self):
        page = _make_page(rects=[
            {'x0': 100, 'x1': 300, 'top': 100, 'bottom': 400},
        ])
        _merge_thin_lines_to_rects(page)
        assert len(page['rects']) == 1


class TestFindVerticalEdge:
    def test_finds_matching(self):
        v_lines = [(0, 100.0, 50.0, 200.0)]
        result = _find_vertical_edge(v_lines, set(), 100.0, 50.0, 200.0, 5.0)
        assert result == 0

    def test_skips_used(self):
        v_lines = [(0, 100.0, 50.0, 200.0)]
        result = _find_vertical_edge(v_lines, {0}, 100.0, 50.0, 200.0, 5.0)
        assert result is None

    def test_no_match(self):
        v_lines = [(0, 500.0, 50.0, 200.0)]
        result = _find_vertical_edge(v_lines, set(), 100.0, 50.0, 200.0, 5.0)
        assert result is None


class TestComputeGridCoords:
    def test_full_pipeline(self):
        page = _make_page(
            words=[{'x0': 50, 'x1': 100, 'top': 50, 'bottom': 70}],
            rects=[{'x0': 40, 'x1': 500, 'top': 30, 'bottom': 800, 'linewidth': 1.0}],
        )
        compute_grid_coords(page, max_rows=39, max_cols=47)
        assert '_content_min_x' in page
        assert '_content_grid_w' in page
        assert '_row' in page['words'][0]
        assert '_col' in page['words'][0]
        assert 'table_border_rects' in page

    def test_grid_coords_in_range(self):
        page = _make_page(
            words=[
                {'x0': 0, 'x1': 595, 'top': 0, 'bottom': 842},
                {'x0': 300, 'x1': 350, 'top': 400, 'bottom': 420},
            ],
        )
        compute_grid_coords(page, max_rows=39, max_cols=47)
        for w in page['words']:
            if '_row' in w:
                assert 1 <= w['_row'] <= 39
                assert 1 <= w['_col'] <= 47


class TestSetupGridParams:
    def test_a4_portrait(self):
        page = {'width': 595, 'height': 842}
        params = setup_grid_params(page, '1pt')
        assert params['orientation'] == 'portrait'
        assert params['paper_size'] == 9
        assert params['max_cols'] == 53
        assert params['max_rows'] == 45

    def test_a4_landscape(self):
        page = {'width': 842, 'height': 595}
        params = setup_grid_params(page, '1pt')
        assert params['orientation'] == 'landscape'
        assert params['max_cols'] == 78
        assert params['max_rows'] == 30

    def test_a3_detection(self):
        page = {'width': 842, 'height': 1190}
        params = setup_grid_params(page, '1pt')
        assert params['paper_size'] == 8
        assert params['max_cols'] == 80

    def test_2pt_grid(self):
        page = {'width': 595, 'height': 842}
        params = setup_grid_params(page, '2pt')
        assert params['max_cols'] == 33
        assert params['excel_col_width'] == 2.74
