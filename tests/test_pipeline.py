import pytest
from src.core.auto_layout_service import (
    _collect_content_bounds,
    _cleanup_extracted_data,
)


class TestCollectContentBounds:
    def test_basic(self):
        extracted = {
            'pages': [{
                'page_number': 1, 'width': 595, 'height': 842,
                '_content_min_x': 50.0, '_content_min_y': 30.0,
                '_content_grid_w': 10.0, '_content_grid_h': 20.0,
            }]
        }
        grid_params = {'max_cols': 47, 'max_rows': 39}
        bounds = _collect_content_bounds(extracted, grid_params)
        assert 1 in bounds
        assert bounds[1]['min_x'] == 50.0
        assert bounds[1]['page_width'] == 595.0

    def test_fallback_values(self):
        extracted = {
            'pages': [{
                'page_number': 1, 'width': 595, 'height': 842,
            }]
        }
        grid_params = {'max_cols': 47, 'max_rows': 39}
        bounds = _collect_content_bounds(extracted, grid_params)
        assert bounds[1]['min_x'] == 0.0


class TestCleanupExtractedData:
    def test_removes_temp_keys(self):
        extracted = {
            'pages': [{
                'page_number': 1,
                'table_data': [['a']], 'table_data_raw': [['a']],
                'table_row_y_positions': [[1]], 'table_cells': [[]],
                '_content_min_x': 0.0, '_content_min_y': 0.0,
                '_content_grid_w': 10.0, '_content_grid_h': 20.0,
                'words': [], 'rects': [],
            }]
        }
        _cleanup_extracted_data(extracted)
        page = extracted['pages'][0]
        assert 'table_data' not in page
        assert '_content_min_x' not in page
        assert 'words' in page
