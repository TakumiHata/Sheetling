import pytest
from src.core.auto_layout_service import _cleanup_extracted_data


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
