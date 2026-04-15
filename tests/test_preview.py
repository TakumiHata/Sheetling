import os
import pytest
from src.renderer.preview import (
    _compute_cell_metrics,
    generate_border_preview,
)


class TestComputeCellMetrics:
    def test_no_content_bounds(self):
        offset_x, offset_y, cell_w, cell_h = _compute_cell_metrics(1000, 800, 50, 40, None)
        assert offset_x == 0.0
        assert offset_y == 0.0
        assert cell_w == 1000 / 50
        assert cell_h == 800 / 40

    def test_with_content_bounds(self):
        bounds = {
            'min_x': 50.0, 'min_y': 30.0,
            'grid_w': 10.0, 'grid_h': 20.0,
            'page_width': 595.0, 'page_height': 842.0,
        }
        offset_x, offset_y, cell_w, cell_h = _compute_cell_metrics(1190, 1684, 47, 39, bounds)
        assert offset_x > 0
        assert offset_y > 0
        assert cell_w > 0
        assert cell_h > 0


class TestGenerateBorderPreview:
    def test_creates_image(self, tmp_path):
        page_layout = {
            'elements': [{
                'type': 'border_rect', 'row': 2, 'end_row': 5,
                'col': 3, 'end_col': 10,
                'borders': {'top': True, 'bottom': True, 'left': True, 'right': True},
            }]
        }
        output = str(tmp_path / 'preview.png')
        generate_border_preview(page_layout, {'max_cols': 47, 'max_rows': 39}, output)
        assert os.path.exists(output)
        assert os.path.getsize(output) > 0

    def test_empty_elements(self, tmp_path):
        output = str(tmp_path / 'empty.png')
        generate_border_preview({'elements': []}, {'max_cols': 47, 'max_rows': 39}, output)
        assert os.path.exists(output)

    def test_with_content_bounds(self, tmp_path):
        page_layout = {
            'elements': [{
                'type': 'border_rect', 'row': 1, 'end_row': 3,
                'col': 1, 'end_col': 5,
                'borders': {'top': True, 'bottom': True, 'left': True, 'right': True},
            }]
        }
        bounds = {
            'min_x': 50.0, 'min_y': 30.0,
            'grid_w': 10.0, 'grid_h': 20.0,
            'page_width': 595.0, 'page_height': 842.0,
        }
        output = str(tmp_path / 'bounded.png')
        generate_border_preview(page_layout, {'max_cols': 47, 'max_rows': 39}, output,
                                content_bounds=bounds)
        assert os.path.exists(output)
