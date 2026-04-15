import json
import os
import pytest
from pathlib import Path
from src.core.pipeline import (
    SheetlingPipeline,
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


class TestApplyCorrections:
    @pytest.fixture
    def setup_pipeline(self, tmp_path):
        layout = [{
            'page_number': 1,
            'elements': [
                {'type': 'border_rect', 'row': 2, 'end_row': 5, 'col': 3, 'end_col': 10,
                 'borders': {'top': True, 'bottom': True, 'left': True, 'right': True}},
                {'type': 'text', 'content': 'Hello', 'row': 3, 'col': 5, 'end_col': 10},
            ]
        }]
        layout_path = tmp_path / 'test_1pt_layout.json'
        layout_path.write_text(json.dumps(layout), encoding='utf-8')
        pipeline = SheetlingPipeline(str(tmp_path))
        return pipeline, tmp_path

    def test_add_border(self, setup_pipeline):
        pipeline, tmp_path = setup_pipeline
        corrections = json.dumps({'corrections': [
            {'action': 'add_border', 'page': 1, 'row': 10, 'end_row': 15, 'col': 5, 'end_col': 20,
             'borders': {'top': True, 'bottom': True, 'left': True, 'right': True}}
        ]})
        pipeline.apply_corrections('test', corrections,
                                   specific_out_dir=str(tmp_path),
                                   layout_json_name='test_1pt_layout.json')
        layout = json.loads((tmp_path / 'test_1pt_layout.json').read_text())
        borders = [e for e in layout[0]['elements'] if e['type'] == 'border_rect']
        assert len(borders) == 2

    def test_remove_border(self, setup_pipeline):
        pipeline, tmp_path = setup_pipeline
        corrections = json.dumps({'corrections': [
            {'action': 'remove_border', 'page': 1, 'row': 2, 'end_row': 5, 'col': 3, 'end_col': 10}
        ]})
        pipeline.apply_corrections('test', corrections,
                                   specific_out_dir=str(tmp_path),
                                   layout_json_name='test_1pt_layout.json')
        layout = json.loads((tmp_path / 'test_1pt_layout.json').read_text())
        borders = [e for e in layout[0]['elements'] if e['type'] == 'border_rect']
        assert len(borders) == 0

    def test_add_text(self, setup_pipeline):
        pipeline, tmp_path = setup_pipeline
        corrections = json.dumps({'corrections': [
            {'action': 'add_text', 'page': 1, 'row': 8, 'col': 2, 'content': 'New text'}
        ]})
        pipeline.apply_corrections('test', corrections,
                                   specific_out_dir=str(tmp_path),
                                   layout_json_name='test_1pt_layout.json')
        layout = json.loads((tmp_path / 'test_1pt_layout.json').read_text())
        texts = [e for e in layout[0]['elements'] if e['type'] == 'text']
        assert any(t['content'] == 'New text' for t in texts)

    def test_fix_text(self, setup_pipeline):
        pipeline, tmp_path = setup_pipeline
        corrections = json.dumps({'corrections': [
            {'action': 'fix_text', 'page': 1, 'row': 3, 'col': 5, 'new_row': 4, 'new_col': 6}
        ]})
        pipeline.apply_corrections('test', corrections,
                                   specific_out_dir=str(tmp_path),
                                   layout_json_name='test_1pt_layout.json')
        layout = json.loads((tmp_path / 'test_1pt_layout.json').read_text())
        hello = [e for e in layout[0]['elements'] if e.get('content') == 'Hello'][0]
        assert hello['row'] == 4
        assert hello['col'] == 6

    def test_invalid_json_raises(self, setup_pipeline):
        pipeline, tmp_path = setup_pipeline
        with pytest.raises(ValueError):
            pipeline.apply_corrections('test', 'not json',
                                       specific_out_dir=str(tmp_path),
                                       layout_json_name='test_1pt_layout.json')

    def test_missing_layout_raises(self):
        pipeline = SheetlingPipeline('/tmp/nonexistent')
        with pytest.raises(FileNotFoundError):
            pipeline.apply_corrections('missing', '{"corrections": []}')
