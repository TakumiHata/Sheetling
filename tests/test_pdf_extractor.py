import pytest
from src.parser.pdf_extractor import (
    _remove_containing_rects,
    _to_hex_color,
    _r05,
    _merge_edge_segments,
)


class TestRemoveContainingRects:
    def test_removes_outer_rect(self):
        rects = [
            {'x0': 0, 'x1': 100, 'top': 0, 'bottom': 100},
            {'x0': 10, 'x1': 50, 'top': 10, 'bottom': 50},
        ]
        result = _remove_containing_rects(rects)
        assert len(result) == 1
        assert result[0]['x0'] == 10

    def test_keeps_identical(self):
        rects = [
            {'x0': 0, 'x1': 100, 'top': 0, 'bottom': 100},
            {'x0': 0, 'x1': 100, 'top': 0, 'bottom': 100},
        ]
        result = _remove_containing_rects(rects)
        assert len(result) == 2

    def test_keeps_non_overlapping(self):
        rects = [
            {'x0': 0, 'x1': 50, 'top': 0, 'bottom': 50},
            {'x0': 60, 'x1': 100, 'top': 60, 'bottom': 100},
        ]
        result = _remove_containing_rects(rects)
        assert len(result) == 2

    def test_empty_input(self):
        assert _remove_containing_rects([]) == []


class TestToHexColor:
    def test_none(self):
        assert _to_hex_color(None) is None

    def test_grayscale_black(self):
        assert _to_hex_color(0.0) == '000000'

    def test_grayscale_white(self):
        assert _to_hex_color(1.0) == 'FFFFFF'

    def test_grayscale_mid(self):
        assert _to_hex_color(0.5) == '808080'

    def test_rgb(self):
        assert _to_hex_color((1.0, 0.0, 0.0)) == 'FF0000'
        assert _to_hex_color((0.0, 1.0, 0.0)) == '00FF00'
        assert _to_hex_color((0.0, 0.0, 1.0)) == '0000FF'

    def test_cmyk(self):
        result = _to_hex_color((0.0, 0.0, 0.0, 0.0))
        assert result == 'FFFFFF'
        result = _to_hex_color((1.0, 1.0, 1.0, 0.0))
        assert result == '000000'

    def test_invalid_type(self):
        assert _to_hex_color('red') is None


class TestR05:
    def test_rounds_to_half(self):
        assert _r05(1.0) == 1.0
        assert _r05(1.25) == 1.0  # banker's rounding: round(2.5) = 2
        assert _r05(1.3) == 1.5
        assert _r05(1.74) == 1.5
        assert _r05(1.75) == 2.0


class TestMergeEdgeSegments:
    def test_merges_adjacent(self):
        edges = [
            {'y': 100, 'x0': 0, 'x1': 50, 'linewidth': 1.0},
            {'y': 100, 'x0': 51, 'x1': 100, 'linewidth': 1.0},
        ]
        result = _merge_edge_segments(edges, 'y', 'x0', 'x1')
        assert len(result) == 1
        assert result[0]['x0'] == 0
        assert result[0]['x1'] == 100

    def test_keeps_separate(self):
        edges = [
            {'y': 100, 'x0': 0, 'x1': 50, 'linewidth': 1.0},
            {'y': 100, 'x0': 60, 'x1': 100, 'linewidth': 1.0},
        ]
        result = _merge_edge_segments(edges, 'y', 'x0', 'x1')
        assert len(result) == 2

    def test_different_axis_values(self):
        edges = [
            {'y': 100, 'x0': 0, 'x1': 50, 'linewidth': 1.0},
            {'y': 200, 'x0': 0, 'x1': 50, 'linewidth': 1.0},
        ]
        result = _merge_edge_segments(edges, 'y', 'x0', 'x1')
        assert len(result) == 2

    def test_takes_max_linewidth(self):
        edges = [
            {'y': 100, 'x0': 0, 'x1': 50, 'linewidth': 1.0},
            {'y': 100, 'x0': 51, 'x1': 100, 'linewidth': 2.0},
        ]
        result = _merge_edge_segments(edges, 'y', 'x0', 'x1')
        assert result[0]['linewidth'] == 2.0

    def test_vertical_edges(self):
        edges = [
            {'x': 50, 'y0': 0, 'y1': 100, 'linewidth': 1.0},
            {'x': 50, 'y0': 101, 'y1': 200, 'linewidth': 1.0},
        ]
        result = _merge_edge_segments(edges, 'x', 'y0', 'y1')
        assert len(result) == 1
        assert result[0]['span'] == 200
