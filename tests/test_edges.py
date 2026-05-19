import pytest

from src.core.edges import (
    decompose_to_cell_edges,
    filter_short_runs,
    group_into_runs,
    runs_to_border_rects,
)


def _full_rect(r, er, c, ec):
    return {
        'type': 'border_rect', 'row': r, 'end_row': er, 'col': c, 'end_col': ec,
        'borders': {'top': True, 'bottom': True, 'left': True, 'right': True},
    }


class TestDecomposeToCellEdges:
    def test_full_rect_yields_perimeter_edges(self):
        elements = [_full_rect(2, 4, 3, 6)]
        edges, _ = decompose_to_cell_edges(elements)
        assert ('H', 2, 3) in edges
        assert ('H', 4, 5) in edges
        assert ('V', 2, 3) in edges
        assert ('V', 3, 6) in edges
        assert len(edges) == 10

    def test_partial_borders(self):
        elem = {
            'type': 'border_rect', 'row': 1, 'end_row': 1, 'col': 2, 'end_col': 5,
            'borders': {'top': True, 'bottom': False, 'left': False, 'right': False},
        }
        edges, _ = decompose_to_cell_edges([elem])
        assert edges == {('H', 1, 2), ('H', 1, 3), ('H', 1, 4)}

    def test_ignores_non_border(self):
        elements = [{'type': 'text', 'row': 1, 'col': 1, 'content': 'x'}]
        edges, _ = decompose_to_cell_edges(elements)
        assert edges == set()


class TestGroupIntoRuns:
    def test_horizontal_contiguous_become_one_run(self):
        edges = {('H', 5, c) for c in range(3, 8)}
        styles = {e: 'thin' for e in edges}
        runs = group_into_runs(edges, styles)
        assert len(runs) == 1
        r = runs[0]
        assert r['type'] == 'H' and r['row'] == 5
        assert r['col_start'] == 3 and r['col_end'] == 8

    def test_horizontal_gap_yields_two_runs(self):
        edges = {('H', 5, 3), ('H', 5, 4), ('H', 5, 7), ('H', 5, 8)}
        styles = {e: 'thin' for e in edges}
        runs = group_into_runs(edges, styles)
        assert len(runs) == 2

    def test_vertical_run(self):
        edges = {('V', r, 4) for r in range(2, 6)}
        styles = {e: 'thin' for e in edges}
        runs = group_into_runs(edges, styles)
        assert len(runs) == 1
        r = runs[0]
        assert r['type'] == 'V' and r['col'] == 4
        assert r['row_start'] == 2 and r['row_end'] == 6


class TestRunsToBorderRects:
    def test_h_run_produces_top_only_zero_height_rect(self):
        rects = runs_to_border_rects([
            {'type': 'H', 'row': 5, 'col_start': 3, 'col_end': 8, 'border_style': 'thin'}
        ])
        assert len(rects) == 1
        r = rects[0]
        assert r['row'] == 5 and r['end_row'] == 5
        assert r['col'] == 3 and r['end_col'] == 8
        assert r['borders']['top'] is True
        assert r['borders']['bottom'] is False

    def test_v_run_produces_left_only_zero_width_rect(self):
        rects = runs_to_border_rects([
            {'type': 'V', 'col': 4, 'row_start': 2, 'row_end': 6, 'border_style': 'thin'}
        ])
        assert len(rects) == 1
        r = rects[0]
        assert r['col'] == 4 and r['end_col'] == 4
        assert r['borders']['left'] is True


class TestFilterShortRuns:
    def test_removes_single_cell_h_span(self):
        # H ライン: col 3〜3 (inclusive span=1, exclusive span=2) → 除去される
        elem = {
            'type': 'border_rect', 'row': 5, 'end_row': 5, 'col': 3, 'end_col': 4,
            'borders': {'top': True, 'bottom': False, 'left': False, 'right': False},
        }
        elements = [elem]
        result = filter_short_runs(elements, min_h_span=2, min_v_span=2)
        assert not any(e.get('type') == 'border_rect' for e in result)

    def test_keeps_sufficient_h_span(self):
        # H ライン: col 3〜5 (inclusive span=3, exclusive span=4) → 保持される
        elem = {
            'type': 'border_rect', 'row': 5, 'end_row': 5, 'col': 3, 'end_col': 6,
            'borders': {'top': True, 'bottom': False, 'left': False, 'right': False},
        }
        elements = [elem]
        result = filter_short_runs(elements, min_h_span=2, min_v_span=2)
        assert any(e.get('type') == 'border_rect' for e in result)

    def test_preserves_text_elements(self):
        elements = [
            _full_rect(2, 4, 3, 6),
            {'type': 'text', 'content': 'hi', 'row': 3, 'col': 4, 'end_col': 6},
        ]
        result = filter_short_runs(elements, min_h_span=2, min_v_span=2)
        texts = [e for e in result if e.get('type') == 'text']
        assert len(texts) == 1
