import pytest

from src.core.edges import (
    apply_edge_corrections,
    decompose_to_cell_edges,
    enumerate_runs_with_ids,
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
        # top: H,2,3 / 2,4 / 2,5
        # bottom: H,4,3 / 4,4 / 4,5
        # left: V,2,3 / 3,3
        # right: V,2,6 / 3,6
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


class TestEnumerateRunsWithIds:
    def test_full_rect_yields_4_runs(self):
        runs = enumerate_runs_with_ids([_full_rect(2, 4, 3, 6)])
        # top H, bottom H, left V, right V
        assert len(runs) == 4
        assert all('id' in r for r in runs)
        assert {r['id'] for r in runs} == {1, 2, 3, 4}

    def test_round_trip_preserves_edge_set(self):
        original = [_full_rect(2, 4, 3, 6), _full_rect(10, 12, 5, 8)]
        runs = enumerate_runs_with_ids(original)
        reconstructed = runs_to_border_rects(
            [{k: v for k, v in r.items() if k != 'id'} for r in runs])
        orig_edges, _ = decompose_to_cell_edges(original)
        new_edges, _ = decompose_to_cell_edges(reconstructed)
        assert orig_edges == new_edges


class TestApplyEdgeCorrections:
    def test_remove_by_id_drops_only_that_edge(self):
        elements = [_full_rect(2, 4, 3, 6)]
        runs = enumerate_runs_with_ids(elements)
        id_map = {r['id']: r for r in runs}
        # remove the top edge (which run has it depends on enumeration order)
        top_id = next(r['id'] for r in runs
                      if r['type'] == 'H' and r['row'] == 2)
        apply_edge_corrections(elements, [top_id], [], id_map)

        edges, _ = decompose_to_cell_edges(elements)
        # top edges (H, 2, *) gone, others remain
        assert not any(e[0] == 'H' and e[1] == 2 for e in edges)
        assert ('H', 4, 3) in edges
        assert ('V', 2, 3) in edges

    def test_add_h_edge_appears_in_layout(self):
        elements = []
        apply_edge_corrections(elements, [], [
            {'type': 'H', 'row': 7, 'col_start': 2, 'col_end': 9}
        ], {})
        edges, _ = decompose_to_cell_edges(elements)
        assert edges == {('H', 7, c) for c in range(2, 9)}

    def test_add_v_edge_appears_in_layout(self):
        elements = []
        apply_edge_corrections(elements, [], [
            {'type': 'V', 'col': 5, 'row_start': 1, 'row_end': 4}
        ], {})
        edges, _ = decompose_to_cell_edges(elements)
        assert edges == {('V', r, 5) for r in range(1, 4)}

    def test_preserves_text_elements(self):
        elements = [_full_rect(2, 4, 3, 6),
                    {'type': 'text', 'content': 'hi', 'row': 3, 'col': 4, 'end_col': 6}]
        apply_edge_corrections(elements, [], [], {})
        texts = [e for e in elements if e.get('type') == 'text']
        assert len(texts) == 1
        assert texts[0]['content'] == 'hi'
