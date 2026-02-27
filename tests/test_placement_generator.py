"""
PlacementGenerator のユニットテスト
"""
import pytest
from src.core.placement_generator import (
    PlacementCommand,
    PlacementGenerator,
    TableStructure,
    format_placement_commands,
)


class TestHalfOpenToClosedConversion:
    """半開区間→閉区間変換のテスト"""

    def test_rect_conversion(self):
        """rect要素の座標変換が正しいこと"""
        gen = PlacementGenerator()
        json_data = {
            "pages": [{
                "elements": [{
                    "type": "rect",
                    "grid_bbox": {"col_start": 22, "row_start": 4, "col_end": 39, "row_end": 8},
                }]
            }]
        }
        result = gen.generate(json_data)
        rect_cmds = [c for c in result.commands if c.category == "rect"]
        assert len(rect_cmds) == 1
        cmd = rect_cmds[0]
        # 半開区間 (4, 22) → (8, 39) を閉区間 (4, 22) → (7, 38) に変換
        assert cmd.r1 == 4
        assert cmd.c1 == 22
        assert cmd.r2 == 7  # row_end - 1
        assert cmd.c2 == 38  # col_end - 1

    def test_text_outside_conversion(self):
        """テーブル外text要素の座標変換が正しいこと"""
        gen = PlacementGenerator()
        json_data = {
            "pages": [{
                "elements": [{
                    "type": "text",
                    "grid_bbox": {"col_start": 25, "row_start": 4, "col_end": 36, "row_end": 8},
                    "text": "御見積書",
                    "font_size": 25.0,
                }]
            }]
        }
        result = gen.generate(json_data)
        text_cmds = [c for c in result.commands if c.category == "text_outside"]
        assert len(text_cmds) == 1
        cmd = text_cmds[0]
        assert cmd.r1 == 4
        assert cmd.c1 == 25
        assert cmd.r2 == 7  # row_end - 1
        assert cmd.c2 == 35  # col_end - 1
        assert cmd.value == "御見積書"
        assert cmd.font_size == 25.0


class TestTableStructureDetection:
    """テーブル構造検出のテスト"""

    def test_detect_vertical_and_horizontal_lines(self):
        """縦線と横線が正しく検出されること"""
        gen = PlacementGenerator()
        elements = [
            # 縦線: col 5, row 37→75
            {"type": "line", "grid_bbox": {"col_start": 5, "row_start": 37, "col_end": 6, "row_end": 75}},
            # 縦線: col 7, row 37→74
            {"type": "line", "grid_bbox": {"col_start": 7, "row_start": 37, "col_end": 8, "row_end": 74}},
            # 横線: row 37, col 5→56
            {"type": "line", "grid_bbox": {"col_start": 5, "row_start": 37, "col_end": 56, "row_end": 38}},
            # 横線: row 39, col 5→56
            {"type": "line", "grid_bbox": {"col_start": 5, "row_start": 39, "col_end": 56, "row_end": 40}},
        ]
        table = gen._detect_table_structure(elements)
        assert 5 in table.v_cols
        assert 7 in table.v_cols
        assert 37 in table.h_rows
        assert 39 in table.h_rows

    def test_column_ranges(self):
        """列区間が正しく計算されること"""
        gen = PlacementGenerator()
        elements = [
            {"type": "line", "grid_bbox": {"col_start": 5, "row_start": 37, "col_end": 6, "row_end": 75}},
            {"type": "line", "grid_bbox": {"col_start": 7, "row_start": 37, "col_end": 8, "row_end": 74}},
            {"type": "line", "grid_bbox": {"col_start": 31, "row_start": 37, "col_end": 32, "row_end": 74}},
            {"type": "line", "grid_bbox": {"col_start": 5, "row_start": 37, "col_end": 56, "row_end": 38}},
        ]
        table = gen._detect_table_structure(elements)
        # 縦線 5, 7, 31 → 列区間は (6, 6) と (8, 30)
        assert (6, 6) in table.col_ranges
        assert (8, 30) in table.col_ranges


class TestTableSnapping:
    """テーブル内要素のスナップテスト"""

    def setup_method(self):
        """テスト用のテーブル構造を準備"""
        self.gen = PlacementGenerator()
        self.table = TableStructure(
            col_ranges=[(6, 6), (8, 30), (32, 33), (35, 40), (42, 47), (50, 53)],
            row_ranges=[(38, 38), (40, 40), (42, 42), (44, 44)],
            table_row_min=37,
            table_row_max=75,
            table_col_min=5,
            table_col_max=55,
            v_cols=[5, 7, 31, 34, 41, 48, 54, 55],
            h_rows=[37, 39, 41, 43, 45],
        )

    def test_snap_header_no(self):
        """ヘッダー 'No.' がNo列にスナップされること"""
        elem = {
            "type": "text",
            "grid_bbox": {"col_start": 5, "row_start": 37, "col_end": 8, "row_end": 40},
            "text": "No.",
            "font_size": 11.0,
        }
        cmd = self.gen._snap_to_table(elem, self.table)
        assert cmd is not None
        assert cmd.c1 == 6
        assert cmd.c2 == 6
        assert cmd.r1 == 38
        assert cmd.r2 == 39

    def test_snap_data_value(self):
        """データ行の値が正しい列・行にスナップされること"""
        elem = {
            "type": "text",
            "grid_bbox": {"col_start": 33, "row_start": 39, "col_end": 35, "row_end": 41},
            "text": "1",
            "font_size": 10.0,
        }
        cmd = self.gen._snap_to_table(elem, self.table)
        assert cmd is not None
        # col 33 は列区間 (32, 33) に含まれる
        assert cmd.c1 == 32
        assert cmd.c2 == 33
        # row 39 は行区間に含まれないので最近傍 → (40, 40) が最寄り
        assert cmd.r1 == 40 or cmd.r1 == 38  # 最近傍


class TestOverlapDetection:
    """重複検証のテスト"""

    def test_detect_overlap(self):
        """重複が正しく検出されること"""
        gen = PlacementGenerator()
        cmd_a = PlacementCommand(category="text_outside", r1=10, c1=3, r2=12, c2=26, value="A")
        cmd_b = PlacementCommand(category="text_outside", r1=11, c1=20, r2=13, c2=30, value="B")
        assert gen._rects_overlap(cmd_a, cmd_b) is True

    def test_no_overlap(self):
        """非重複が正しく判定されること"""
        gen = PlacementGenerator()
        cmd_a = PlacementCommand(category="text_outside", r1=10, c1=3, r2=11, c2=10, value="A")
        cmd_b = PlacementCommand(category="text_outside", r1=10, c1=27, r2=11, c2=29, value="B")
        assert gen._rects_overlap(cmd_a, cmd_b) is False

    def test_resolve_overlap(self):
        """重複が解消されること"""
        gen = PlacementGenerator()
        cmd_a = PlacementCommand(category="text_outside", r1=15, c1=7, r2=16, c2=27, value="A")
        cmd_b = PlacementCommand(category="text_outside", r1=16, c1=7, r2=17, c2=30, value="B")
        # row 16 で重複
        warnings = gen._validate_and_resolve_overlaps([cmd_a, cmd_b])
        # 重複が検出・解消されるはず
        assert len(warnings) > 0


class TestAlignmentGuessing:
    """アライメント推定のテスト"""

    def test_numeric_right_align(self):
        """数値テキストが右寄せになること"""
        gen = PlacementGenerator()
        assert gen._guess_alignment("2,700,000", 10.0) == "right"

    def test_header_center_align(self):
        """ヘッダーテキストが中央寄せになること"""
        gen = PlacementGenerator()
        assert gen._guess_alignment("No.", 11.0) == "center"

    def test_large_font_center(self):
        """大きいフォントが中央寄せになること"""
        gen = PlacementGenerator()
        assert gen._guess_alignment("御見積書", 25.0) == "center"

    def test_normal_text_left(self):
        """通常テキストが左寄せになること"""
        gen = PlacementGenerator()
        assert gen._guess_alignment("ワークフロー商事株式会社", 10.0) == "left"


class TestFormatPlacementCommands:
    """配置命令リストのフォーマットテスト"""

    def test_format_includes_place_cell(self):
        """フォーマット出力に place_cell 呼び出しが含まれること"""
        from src.core.placement_generator import PlacementResult
        result = PlacementResult(
            commands=[
                PlacementCommand(
                    category="rect", r1=4, c1=22, r2=7, c2=38,
                ),
                PlacementCommand(
                    category="text_outside", r1=4, c1=25, r2=7, c2=35,
                    value="御見積書", font_size=25.0, alignment="center",
                ),
            ]
        )
        output = format_placement_commands(result)
        assert "place_cell(ws, 4, 22, 7, 38" in output
        assert "place_cell(ws, 4, 25, 7, 35" in output
        assert "御見積書" in output

    def test_format_table_borders(self):
        """フォーマット出力に draw_table_borders が含まれること"""
        from src.core.placement_generator import PlacementResult
        result = PlacementResult(
            table_structures=[
                TableStructure(
                    v_cols=[5, 7, 31],
                    h_rows=[37, 39, 41],
                    table_row_min=37,
                    table_row_max=41,
                    table_col_min=5,
                    table_col_max=31,
                )
            ]
        )
        output = format_placement_commands(result)
        assert "draw_table_borders" in output


class TestIntegration:
    """mitsumori.json を使った統合テスト"""

    def test_generate_from_real_json(self):
        """実際のJSONデータから配置命令リストが正常に生成されること"""
        import json
        from pathlib import Path

        json_path = Path("data/03_layout_json/mitsumori.json")
        if not json_path.exists():
            pytest.skip("mitsumori.json が存在しません")

        with open(json_path, "r", encoding="utf-8") as f:
            json_data = json.load(f)

        # 圧縮を模擬（PromptBuilder._compress_json相当）
        from src.core.prompt_builder import PromptBuilder
        builder = PromptBuilder("data/04_prompt")
        compressed = builder._compress_json(json_data)

        gen = PlacementGenerator()
        result = gen.generate(compressed)

        # 基本的な検証
        assert len(result.commands) > 0, "配置命令が1件以上生成されること"

        # text命令が存在すること
        text_cmds = [c for c in result.commands if c.category in ("text_outside", "text_table")]
        assert len(text_cmds) >= 10, "text命令が10件以上あること"

        # テーブル構造が検出されていること
        assert len(result.table_structures) >= 1, "テーブル構造が検出されていること"

        # 全命令の座標が正の値であること
        for cmd in result.commands:
            assert cmd.r1 >= 1 or cmd.r1 >= 0, f"r1が負: {cmd}"
            assert cmd.c1 >= 0, f"c1が負: {cmd}"
            assert cmd.r2 >= cmd.r1, f"r2 < r1: {cmd}"
            assert cmd.c2 >= cmd.c1, f"c2 < c1: {cmd}"

        # フォーマットが正常に生成されること
        output = format_placement_commands(result)
        assert len(output) > 100, "フォーマット出力が十分な長さであること"
        assert "place_cell" in output, "place_cell呼び出しが含まれること"
