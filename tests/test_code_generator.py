"""
CodeGenerator のユニットテスト
"""
import pytest
from src.core.code_generator import CodeGenerator
from src.core.placement_generator import (
    PlacementCommand,
    PlacementResult,
    TableStructure,
    LineElement,
)


class TestCodeGeneration:
    """コード生成の基本テスト"""

    def _make_simple_result(self) -> PlacementResult:
        """テスト用の簡単な PlacementResult を作成する"""
        return PlacementResult(
            commands=[
                PlacementCommand(
                    category="rect", r1=4, c1=22, r2=7, c2=38,
                    fill_color="#F2F2F2",
                ),
                PlacementCommand(
                    category="text_outside", r1=4, c1=25, r2=7, c2=35,
                    value="御見積書", font_size=25.0, font_bold=True,
                    alignment="center",
                ),
                PlacementCommand(
                    category="text_table", r1=38, c1=6, r2=38, c2=6,
                    value="No.", font_size=11.0, font_bold=True,
                    alignment="center",
                ),
            ],
            table_structures=[
                TableStructure(
                    v_cols=[5, 7, 31],
                    h_rows=[37, 39, 41],
                    table_row_min=37,
                    table_row_max=41,
                    table_col_min=5,
                    table_col_max=31,
                )
            ],
            line_elements=[
                LineElement(row_start=37, row_end=37, col_start=5, col_end=31, orientation="horizontal"),
                LineElement(row_start=39, row_end=39, col_start=5, col_end=31, orientation="horizontal"),
                LineElement(row_start=41, row_end=41, col_start=5, col_end=31, orientation="horizontal"),
                LineElement(row_start=37, row_end=41, col_start=5, col_end=5, orientation="vertical"),
                LineElement(row_start=37, row_end=41, col_start=7, col_end=7, orientation="vertical"),
                LineElement(row_start=37, row_end=41, col_start=31, col_end=31, orientation="vertical"),
            ],
        )

    def test_generates_valid_python(self):
        """生成コードが有効なPythonであること"""
        gen = CodeGenerator()
        result = self._make_simple_result()
        code = gen.generate(
            placement_result=result,
            grid_cols=60, grid_rows=85,
            col_width=2.4, row_height=18.0,
            page_count=1,
            output_filename="test.xlsx",
            pdf_name="test",
        )
        # SyntaxErrorが出ないこと
        compile(code, "test_gen.py", "exec")

    def test_contains_place_cell_calls(self):
        """生成コードに place_cell 呼び出しが含まれること"""
        gen = CodeGenerator()
        result = self._make_simple_result()
        code = gen.generate(
            placement_result=result,
            grid_cols=60, grid_rows=85,
            col_width=2.4, row_height=18.0,
            page_count=1,
            output_filename="test.xlsx",
            pdf_name="test",
        )
        assert "place_cell(ws, 4, 22, 7, 38" in code  # rect
        assert "place_cell(ws, 4, 25, 7, 35" in code  # text_outside
        assert "place_cell(ws, 38, 6, 38, 6" in code  # text_table
        assert '御見積書' in code
        assert 'No.' in code

    def test_contains_draw_line(self):
        """生成コードに draw_line 呼び出しが含まれること"""
        gen = CodeGenerator()
        result = self._make_simple_result()
        code = gen.generate(
            placement_result=result,
            grid_cols=60, grid_rows=85,
            col_width=2.4, row_height=18.0,
            page_count=1,
            output_filename="test.xlsx",
            pdf_name="test",
        )
        assert 'draw_line(ws, "horizontal"' in code
        assert 'draw_line(ws, "vertical"' in code
        # draw_table_bordersは廃止
        assert 'draw_table_borders' not in code

    def test_contains_grid_setup(self):
        """生成コードにグリッド設定が含まれること"""
        gen = CodeGenerator()
        result = self._make_simple_result()
        code = gen.generate(
            placement_result=result,
            grid_cols=60, grid_rows=85,
            col_width=2.4, row_height=18.0,
            page_count=1,
            output_filename="test.xlsx",
            pdf_name="test",
        )
        assert "width = 2.4" in code
        assert "height = 18.0" in code
        assert "range(1, 60 + 1)" in code
        assert "range(1, 85 + 1)" in code

    def test_contains_print_setup(self):
        """生成コードに印刷設定が含まれること"""
        gen = CodeGenerator()
        result = self._make_simple_result()
        code = gen.generate(
            placement_result=result,
            grid_cols=60, grid_rows=85,
            col_width=2.4, row_height=18.0,
            page_count=1,
            output_filename="test.xlsx",
            pdf_name="test",
        )
        assert "fitToPage = True" in code
        assert "paperSize = 9" in code
        assert "fitToHeight = 1" in code
        assert 'wb.save("test.xlsx")' in code

    def test_escapes_special_characters(self):
        """特殊文字がエスケープされること"""
        gen = CodeGenerator()
        result = PlacementResult(
            commands=[
                PlacementCommand(
                    category="text_outside", r1=10, c1=5, r2=10, c2=20,
                    value='\\3,700,000', font_size=18.0, font_bold=True,
                    alignment="right",
                ),
            ],
        )
        code = gen.generate(
            placement_result=result,
            grid_cols=60, grid_rows=85,
            col_width=2.4, row_height=18.0,
            page_count=1,
            output_filename="test.xlsx",
            pdf_name="test",
        )
        # コード内でエスケープされていること
        compile(code, "test_gen.py", "exec")
        # 値が含まれていること
        assert "3,700,000" in code

    def test_correct_place_cell_count(self):
        """place_cell呼び出し数が正しいこと"""
        gen = CodeGenerator()
        result = self._make_simple_result()
        code = gen.generate(
            placement_result=result,
            grid_cols=60, grid_rows=85,
            col_width=2.4, row_height=18.0,
            page_count=1,
            output_filename="test.xlsx",
            pdf_name="test",
        )
        # main()内のplace_cell呼び出し数 = 3 (rect + text_outside + text_table)
        # 関数定義内のplace_cellは除外
        main_section = code.split("def main():")[1]
        call_count = main_section.count("place_cell(ws,")
        assert call_count == 3


class TestIntegrationWithRealData:
    """実データを使った統合テスト"""

    def test_generate_from_real_json(self):
        """実際のJSONデータからコード生成が正常に動作すること"""
        import json
        from pathlib import Path
        from src.core.prompt_builder import PromptBuilder
        from src.core.placement_generator import PlacementGenerator

        json_path = Path("data/03_layout_json/mitsumori.json")
        if not json_path.exists():
            pytest.skip("mitsumori.json が存在しません")

        with open(json_path, "r", encoding="utf-8") as f:
            json_data = json.load(f)

        builder = PromptBuilder("data/04_prompt")
        compressed = builder._compress_json(json_data)

        # PlacementGenerator で配置命令を生成
        placement_gen = PlacementGenerator()
        result = placement_gen.generate(compressed)

        # CodeGenerator でコードを生成
        code_gen = CodeGenerator()
        code = code_gen.generate(
            placement_result=result,
            grid_cols=60, grid_rows=85,
            col_width=2.4, row_height=18.0,
            page_count=1,
            output_filename="mitsumori.xlsx",
            pdf_name="mitsumori",
        )

        # 基本検証
        compile(code, "mitsumori_gen.py", "exec")
        assert "place_cell(ws," in code
        assert 'draw_line(ws, "horizontal"' in code or 'draw_line(ws, "vertical"' in code
        assert 'wb.save("mitsumori.xlsx")' in code

        # テキスト内容が含まれていること
        assert "御見積書" in code or "見積" in code
        assert "No." in code

        # 重複解消不能が大幅に減っていること（修正前は12件以上、テーブル外テキストの元データ重複は残る）
        unresolvable = [w for w in result.warnings if "重複解消不能" in w]
        assert len(unresolvable) < 10, f"重複解消不能が多すぎます: {len(unresolvable)}件"
