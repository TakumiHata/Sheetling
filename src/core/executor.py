"""
AI出力のPythonソースコードを実行し、3シートのExcelファイルを生成するモジュール。

- 1シート目: AI生成コードによるExcel描画（PDFの再現）
- 2シート目: カラーコード・フォント情報の一覧
- 3シート目: PDF画像の添付
"""

import os
import traceback

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from src.core.config import config
from src.utils.logger import get_logger

logger = get_logger(__name__)


class Executor:
    """AI出力のPythonコードを実行し、2シート構成のExcelを生成する"""

    def __init__(self):
        self.col_width = config.excel.col_width_chars
        self.row_height = config.excel.row_height_pt
        self.max_cols = config.grid.target_cols
        self.max_rows = config.grid.target_rows

    def execute(
        self,
        gen_py_path: str,
        output_xlsx_path: str,
        fonts: list[dict],
        colors: list,
    ) -> str:
        """
        AI生成のPythonソースを実行し、2シートExcelを出力する。

        Args:
            gen_py_path: AI出力の .py ファイルパス
            output_xlsx_path: 出力する .xlsx ファイルパス
            fonts: 抽出されたフォント情報のリスト
            colors: 抽出されたカラー情報のリスト

        Returns:
            出力されたExcelファイルパス
        """
        logger.info(f"--- Executing AI-generated code: {gen_py_path} ---")

        # Workbook初期化（方眼設定済み1シート目付き）
        wb = self._create_workbook()
        ws = wb.active

        # --- 1シート目: AI生成コードの実行 ---
        self._execute_generated_code(gen_py_path, wb, ws)

        # --- 2シート目: フォント・カラー情報の一覧 ---
        self._create_info_sheet(wb, fonts, colors)

        # 保存
        os.makedirs(os.path.dirname(os.path.abspath(output_xlsx_path)), exist_ok=True)
        wb.save(output_xlsx_path)
        logger.info(f"✅ 2-sheet Excel saved: {output_xlsx_path}")

        return output_xlsx_path

    def _create_workbook(self) -> openpyxl.Workbook:
        """A4方眼設定済みのWorkbookを作成する"""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "変換結果"

        # ページ設定（A4）
        ws.page_setup.paperSize = ws.PAPERSIZE_A4

        # 印刷余白
        ws.page_margins.left = 0.25
        ws.page_margins.right = 0.25
        ws.page_margins.top = 0.75
        ws.page_margins.bottom = 0.75
        ws.page_margins.header = 0.3
        ws.page_margins.footer = 0.3

        # 印刷時にA4幅いっぱいに拡大（方眼サイズはそのまま、印刷スケーリングで対応）
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True

        # 方眼の列幅・行高さを設定
        for col_idx in range(1, self.max_cols + 1):
            ws.column_dimensions[get_column_letter(col_idx)].width = self.col_width

        for row_idx in range(1, self.max_rows + 1):
            ws.row_dimensions[row_idx].height = self.row_height

        return wb

    def _execute_generated_code(self, gen_py_path: str, wb, ws):
        """AI生成のPythonコードを読み込んで実行する"""
        logger.info(f"Loading generated code: {gen_py_path}")

        # AIが生成したコードを文字列として読み込む
        with open(gen_py_path, "r", encoding="utf-8") as f:
            code = f.read()

        # 動的実行のための独立した名前空間を用意する
        exec_globals = {
            "__builtins__": __builtins__,
        }

        try:
            # コード文字列をコンパイルし、用意した名前空間上で実行する
            exec(compile(code, gen_py_path, "exec"), exec_globals)

            if "generate" not in exec_globals:
                raise RuntimeError(
                    f"AI生成コードに `generate(wb, ws)` 関数が定義されていません: {gen_py_path}"
                )

            # generate関数を呼び出し
            exec_globals["generate"](wb, ws)
            logger.info("✅ AI generated code executed successfully (Sheet 1)")

        except Exception as e:
            logger.error(f"❌ AI generated code execution failed: {e}")
            logger.error(traceback.format_exc())

            # エラー発生時に内容がわかるよう、最初の数行分だけセルの高さを広げる
            ws.row_dimensions[1].height = 30
            ws.row_dimensions[2].height = 20
            ws.row_dimensions[3].height = 60

            ws["A1"] = "⚠ AI生成コードの実行に失敗しました"
            ws["A1"].font = Font(color="FF0000", size=14, bold=True)

            ws["A2"] = f"エラー種別: {type(e).__name__}"
            ws["A2"].font = Font(size=10)

            ws["A3"] = f"詳細: {str(e)}"
            ws["A3"].font = Font(size=10)
            ws["A3"].alignment = Alignment(wrap_text=True)

            # エラー内容が見えるようにA列のみ幅を拡大
            ws.column_dimensions["A"].width = 80

    def _create_info_sheet(self, wb, fonts: list[dict], colors: list):
        """2シート目: フォント・カラー情報の一覧を作成する"""
        ws = wb.create_sheet(title="フォント・カラー情報")

        # ヘッダースタイル
        header_font = Font(bold=True, size=11)
        header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        thin_border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin"),
        )

        # --- フォント一覧 ---
        ws["A1"] = "■ フォント一覧"
        ws["A1"].font = Font(bold=True, size=14)

        headers = ["No.", "フォント名", "サイズ (pt)"]
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col_idx, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.border = thin_border

        for i, font_info in enumerate(fonts):
            row = 4 + i
            ws.cell(row=row, column=1, value=i + 1).border = thin_border
            ws.cell(row=row, column=2, value=font_info.get("fontname", "")).border = thin_border
            ws.cell(row=row, column=3, value=font_info.get("size", 0)).border = thin_border

        # --- カラー一覧 ---
        color_start_row = 4 + len(fonts) + 2
        ws.cell(row=color_start_row, column=1, value="■ カラーコード一覧")
        ws.cell(row=color_start_row, column=1).font = Font(bold=True, size=14)

        color_headers = ["No.", "カラーコード", "プレビュー"]
        for col_idx, header in enumerate(color_headers, 1):
            cell = ws.cell(row=color_start_row + 2, column=col_idx, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.border = thin_border

        for i, color in enumerate(colors):
            row = color_start_row + 3 + i
            color_hex = color if isinstance(color, str) else str(color)
            ws.cell(row=row, column=1, value=i + 1).border = thin_border
            ws.cell(row=row, column=2, value=color_hex).border = thin_border
            # プレビュー（背景色）
            preview_cell = ws.cell(row=row, column=3, value="")
            preview_cell.border = thin_border
            try:
                fill_color = color_hex.lstrip("#")
                if len(fill_color) == 6:
                    preview_cell.fill = PatternFill(
                        start_color=fill_color, end_color=fill_color, fill_type="solid"
                    )
            except Exception:
                pass

        # 列幅調整
        ws.column_dimensions["A"].width = 8
        ws.column_dimensions["B"].width = 30
        ws.column_dimensions["C"].width = 15
