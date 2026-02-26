"""
Pythonコード直接生成モジュール

PlacementResult から実行可能な openpyxl Pythonスクリプトを直接生成する。
LLM の責務を「コードの展開」から「コードのレビュー・微調整」に変更するための中核モジュール。
"""
from src.core.placement_generator import PlacementResult
from src.utils.logger import get_logger

logger = get_logger(__name__)


class CodeGenerator:
    """
    配置命令リストから完全な Python スクリプトを生成する。

    生成されるスクリプトは以下を含む:
    - place_cell / draw_table_borders 関数定義
    - グリッド設定（列幅・行高の均一設定）
    - rect 要素の背景色設定
    - text 要素の配置（テーブル内・外）
    - テーブル罫線の描画
    - 印刷設定（A4, fitToPage）
    """

    def generate(
        self,
        placement_result: PlacementResult,
        grid_cols: int,
        grid_rows: int,
        col_width: float,
        row_height: float,
        page_count: int,
        output_filename: str,
        pdf_name: str,
    ) -> str:
        """
        配置命令リストから完全な Python スクリプトを生成する。

        Args:
            placement_result: PlacementGenerator の出力
            grid_cols: グリッド列数
            grid_rows: グリッド行数
            col_width: openpyxl 列幅単位
            row_height: 行高（pt）
            page_count: ページ数（印刷設定用）
            output_filename: 出力ファイル名
            pdf_name: 元 PDF 名

        Returns:
            実行可能な Python スクリプト文字列
        """
        lines = []

        # --- ヘッダー ---
        lines.append('#!/usr/bin/env python3')
        lines.append(f'"""方眼Excel生成スクリプト - {pdf_name}"""')
        lines.append('from openpyxl import Workbook')
        lines.append('from openpyxl.styles import PatternFill, Alignment, Border, Side, Font')
        lines.append('from openpyxl.utils import get_column_letter')
        lines.append('')
        lines.append('')

        # --- place_cell 関数 ---
        lines.append(self._place_cell_function())
        lines.append('')
        lines.append('')

        # --- draw_table_borders 関数 ---
        lines.append(self._draw_table_borders_function())
        lines.append('')
        lines.append('')

        # --- main 関数 ---
        lines.append('def main():')
        lines.append('    wb = Workbook()')
        lines.append('    ws = wb.active')
        lines.append('    ws.title = "Sheet1"')
        lines.append('')

        # --- 1. グリッド設定 ---
        lines.append('    # --- 1. グリッド設定（方眼紙） ---')
        lines.append(f'    for col_idx in range(1, {grid_cols} + 1):')
        lines.append(f'        ws.column_dimensions[get_column_letter(col_idx)].width = {col_width}')
        lines.append(f'    for row_idx in range(1, {grid_rows} + 1):')
        lines.append(f'        ws.row_dimensions[row_idx].height = {row_height}')
        lines.append('')

        # --- 2. rect 要素 ---
        rect_cmds = [c for c in placement_result.commands if c.category == "rect"]
        if rect_cmds:
            lines.append('    # --- 2. rect要素（背景色） ---')
            for cmd in rect_cmds:
                fill = cmd.fill_color.replace("#", "") if cmd.fill_color else "FFFFFF"
                lines.append(
                    f'    place_cell(ws, {cmd.r1}, {cmd.c1}, {cmd.r2}, {cmd.c2}, '
                    f'fill=PatternFill(start_color="{fill}", end_color="{fill}", fill_type="solid"))'
                )
            lines.append('')

        # --- 3. text 要素 ---
        text_cmds = [c for c in placement_result.commands if c.category in ("text_outside", "text_table")]
        if text_cmds:
            lines.append('    # --- 3. text要素（テキスト・フォント・配置） ---')
            for cmd in text_cmds:
                bold_str = ", bold=True" if cmd.font_bold else ""
                align = cmd.alignment or "left"
                # 値のエスケープ
                escaped_value = cmd.value.replace('\\', '\\\\').replace('"', '\\"')
                lines.append(
                    f'    place_cell(ws, {cmd.r1}, {cmd.c1}, {cmd.r2}, {cmd.c2}, '
                    f'value="{escaped_value}", '
                    f'font=Font(name="Meiryo", size={cmd.font_size}{bold_str}), '
                    f'alignment=Alignment(horizontal="{align}", vertical="center", wrap_text=True))'
                )
            lines.append('')

        # --- 4. テーブル罫線 ---
        for ts in placement_result.table_structures:
            if ts.v_cols and ts.h_rows:
                lines.append('    # --- 4. テーブル罫線（place_cellを使わないこと） ---')
                lines.append(
                    f'    draw_table_borders(ws, '
                    f'v_cols={ts.v_cols}, '
                    f'h_rows={ts.h_rows}, '
                    f'row_min={ts.table_row_min}, '
                    f'row_max={ts.table_row_max}, '
                    f'col_min={ts.table_col_min}, '
                    f'col_max={ts.table_col_max})'
                )
                lines.append('')

        # --- 5. 印刷設定 ---
        lines.append('    # --- 5. 印刷設定（必須） ---')
        lines.append('    ws.sheet_properties.pageSetUpPr.fitToPage = True')
        lines.append('    ws.page_setup.paperSize = 9  # A4')
        lines.append('    ws.page_setup.orientation = "portrait"')
        lines.append('    ws.page_setup.fitToWidth = 1')
        lines.append(f'    ws.page_setup.fitToHeight = {page_count}')
        lines.append('')
        lines.append(f'    wb.save("{output_filename}")')
        lines.append(f'    print(f"Saved: {output_filename}")')
        lines.append('')
        lines.append('')
        lines.append('if __name__ == "__main__":')
        lines.append('    main()')
        lines.append('')

        code = '\n'.join(lines)

        # 生成コードの文法チェック
        try:
            compile(code, f"{pdf_name}_gen.py", "exec")
            logger.info(f"コード生成完了: {len(text_cmds)}件のtext配置, {len(rect_cmds)}件のrect配置")
        except SyntaxError as e:
            logger.error(f"生成コードにSyntaxError: {e}")

        return code

    @staticmethod
    def _place_cell_function() -> str:
        """place_cell 関数のソースコードを返す。"""
        return '''def place_cell(ws, r1, c1, r2, c2, value="", font=None, alignment=None, fill=None, border=None):
    """セルに値・スタイルを設定し結合する。重複する既存結合は自動解除される。"""
    # 重複する既存の結合範囲を自動解除
    overlapping = [mr.coord for mr in ws.merged_cells.ranges
                   if mr.min_row <= r2 and mr.max_row >= r1
                   and mr.min_col <= c2 and mr.max_col >= c1]
    for coord in overlapping:
        ws.unmerge_cells(coord)
    # 左上セルに値・スタイルを設定
    cell = ws.cell(row=r1, column=c1, value=value)
    if font:
        cell.font = font
    if alignment:
        cell.alignment = alignment
    if fill:
        cell.fill = fill
    # 罫線は全セルに適用（結合後はMergedCellとなり設定不可のため）
    if border:
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                ws.cell(row=r, column=c).border = border
    # セル結合
    if r2 > r1 or c2 > c1:
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)'''

    @staticmethod
    def _draw_table_borders_function() -> str:
        """draw_table_borders 関数のソースコードを返す。"""
        return '''def draw_table_borders(ws, v_cols, h_rows, row_min, row_max, col_min, col_max):
    """テーブルの罫線を描画する。place_cellとの競合を避けるため直接cell.borderを設定する。"""
    side_thin = Side(border_style="thin", color="000000")
    # 縦線
    for c in v_cols:
        for r in range(row_min, row_max + 1):
            cell = ws.cell(row=r, column=c)
            existing = cell.border
            cell.border = Border(
                left=side_thin,
                right=existing.right,
                top=existing.top,
                bottom=existing.bottom
            )
    # 横線
    for r in h_rows:
        for c in range(col_min, col_max + 1):
            cell = ws.cell(row=r, column=c)
            existing = cell.border
            cell.border = Border(
                left=existing.left,
                right=existing.right,
                top=side_thin,
                bottom=existing.bottom
            )'''
