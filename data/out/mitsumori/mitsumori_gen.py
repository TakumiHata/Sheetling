#!/usr/bin/env python3
"""方眼Excel生成スクリプト - mitsumori"""

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Border, Side, Font
from openpyxl.utils import get_column_letter


def place_cell(
    ws, r1, c1, r2, c2, value="", font=None, alignment=None, fill=None, border=None
):
    """セルに値・スタイルを設定し結合する。重複する既存結合は自動解除される。"""
    # 重複する既存の結合範囲を自動解除
    overlapping = [
        mr.coord
        for mr in ws.merged_cells.ranges
        if mr.min_row <= r2
        and mr.max_row >= r1
        and mr.min_col <= c2
        and mr.max_col >= c1
    ]
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
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)


def draw_line(
    ws,
    orientation,
    row=None,
    col=None,
    row_start=None,
    row_end=None,
    col_start=None,
    col_end=None,
    side=None,
):
    """1本の罫線を描画する。横線=top border, 縦線=left border。"""
    if side is None:
        side = Side(border_style="thin", color="000000")
    if orientation == "horizontal" and row is not None:
        for c in range(col_start, col_end + 1):
            cell = ws.cell(row=row, column=c)
            existing = cell.border
            cell.border = Border(
                left=existing.left,
                right=existing.right,
                top=side,
                bottom=existing.bottom,
            )
    elif orientation == "vertical" and col is not None:
        for r in range(row_start, row_end + 1):
            cell = ws.cell(row=r, column=col)
            existing = cell.border
            cell.border = Border(
                left=side,
                right=existing.right,
                top=existing.top,
                bottom=existing.bottom,
            )


def main():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # --- 1. グリッド設定（方眼紙） ---
    for col_idx in range(1, 120 + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 0.75
    for row_idx in range(1, 170 + 1):
        ws.row_dimensions[row_idx].height = 4.96

    # --- 2. rect要素（色彩情報除外のためスキップ） ---

    # --- 3. text要素（テキスト・フォント・配置） ---

    place_cell(
        ws,
        9,
        50,
        14,
        71,
        value="御見積書",
        font=Font(name="Meiryo", size=25.0, bold=True),
        alignment=Alignment(
            horizontal="center", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        19,
        80,
        21,
        106,
        value="見積書Ｎｏ．20121119_00001",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="left", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        21,
        6,
        23,
        22,
        value="△△株式会社",
        font=Font(name="Meiryo", size=12.5),
        alignment=Alignment(
            horizontal="left", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        22,
        23,
        23,
        61,
        value="御中",
        font=Font(name="Meiryo", size=12.5, bold=True),
        alignment=Alignment(
            horizontal="left", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        25,
        71,
        27,
        82,
        value="作成日:",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="left", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        25,
        84,
        27,
        89,
        value="2012年11月19日",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="left", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        29,
        12,
        42,
        14,
        value="下記のとおり御見積申し上げます。",
        font=Font(name="Meiryo", size=12.0, bold=True),
        alignment=Alignment(
            horizontal="left", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        33,
        15,
        35,
        61,
        value="何卒ご用命の程、お願い申し上げます。",
        font=Font(name="Meiryo", size=12.0),
        alignment=Alignment(
            horizontal="left", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        40,
        15,
        42,
        32,
        value="受渡期日:別途御打合せ",
        font=Font(name="Meiryo", size=12.0),
        alignment=Alignment(
            horizontal="left", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        29,
        71,
        42,
        82,
        value="ワークフロー商事株式会社",
        font=Font(name="Meiryo", size=11.0, bold=True),
        alignment=Alignment(
            horizontal="left", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        45,
        5,
        47,
        32,
        value="取引方法:別途御打合せ",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="left", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        44,
        71,
        47,
        82,
        value="東京都新宿区千代田９−９−９",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="left", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        49,
        101,
        52,
        109,
        value="WSビル",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="left", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        50,
        5,
        52,
        34,
        value="有効期限:発行日から30日",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="left", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        49,
        71,
        52,
        82,
        value="TEL",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="left", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        49,
        84,
        52,
        89,
        value="０３-××××−９９９９",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="left", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        54,
        71,
        56,
        82,
        value="FAX",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="left", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        54,
        84,
        56,
        89,
        value="０３-９９９９−××××",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="left", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        56,
        5,
        59,
        22,
        value="貴社管理番号：",
        font=Font(name="Meiryo", size=12.0),
        alignment=Alignment(
            horizontal="left", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        58,
        91,
        59,
        96,
        value="承認",
        font=Font(name="Meiryo", size=8.0),
        alignment=Alignment(
            horizontal="center", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        58,
        101,
        59,
        109,
        value="担当営業",
        font=Font(name="Meiryo", size=8.0),
        alignment=Alignment(
            horizontal="center", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        68,
        16,
        72,
        61,
        value="合計金額：",
        font=Font(name="Meiryo", size=14.0, bold=True),
        alignment=Alignment(
            horizontal="left", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        68,
        62,
        72,
        81,
        value="\\3,700,000(消費税別)",
        font=Font(name="Meiryo", size=18.0, bold=True),
        alignment=Alignment(
            horizontal="center", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        77,
        12,
        78,
        14,
        value="No.",
        font=Font(name="Meiryo", size=11.0, bold=True),
        alignment=Alignment(
            horizontal="center", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        77,
        16,
        78,
        61,
        value="摘",
        font=Font(name="Meiryo", size=11.0, bold=True),
        alignment=Alignment(
            horizontal="center", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        76,
        35,
        78,
        37,
        value="要",
        font=Font(name="Meiryo", size=11.0),
        alignment=Alignment(
            horizontal="center", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        77,
        63,
        78,
        69,
        value="数量",
        font=Font(name="Meiryo", size=11.0, bold=True),
        alignment=Alignment(
            horizontal="center", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        77,
        71,
        78,
        82,
        value="標準価格",
        font=Font(name="Meiryo", size=11.0, bold=True),
        alignment=Alignment(
            horizontal="center", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        77,
        84,
        78,
        89,
        value="見積価格",
        font=Font(name="Meiryo", size=11.0, bold=True),
        alignment=Alignment(
            horizontal="center", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        77,
        101,
        78,
        109,
        value="合計金額",
        font=Font(name="Meiryo", size=11.0, bold=True),
        alignment=Alignment(
            horizontal="center", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        80,
        12,
        81,
        14,
        value="1",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="right", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        80,
        16,
        81,
        61,
        value="ワークフローシステム",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="left", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        82,
        38,
        82,
        58,
        value="30ユーザーライセンス",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="left", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        80,
        63,
        81,
        69,
        value="1",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="right", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        80,
        71,
        81,
        82,
        value="2,700,000",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="right", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        80,
        84,
        81,
        89,
        value="2,700,000",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="right", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        80,
        101,
        81,
        109,
        value="2,700,000",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="right", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        83,
        12,
        85,
        14,
        value="2",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="right", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        83,
        16,
        85,
        61,
        value="初期設定費用",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="left", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        83,
        63,
        85,
        69,
        value="1",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="right", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        83,
        71,
        85,
        82,
        value="500,000",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="right", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        83,
        91,
        85,
        96,
        value="500,000",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="right", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        83,
        101,
        85,
        109,
        value="500,000",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="right", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        87,
        12,
        88,
        14,
        value="3",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="right", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        87,
        16,
        88,
        61,
        value="管理者費用",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="left", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        87,
        63,
        88,
        69,
        value="1",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="right", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        87,
        71,
        88,
        82,
        value="500,000",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="right", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        87,
        91,
        88,
        96,
        value="500,000",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="right", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        87,
        101,
        88,
        109,
        value="500,000",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="right", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        90,
        12,
        91,
        14,
        value="4",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="right", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        93,
        12,
        95,
        14,
        value="5",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="right", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        97,
        12,
        98,
        14,
        value="6",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="right", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        100,
        12,
        102,
        14,
        value="7",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="right", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        104,
        12,
        105,
        14,
        value="8",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="right", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        107,
        12,
        108,
        14,
        value="9",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="right", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        110,
        12,
        112,
        14,
        value="10",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="right", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        114,
        12,
        115,
        14,
        value="11",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="right", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        117,
        12,
        119,
        14,
        value="12",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="right", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        121,
        12,
        122,
        14,
        value="13",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="right", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        124,
        12,
        126,
        14,
        value="14",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="right", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        128,
        12,
        129,
        14,
        value="15",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="right", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        131,
        12,
        132,
        14,
        value="16",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="right", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        134,
        12,
        136,
        14,
        value="17",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="right", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        138,
        12,
        139,
        14,
        value="18",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="right", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        141,
        12,
        143,
        14,
        value="19",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="right", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        145,
        12,
        146,
        14,
        value="20",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="right", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        148,
        16,
        149,
        61,
        value="合",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="center", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        150,
        33,
        150,
        36,
        value="計",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="center", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        148,
        101,
        149,
        109,
        value="3,700,000",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="right", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        152,
        12,
        153,
        14,
        value="備",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="center", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        152,
        16,
        153,
        61,
        value="考",
        font=Font(name="Meiryo", size=10.0),
        alignment=Alignment(
            horizontal="center", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        155,
        16,
        163,
        61,
        value="・消費税は別途計上させていただきます。",
        font=Font(name="Meiryo", size=8.0),
        alignment=Alignment(
            horizontal="left", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    place_cell(
        ws,
        157,
        11,
        159,
        50,
        value="・製品の瑕疵、無償保証期間は御購入後3ヶ月間です。",
        font=Font(name="Meiryo", size=8.0),
        alignment=Alignment(
            horizontal="left", vertical="center", wrap_text=False, shrink_to_fit=True
        ),
    )

    # --- 4. 罫線（元PDF上のline座標を忠実に再現） ---

    side_thin = Side(border_style="thin", color="000000")

    draw_line(ws, "horizontal", row=79, col_start=11, col_end=113, side=side_thin)

    draw_line(ws, "horizontal", row=147, col_start=11, col_end=113, side=side_thin)

    draw_line(ws, "horizontal", row=43, col_start=5, col_end=62, side=side_thin)

    draw_line(ws, "horizontal", row=48, col_start=5, col_end=62, side=side_thin)

    draw_line(ws, "horizontal", row=53, col_start=5, col_end=62, side=side_thin)

    draw_line(ws, "horizontal", row=73, col_start=11, col_end=113, side=side_thin)

    draw_line(ws, "horizontal", row=59, col_start=5, col_end=48, side=side_thin)

    draw_line(ws, "horizontal", row=21, col_start=79, col_end=109, side=side_thin)

    draw_line(ws, "horizontal", row=154, col_start=11, col_end=113, side=side_thin)

    draw_line(ws, "horizontal", row=59, col_start=90, col_end=111, side=side_thin)

    draw_line(ws, "horizontal", row=28, col_start=79, col_end=108, side=side_thin)

    draw_line(ws, "horizontal", row=24, col_start=5, col_end=62, side=side_thin)

    draw_line(ws, "vertical", col=11, row_start=76, row_end=151, side=side_thin)

    draw_line(ws, "vertical", col=112, row_start=76, row_end=151, side=side_thin)

    draw_line(ws, "horizontal", row=76, col_start=11, col_end=113, side=side_thin)

    draw_line(ws, "horizontal", row=78, col_start=11, col_end=113, side=side_thin)

    draw_line(ws, "vertical", col=15, row_start=76, row_end=148, side=side_thin)

    draw_line(ws, "vertical", col=62, row_start=76, row_end=148, side=side_thin)

    draw_line(ws, "vertical", col=70, row_start=76, row_end=148, side=side_thin)

    draw_line(ws, "vertical", col=83, row_start=76, row_end=148, side=side_thin)

    draw_line(ws, "vertical", col=97, row_start=76, row_end=148, side=side_thin)

    draw_line(ws, "horizontal", row=150, col_start=11, col_end=113, side=side_thin)

    draw_line(ws, "horizontal", row=89, col_start=11, col_end=113, side=side_thin)

    draw_line(ws, "horizontal", row=92, col_start=11, col_end=113, side=side_thin)

    draw_line(ws, "horizontal", row=96, col_start=11, col_end=113, side=side_thin)

    draw_line(ws, "horizontal", row=99, col_start=11, col_end=113, side=side_thin)

    draw_line(ws, "horizontal", row=103, col_start=11, col_end=113, side=side_thin)

    draw_line(ws, "horizontal", row=106, col_start=11, col_end=113, side=side_thin)

    draw_line(ws, "horizontal", row=109, col_start=11, col_end=113, side=side_thin)

    draw_line(ws, "horizontal", row=113, col_start=10, col_end=113, side=side_thin)

    draw_line(ws, "horizontal", row=116, col_start=11, col_end=113, side=side_thin)

    draw_line(ws, "horizontal", row=120, col_start=10, col_end=113, side=side_thin)

    draw_line(ws, "horizontal", row=123, col_start=10, col_end=113, side=side_thin)

    draw_line(ws, "horizontal", row=127, col_start=11, col_end=113, side=side_thin)

    draw_line(ws, "horizontal", row=130, col_start=11, col_end=113, side=side_thin)

    draw_line(ws, "horizontal", row=133, col_start=10, col_end=113, side=side_thin)

    draw_line(ws, "horizontal", row=137, col_start=10, col_end=113, side=side_thin)

    draw_line(ws, "horizontal", row=140, col_start=11, col_end=113, side=side_thin)

    draw_line(ws, "horizontal", row=144, col_start=11, col_end=113, side=side_thin)

    draw_line(ws, "vertical", col=11, row_start=151, row_end=165, side=side_thin)

    draw_line(ws, "horizontal", row=151, col_start=11, col_end=113, side=side_thin)

    draw_line(ws, "vertical", col=112, row_start=151, row_end=165, side=side_thin)

    draw_line(ws, "horizontal", row=164, col_start=11, col_end=113, side=side_thin)

    draw_line(ws, "horizontal", row=67, col_start=90, col_end=111, side=side_thin)

    draw_line(ws, "horizontal", row=57, col_start=90, col_end=111, side=side_thin)

    draw_line(ws, "vertical", col=90, row_start=57, row_end=68, side=side_thin)

    draw_line(ws, "vertical", col=110, row_start=57, row_end=68, side=side_thin)

    draw_line(ws, "vertical", col=100, row_start=57, row_end=68, side=side_thin)

    draw_line(ws, "horizontal", row=82, col_start=11, col_end=113, side=side_thin)

    draw_line(ws, "horizontal", row=86, col_start=11, col_end=113, side=side_thin)

    # --- 5. 印刷設定（必須） ---
    from openpyxl.worksheet.page import PageMargins

    # デフォルトのビュー設定（倍率）
    ws.sheet_view.zoomScale = (
        200  # 実寸A4サイズセルが小さいため、作業用の表示倍率を拡大
    )
    ws.page_margins = PageMargins(left=0, right=0, top=0, bottom=0, header=0, footer=0)
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.paperSize = 9  # A4
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    # 印刷範囲（Print Area）を明示的に指定することで、右側が余る現象を防止する
    ws.print_area = f"A1:{get_column_letter(113)}165"
    ws.print_options.horizontalCentered = True

    # 改ページ指定（最後のページは不要）
    from openpyxl.worksheet.pagebreak import Break

    ws.cell(row=1, column=1, value=" ").font = Font(color="FFFFFF")
    ws.cell(row=165, column=113, value=" ").font = Font(color="FFFFFF")

    wb.save("mitsumori.xlsx")
    print("Saved: mitsumori.xlsx")


if __name__ == "__main__":
    main()
