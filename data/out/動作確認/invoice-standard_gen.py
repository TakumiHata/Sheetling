from openpyxl import Workbook
from openpyxl.styles import Border, Side, Alignment, PatternFill, Font
from openpyxl.utils import get_column_letter
import json

thin = Side(style='thin')

def apply_outer_border(ws, s_row, e_row, s_col, e_col, fill_color=None):
    fill = PatternFill(patternType='solid', fgColor=fill_color) if fill_color else None
    for r in range(s_row, e_row + 1):
        for c in range(s_col, e_col + 1):
            target = ws.cell(row=r, column=c)
            try:
                target.alignment = Alignment(horizontal='left', vertical='center', wrap_text=False)
            except AttributeError:
                pass
            top    = thin if r == s_row else None
            bottom = thin if r == e_row else None
            left   = thin if c == s_col else None
            right  = thin if c == e_col else None
            try:
                target.border = Border(top=top, bottom=bottom, left=left, right=right)
            except AttributeError:
                pass
            if fill:
                try:
                    target.fill = fill
                except AttributeError:
                    pass

def create_excel_from_json(json_data):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # グリッド設定（全列・全行）
    for c in range(1, 37):
        ws.column_dimensions[get_column_letter(c)].width = 2.53
    for r in range(1, 50 * len(json_data) + 1): # Assume max 50 rows per page for grid setting
        ws.row_dimensions[r].height = 17.01

    # 印刷設定
    ws.page_setup.paperSize = 9  # A4
    ws.page_setup.orientation = 'portrait'
    ws.page_margins.left = 0.47
    ws.page_margins.right = 0.47
    ws.page_margins.top = 0.41
    ws.page_margins.bottom = 0.41

    max_element_row = 0
    max_element_col = 0

    for page_data in json_data:
        page_number = page_data["page_number"]
        row_offset = (page_number - 1) * 50

        # ページ境界の改ページ設定
        if page_number > 1:
            try:
                ws.row_breaks.append(Break(id=row_offset))
            except AttributeError:
                pass # Older openpyxl versions might not support this directly

        for item in page_data["elements"]:
            if item["type"] == "text":
                r = item["row"] + row_offset
                c = item["col"]
                # Clamp coordinates
                r = min(r, 50)
                c = min(c, 36)

                if r > max_element_row: max_element_row = r
                if c > max_element_col: max_element_col = c

                cell = ws.cell(row=r, column=c)
                cell.value = item["content"]
                cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=False)
                font_kwargs = {}
                if item.get("font_color"):
                    font_kwargs["color"] = item["font_color"]
                if item.get("font_size"):
                    font_kwargs["size"] = item["font_size"]
                if font_kwargs:
                    cell.font = Font(**font_kwargs)

            elif item["type"] == "border_rect":
                # Clamp coordinates
                s_row = min(item["row"] + row_offset, 50)
                e_row = min(item["end_row"] + row_offset, 50)
                s_col = min(item["col"], 36)
                e_col = min(item["end_col"], 36)

                if e_row > max_element_row: max_element_row = e_row
                if e_col > max_element_col: max_element_col = e_col

                apply_outer_border(
                    ws,
                    s_row, e_row,
                    s_col, e_col,
                    fill_color=item.get("fill_color")
                )

    # 印刷範囲設定
    if max_element_row > 0 and max_element_col > 0:
        ws.print_area = f"A1:{get_column_letter(max_element_col)}{max_element_row}"
    else:
        ws.print_area = "A1" # Default if no elements found

    wb.save("output.xlsx")

# Input data from Step 1
data = [
  {
    "page_number": 1,
    "elements": [
      {
        "type": "border_rect",
        "row": 15,
        "end_row": 16,
        "col": 26,
        "end_col": 30
      },
      {
        "type": "border_rect",
        "row": 16,
        "end_row": 19,
        "col": 26,
        "end_col": 30
      },
      {
        "type": "border_rect",
        "row": 15,
        "end_row": 16,
        "col": 30,
        "end_col": 33
      },
      {
        "type": "border_rect",
        "row": 16,
        "end_row": 19,
        "col": 30,
        "end_col": 33
      },
      {
        "type": "border_rect",
        "row": 20,
        "end_row": 21,
        "col": 4,
        "end_col": 16
      },
      {
        "type": "border_rect",
        "row": 21,
        "end_row": 23,
        "col": 4,
        "end_col": 16
      },
      {
        "type": "border_rect",
        "row": 23,
        "end_row": 24,
        "col": 4,
        "end_col": 16
      },
      {
        "type": "border_rect",
        "row": 24,
        "end_row": 25,
        "col": 4,
        "end_col": 16
      },
      {
        "type": "border_rect",
        "row": 25,
        "end_row": 27,
        "col": 4,
        "end_col": 16
      },
      {
        "type": "border_rect",
        "row": 27,
        "end_row": 28,
        "col": 4,
        "end_col": 16
      },
      {
        "type": "border_rect",
        "row": 28,
        "end_row": 29,
        "col": 4,
        "end_col": 16
      },
      {
        "type": "border_rect",
        "row": 29,
        "end_row": 31,
        "col": 4,
        "end_col": 16
      },
      {
        "type": "border_rect",
        "row": 31,
        "end_row": 32,
        "col": 4,
        "end_col": 16
      },
      {
        "type": "border_rect",
        "row": 32,
        "end_row": 33,
        "col": 4,
        "end_col": 16
      },
      {
        "type": "border_rect",
        "row": 33,
        "end_row": 35,
        "col": 4,
        "end_col": 16
      },
      {
        "type": "border_rect",
        "row": 35,
        "end_row": 36,
        "col": 4,
        "end_col": 16
      },
      {
        "type": "border_rect",
        "row": 36,
        "end_row": 37,
        "col": 4,
        "end_col": 16
      },
      {
        "type": "border_rect",
        "row": 37,
        "end_row": 39,
        "col": 4,
        "end_col": 16
      },
      {
        "type": "border_rect",
        "row": 39,
        "end_row": 40,
        "col": 4,
        "end_col": 24
      },
      {
        "type": "border_rect",
        "row": 40,
        "end_row": 41,
        "col": 4,
        "end_col": 24
      },
      {
        "type": "border_rect",
        "row": 41,
        "end_row": 43,
        "col": 4,
        "end_col": 24
      },
      {
        "type": "border_rect",
        "row": 43,
        "end_row": 47,
        "col": 4,
        "end_col": 33
      },
      {
        "type": "border_rect",
        "row": 20,
        "end_row": 21,
        "col": 16,
        "end_col": 18
      },
      {
        "type": "border_rect",
        "row": 21,
        "end_row": 23,
        "col": 16,
        "end_col": 18
      },
      {
        "type": "border_rect",
        "row": 23,
        "end_row": 24,
        "col": 16,
        "end_col": 18
      },
      {
        "type": "border_rect",
        "row": 24,
        "end_row": 25,
        "col": 16,
        "end_col": 18
      },
      {
        "type": "border_rect",
        "row": 25,
        "end_row": 27,
        "col": 16,
        "end_col": 18
      },
      {
        "type": "border_rect",
        "row": 27,
        "end_row": 28,
        "col": 16,
        "end_col": 18
      },
      {
        "type": "border_rect",
        "row": 28,
        "end_row": 29,
        "col": 16,
        "end_col": 18
      },
      {
        "type": "border_rect",
        "row": 29,
        "end_row": 31,
        "col": 16,
        "end_col": 18
      },
      {
        "type": "border_rect",
        "row": 31,
        "end_row": 32,
        "col": 16,
        "end_col": 18
      },
      {
        "type": "border_rect",
        "row": 32,
        "end_row": 33,
        "col": 16,
        "end_col": 18
      },
      {
        "type": "border_rect",
        "row": 33,
        "end_row": 35,
        "col": 16,
        "end_col": 18
      },
      {
        "type": "border_rect",
        "row": 35,
        "end_row": 36,
        "col": 16,
        "end_col": 18
      },
      {
        "type": "border_rect",
        "row": 36,
        "end_row": 37,
        "col": 16,
        "end_col": 18
      },
      {
        "type": "border_rect",
        "row": 37,
        "end_row": 39,
        "col": 16,
        "end_col": 18
      },
      {
        "type": "border_rect",
        "row": 20,
        "end_row": 21,
        "col": 18,
        "end_col": 24
      },
      {
        "type": "border_rect",
        "row": 21,
        "end_row": 23,
        "col": 18,
        "end_col": 24
      },
      {
        "type": "border_rect",
        "row": 23,
        "end_row": 24,
        "col": 18,
        "end_col": 24
      },
      {
        "type": "border_rect",
        "row": 24,
        "end_row": 25,
        "col": 18,
        "end_col": 24
      },
      {
        "type": "border_rect",
        "row": 25,
        "end_row": 27,
        "col": 18,
        "end_col": 24
      },
      {
        "type": "border_rect",
        "row": 27,
        "end_row": 28,
        "col": 18,
        "end_col": 24
      },
      {
        "type": "border_rect",
        "row": 28,
        "end_row": 29,
        "col": 18,
        "end_col": 24
      },
      {
        "type": "border_rect",
        "row": 29,
        "end_row": 31,
        "col": 18,
        "end_col": 24
      },
      {
        "type": "border_rect",
        "row": 31,
        "end_row": 32,
        "col": 18,
        "end_col": 24
      },
      {
        "type": "border_rect",
        "row": 32,
        "end_row": 33,
        "col": 18,
        "end_col": 24
      },
      {
        "type": "border_rect",
        "row": 33,
        "end_row": 35,
        "col": 18,
        "end_col": 24
      },
      {
        "type": "border_rect",
        "row": 35,
        "end_row": 36,
        "col": 18,
        "end_col": 24
      },
      {
        "type": "border_rect",
        "row": 36,
        "end_row": 37,
        "col": 18,
        "end_col": 24
      },
      {
        "type": "border_rect",
        "row": 37,
        "end_row": 39,
        "col": 18,
        "end_col": 24
      },
      {
        "type": "border_rect",
        "row": 39,
        "end_row": 40,
        "col": 24,
        "end_col": 27
      },
      {
        "type": "border_rect",
        "row": 40,
        "end_row": 41,
        "col": 24,
        "end_col": 27
      },
      {
        "type": "border_rect",
        "row": 41,
        "end_row": 43,
        "col": 24,
        "end_col": 27
      },
      {
        "type": "border_rect",
        "row": 20,
        "end_row": 21,
        "col": 27,
        "end_col": 33
      },
      {
        "type": "border_rect",
        "row": 21,
        "end_row": 23,
        "col": 27,
        "end_col": 33
      },
      {
        "type": "border_rect",
        "row": 23,
        "end_row": 24,
        "col": 27,
        "end_col": 33
      },
      {
        "type": "border_rect",
        "row": 24,
        "end_row": 25,
        "col": 27,
        "end_col": 33
      },
      {
        "type": "border_rect",
        "row": 25,
        "end_row": 27,
        "col": 27,
        "end_col": 33
      },
      {
        "type": "border_rect",
        "row": 27,
        "end_row": 28,
        "col": 27,
        "end_col": 33
      },
      {
        "type": "border_rect",
        "row": 28,
        "end_row": 29,
        "col": 27,
        "end_col": 33
      },
      {
        "type": "border_rect",
        "row": 29,
        "end_row": 31,
        "col": 27,
        "end_col": 33
      },
      {
        "type": "border_rect",
        "row": 31,
        "end_row": 32,
        "col": 27,
        "end_col": 33
      },
      {
        "type": "border_rect",
        "row": 32,
        "end_row": 33,
        "col": 27,
        "end_col": 33
      },
      {
        "type": "border_rect",
        "row": 33,
        "end_row": 35,
        "col": 27,
        "end_col": 33
      },
      {
        "type": "border_rect",
        "row": 35,
        "end_row": 36,
        "col": 27,
        "end_col": 33
      },
      {
        "type": "border_rect",
        "row": 36,
        "end_row": 37,
        "col": 27,
        "end_col": 33
      },
      {
        "type": "border_rect",
        "row": 37,
        "end_row": 39,
        "col": 27,
        "end_col": 33
      },
      {
        "type": "border_rect",
        "row": 39,
        "end_row": 40,
        "col": 27,
        "end_col": 33
      },
      {
        "type": "border_rect",
        "row": 40,
        "end_row": 41,
        "col": 27,
        "end_col": 33
      },
      {
        "type": "border_rect",
        "row": 41,
        "end_row": 43,
        "col": 27,
        "end_col": 33
      },
      {
        "type": "text",
        "content": "請",
        "row": 5,
        "col": 14,
        "end_col": 15,
        "font_size": 20.0
      },
      {
        "type": "text",
        "content": "求",
        "row": 5,
        "col": 16,
        "end_col": 17,
        "font_size": 20.0
      },
      {
        "type": "text",
        "content": "書",
        "row": 5,
        "col": 21,
        "end_col": 22,
        "font_size": 20.0
      },
      {
        "type": "text",
        "content": "御請求金額",
        "row": 18,
        "col": 5,
        "end_col": 14,
        "font_size": 14.0
      },
      {
        "type": "text",
        "content": "¥32,400-",
        "row": 18,
        "col": 14,
        "end_col": 23,
        "font_size": 18.0
      },
      {
        "type": "text",
        "content": "（消費税込み）",
        "row": 18,
        "col": 21,
        "end_col": 29,
        "font_size": 11.0
      },
      {
        "type": "text",
        "content": "下記のとおりご請求いたします。",
        "row": 11,
        "col": 4,
        "end_col": 23,
        "font_size": 11.0
      },
      {
        "type": "text",
        "content": "（会社名）",
        "row": 9,
        "col": 21,
        "end_col": 26,
        "font_size": 11.0
      },
      {
        "type": "text",
        "content": "様",
        "row": 9,
        "col": 14,
        "end_col": 15,
        "font_size": 14.0
      },
      {
        "type": "text",
        "content": "〒",
        "row": 10,
        "col": 21,
        "end_col": 22,
        "font_size": 11.0
      },
      {
        "type": "text",
        "content": "（住所）",
        "row": 11,
        "col": 21,
        "end_col": 25,
        "font_size": 11.0
      },
      {
        "type": "text",
        "content": "TEL.",
        "row": 13,
        "col": 21,
        "end_col": 23,
        "font_size": 11.0
      },
      {
        "type": "text",
        "content": "FAX.",
        "row": 14,
        "col": 21,
        "end_col": 23,
        "font_size": 11.0
      },
      {
        "type": "text",
        "content": "（振込先）",
        "row": 13,
        "col": 5,
        "end_col": 10,
        "font_size": 11.0
      },
      {
        "type": "text",
        "content": "◯◯銀行◯◯支店",
        "row": 14,
        "col": 5,
        "end_col": 13,
        "font_size": 11.0
      },
      {
        "type": "text",
        "content": "預金種別：",
        "row": 14,
        "col": 5,
        "end_col": 10,
        "font_size": 11.0
      },
      {
        "type": "text",
        "content": "普通",
        "row": 14,
        "col": 10,
        "end_col": 12,
        "font_size": 11.0
      },
      {
        "type": "text",
        "content": "口座番号：",
        "row": 15,
        "col": 5,
        "end_col": 10,
        "font_size": 11.0
      },
      {
        "type": "text",
        "content": "1234567",
        "row": 15,
        "col": 10,
        "end_col": 13,
        "font_size": 11.0
      },
      {
        "type": "text",
        "content": "口座名義：",
        "row": 16,
        "col": 5,
        "end_col": 10,
        "font_size": 11.0
      },
      {
        "type": "text",
        "content": "◯◯株式会社",
        "row": 16,
        "col": 10,
        "end_col": 16,
        "font_size": 11.0
      },
      {
        "type": "text",
        "content": "品",
        "row": 20,
        "col": 10,
        "end_col": 11,
        "font_size": 11.0
      },
      {
        "type": "text",
        "content": "名",
        "row": 20,
        "col": 10,
        "end_col": 11,
        "font_size": 11.0
      },
      {
        "type": "text",
        "content": "数量",
        "row": 20,
        "col": 16,
        "end_col": 18,
        "font_size": 11.0
      },
      {
        "type": "text",
        "content": "単価",
        "row": 20,
        "col": 21,
        "end_col": 23,
        "font_size": 11.0
      },
      {
        "type": "text",
        "content": "金額",
        "row": 20,
        "col": 24,
        "end_col": 26,
        "font_size": 11.0
      },
      {
        "type": "text",
        "content": "摘要",
        "row": 20,
        "col": 29,
        "end_col": 31,
        "font_size": 11.0
      },
      {
        "type": "text",
        "content": "洗面台パイプ修理",
        "row": 21,
        "col": 4,
        "end_col": 12,
        "font_size": 11.0
      },
      {
        "type": "text",
        "content": "1",
        "row": 21,
        "col": 16,
        "end_col": 17,
        "font_size": 11.0
      },
      {
        "type": "text",
        "content": "21,000",
        "row": 21,
        "col": 21,
        "end_col": 25,
        "font_size": 11.0
      },
      {
        "type": "text",
        "content": "21,000",
        "row": 21,
        "col": 24,
        "end_col": 28,
        "font_size": 11.0
      },
      {
        "type": "text",
        "content": "サンプル",
        "row": 21,
        "col": 27,
        "end_col": 31,
        "font_size": 9.0
      },
      {
        "type": "text",
        "content": "清掃一式",
        "row": 23,
        "col": 4,
        "end_col": 8,
        "font_size": 11.0
      },
      {
        "type": "text",
        "content": "1",
        "row": 23,
        "col": 16,
        "end_col": 17,
        "font_size": 11.0
      },
      {
        "type": "text",
        "content": "9,000",
        "row": 23,
        "col": 21,
        "end_col": 24,
        "font_size": 11.0
      },
      {
        "type": "text",
        "content": "9,000",
        "row": 23,
        "col": 24,
        "end_col": 27,
        "font_size": 11.0
      },
      {
        "type": "text",
        "content": "サンプル",
        "row": 23,
        "col": 27,
        "end_col": 31,
        "font_size": 9.0
      },
      {
        "type": "text",
        "content": "小 計",
        "row": 39,
        "col": 12,
        "end_col": 14,
        "font_size": 11.0
      },
      {
        "type": "text",
        "content": "30,000",
        "row": 39,
        "col": 24,
        "end_col": 28,
        "font_size": 11.0
      },
      {
        "type": "text",
        "content": "消費税等",
        "row": 40,
        "col": 12,
        "end_col": 16,
        "font_size": 11.0
      },
      {
        "type": "text",
        "content": "2,400",
        "row": 40,
        "col": 24,
        "end_col": 27,
        "font_size": 11.0
      },
      {
        "type": "text",
        "content": "合 計",
        "row": 41,
        "col": 14,
        "end_col": 16,
        "font_size": 11.0
      },
      {
        "type": "text",
        "content": "32,400",
        "row": 41,
        "col": 24,
        "end_col": 28,
        "font_size": 11.0
      },
      {
        "type": "text",
        "content": "備考：",
        "row": 43,
        "col": 4,
        "end_col": 7,
        "font_size": 11.0
      },
      {
        "type": "text",
        "content": "検",
        "row": 15,
        "col": 26,
        "end_col": 27,
        "font_size": 11.0
      },
      {
        "type": "text",
        "content": "印",
        "row": 15,
        "col": 27,
        "end_col": 28,
        "font_size": 11.0
      },
      {
        "type": "text",
        "content": "担当者印",
        "row": 15,
        "col": 30,
        "end_col": 32,
        "font_size": 11.0
      }
    ]
  }
]

create_excel_from_json(data)