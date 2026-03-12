import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Border, Side, Alignment
from openpyxl.worksheet.pagebreak import Break
import json

# Input data from Step 1
data = [
  {
    "page_number": 1,
    "elements": [
      {
        "type": "border_rect",
        "row": 15,
        "end_row": 16,
        "col": 27,
        "end_col": 30
      },
      {
        "type": "border_rect",
        "row": 16,
        "end_row": 19,
        "col": 27,
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
        "end_col": 22
      },
      {
        "type": "border_rect",
        "row": 40,
        "end_row": 41,
        "col": 4,
        "end_col": 22
      },
      {
        "type": "border_rect",
        "row": 41,
        "end_row": 43,
        "col": 4,
        "end_col": 22
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
        "end_col": 16
      },
      {
        "type": "border_rect",
        "row": 21,
        "end_row": 23,
        "col": 16,
        "end_col": 16
      },
      {
        "type": "border_rect",
        "row": 23,
        "end_row": 24,
        "col": 16,
        "end_col": 16
      },
      {
        "type": "border_rect",
        "row": 24,
        "end_row": 25,
        "col": 16,
        "end_col": 16
      },
      {
        "type": "border_rect",
        "row": 25,
        "end_row": 27,
        "col": 16,
        "end_col": 16
      },
      {
        "type": "border_rect",
        "row": 27,
        "end_row": 28,
        "col": 16,
        "end_col": 16
      },
      {
        "type": "border_rect",
        "row": 28,
        "end_row": 29,
        "col": 16,
        "end_col": 16
      },
      {
        "type": "border_rect",
        "row": 29,
        "end_row": 31,
        "col": 16,
        "end_col": 16
      },
      {
        "type": "border_rect",
        "row": 31,
        "end_row": 32,
        "col": 16,
        "end_col": 16
      },
      {
        "type": "border_rect",
        "row": 32,
        "end_row": 33,
        "col": 16,
        "end_col": 16
      },
      {
        "type": "border_rect",
        "row": 33,
        "end_row": 35,
        "col": 16,
        "end_col": 16
      },
      {
        "type": "border_rect",
        "row": 35,
        "end_row": 36,
        "col": 16,
        "end_col": 16
      },
      {
        "type": "border_rect",
        "row": 36,
        "end_row": 37,
        "col": 16,
        "end_col": 16
      },
      {
        "type": "border_rect",
        "row": 37,
        "end_row": 39,
        "col": 16,
        "end_col": 16
      },
      {
        "type": "border_rect",
        "row": 20,
        "end_row": 21,
        "col": 16,
        "end_col": 22
      },
      {
        "type": "border_rect",
        "row": 21,
        "end_row": 23,
        "col": 16,
        "end_col": 22
      },
      {
        "type": "border_rect",
        "row": 23,
        "end_row": 24,
        "col": 16,
        "end_col": 22
      },
      {
        "type": "border_rect",
        "row": 24,
        "end_row": 25,
        "col": 16,
        "end_col": 22
      },
      {
        "type": "border_rect",
        "row": 25,
        "end_row": 27,
        "col": 16,
        "end_col": 22
      },
      {
        "type": "border_rect",
        "row": 27,
        "end_row": 28,
        "col": 16,
        "end_col": 22
      },
      {
        "type": "border_rect",
        "row": 28,
        "end_row": 29,
        "col": 16,
        "end_col": 22
      },
      {
        "type": "border_rect",
        "row": 29,
        "end_row": 31,
        "col": 16,
        "end_col": 22
      },
      {
        "type": "border_rect",
        "row": 31,
        "end_row": 32,
        "col": 16,
        "end_col": 22
      },
      {
        "type": "border_rect",
        "row": 32,
        "end_row": 33,
        "col": 16,
        "end_col": 22
      },
      {
        "type": "border_rect",
        "row": 33,
        "end_row": 35,
        "col": 16,
        "end_col": 22
      },
      {
        "type": "border_rect",
        "row": 35,
        "end_row": 36,
        "col": 16,
        "end_col": 22
      },
      {
        "type": "border_rect",
        "row": 36,
        "end_row": 37,
        "col": 16,
        "end_col": 22
      },
      {
        "type": "border_rect",
        "row": 37,
        "end_row": 39,
        "col": 16,
        "end_col": 22
      },
      {
        "type": "border_rect",
        "row": 20,
        "end_row": 21,
        "col": 22,
        "end_col": 27
      },
      {
        "type": "border_rect",
        "row": 21,
        "end_row": 23,
        "col": 22,
        "end_col": 27
      },
      {
        "type": "border_rect",
        "row": 23,
        "end_row": 24,
        "col": 22,
        "end_col": 27
      },
      {
        "type": "border_rect",
        "row": 24,
        "end_row": 25,
        "col": 22,
        "end_col": 27
      },
      {
        "type": "border_rect",
        "row": 25,
        "end_row": 27,
        "col": 22,
        "end_col": 27
      },
      {
        "type": "border_rect",
        "row": 27,
        "end_row": 28,
        "col": 22,
        "end_col": 27
      },
      {
        "type": "border_rect",
        "row": 28,
        "end_row": 29,
        "col": 22,
        "end_col": 27
      },
      {
        "type": "border_rect",
        "row": 29,
        "end_row": 31,
        "col": 22,
        "end_col": 27
      },
      {
        "type": "border_rect",
        "row": 31,
        "end_row": 32,
        "col": 22,
        "end_col": 27
      },
      {
        "type": "border_rect",
        "row": 32,
        "end_row": 33,
        "col": 22,
        "end_col": 27
      },
      {
        "type": "border_rect",
        "row": 33,
        "end_row": 35,
        "col": 22,
        "end_col": 27
      },
      {
        "type": "border_rect",
        "row": 35,
        "end_row": 36,
        "col": 22,
        "end_col": 27
      },
      {
        "type": "border_rect",
        "row": 36,
        "end_row": 37,
        "col": 22,
        "end_col": 27
      },
      {
        "type": "border_rect",
        "row": 37,
        "end_row": 39,
        "col": 22,
        "end_col": 27
      },
      {
        "type": "border_rect",
        "row": 39,
        "end_row": 40,
        "col": 22,
        "end_col": 27
      },
      {
        "type": "border_rect",
        "row": 40,
        "end_row": 41,
        "col": 22,
        "end_col": 27
      },
      {
        "type": "border_rect",
        "row": 41,
        "end_row": 43,
        "col": 22,
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
        "col": 16,
        "end_col": 18
      },
      {
        "type": "text",
        "content": "求",
        "row": 5,
        "col": 16,
        "end_col": 18
      },
      {
        "type": "text",
        "content": "書",
        "row": 5,
        "col": 22,
        "end_col": 24
      },
      {
        "type": "text",
        "content": "年",
        "row": 7,
        "col": 27,
        "end_col": 28
      },
      {
        "type": "text",
        "content": "月",
        "row": 7,
        "col": 30,
        "end_col": 31
      },
      {
        "type": "text",
        "content": "日",
        "row": 7,
        "col": 33,
        "end_col": 34
      },
      {
        "type": "text",
        "content": "様",
        "row": 9,
        "col": 16,
        "end_col": 17
      },
      {
        "type": "text",
        "content": "（会社名）",
        "row": 9,
        "col": 22,
        "end_col": 27
      },
      {
        "type": "text",
        "content": "〒",
        "row": 10,
        "col": 22,
        "end_col": 23
      },
      {
        "type": "text",
        "content": "下記のとおりご請求いたします。",
        "row": 11,
        "col": 4,
        "end_col": 17
      },
      {
        "type": "text",
        "content": "（住所）",
        "row": 11,
        "col": 22,
        "end_col": 25
      },
      {
        "type": "text",
        "content": "（振込先）",
        "row": 13,
        "col": 5,
        "end_col": 10
      },
      {
        "type": "text",
        "content": "TEL.",
        "row": 13,
        "col": 22,
        "end_col": 24
      },
      {
        "type": "text",
        "content": "◯◯銀行◯◯支店",
        "row": 14,
        "col": 5,
        "end_col": 13
      },
      {
        "type": "text",
        "content": "FAX.",
        "row": 14,
        "col": 22,
        "end_col": 24
      },
      {
        "type": "text",
        "content": "預金種別：普通",
        "row": 14,
        "col": 5,
        "end_col": 12
      },
      {
        "type": "text",
        "content": "口座番号：1234567",
        "row": 15,
        "col": 5,
        "end_col": 16
      },
      {
        "type": "text",
        "content": "検",
        "row": 15,
        "col": 27,
        "end_col": 28
      },
      {
        "type": "text",
        "content": "印",
        "row": 15,
        "col": 27,
        "end_col": 28
      },
      {
        "type": "text",
        "content": "担当者印",
        "row": 15,
        "col": 30,
        "end_col": 32
      },
      {
        "type": "text",
        "content": "口座名義：◯◯株式会社",
        "row": 16,
        "col": 5,
        "end_col": 15
      },
      {
        "type": "text",
        "content": "御請求金額",
        "row": 18,
        "col": 5,
        "end_col": 11
      },
      {
        "type": "text",
        "content": "¥32,400-",
        "row": 18,
        "col": 16,
        "end_col": 23
      },
      {
        "type": "text",
        "content": "（消費税込み）",
        "row": 18,
        "col": 22,
        "end_col": 28
      },
      {
        "type": "text",
        "content": "品",
        "row": 20,
        "col": 8,
        "end_col": 9
      },
      {
        "type": "text",
        "content": "名",
        "row": 20,
        "col": 10,
        "end_col": 11
      },
      {
        "type": "text",
        "content": "数量",
        "row": 20,
        "col": 16,
        "end_col": 18
      },
      {
        "type": "text",
        "content": "単価",
        "row": 20,
        "col": 22,
        "end_col": 24
      },
      {
        "type": "text",
        "content": "金額",
        "row": 20,
        "col": 22,
        "end_col": 24
      },
      {
        "type": "text",
        "content": "摘要",
        "row": 20,
        "col": 30,
        "end_col": 32
      },
      {
        "type": "text",
        "content": "洗面台パイプ修理",
        "row": 21,
        "col": 4,
        "end_col": 12
      },
      {
        "type": "text",
        "content": "1",
        "row": 21,
        "col": 16,
        "end_col": 17
      },
      {
        "type": "text",
        "content": "21,000",
        "row": 21,
        "col": 22,
        "end_col": 26
      },
      {
        "type": "text",
        "content": "21,000",
        "row": 21,
        "col": 22,
        "end_col": 26
      },
      {
        "type": "text",
        "content": "サンプル",
        "row": 21,
        "col": 27,
        "end_col": 30
      },
      {
        "type": "text",
        "content": "清掃一式",
        "row": 23,
        "col": 4,
        "end_col": 8
      },
      {
        "type": "text",
        "content": "1",
        "row": 23,
        "col": 16,
        "end_col": 17
      },
      {
        "type": "text",
        "content": "9,000",
        "row": 23,
        "col": 22,
        "end_col": 25
      },
      {
        "type": "text",
        "content": "9,000",
        "row": 23,
        "col": 22,
        "end_col": 25
      },
      {
        "type": "text",
        "content": "サンプル",
        "row": 23,
        "col": 27,
        "end_col": 30
      },
      {
        "type": "text",
        "content": "小",
        "row": 39,
        "col": 12,
        "end_col": 13
      },
      {
        "type": "text",
        "content": "計",
        "row": 39,
        "col": 16,
        "end_col": 17
      },
      {
        "type": "text",
        "content": "30,000",
        "row": 39,
        "col": 22,
        "end_col": 26
      },
      {
        "type": "text",
        "content": "消費税等",
        "row": 40,
        "col": 12,
        "end_col": 16
      },
      {
        "type": "text",
        "content": "2,400",
        "row": 40,
        "col": 22,
        "end_col": 26
      },
      {
        "type": "text",
        "content": "合",
        "row": 41,
        "col": 12,
        "end_col": 13
      },
      {
        "type": "text",
        "content": "計",
        "row": 41,
        "col": 16,
        "end_col": 17
      },
      {
        "type": "text",
        "content": "32,400",
        "row": 41,
        "col": 22,
        "end_col": 26
      },
      {
        "type": "text",
        "content": "備考：",
        "row": 43,
        "col": 4,
        "end_col": 7
      }
    ]
  }
]

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Sheet1"

# Grid settings
for c in range(1, 37):
    ws.column_dimensions[get_column_letter(c)].width = 2.53
for r in range(1, 51): # Assuming a single page for now, adjust if more pages are expected
    ws.row_dimensions[r].height = 17.01

# Border style
thin = Side(style='thin')

def apply_outer_border(ws, s_row, e_row, s_col, e_col):
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

row_offset = 0
max_row_num = 0
max_col_num = 0

for page in data:
    for item in page["elements"]:
        if item["type"] == "text":
            r = item["row"] + row_offset
            try:
                cell = ws.cell(row=r, column=item["col"])
                cell.value = item["content"]
                cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=False)
            except AttributeError:
                pass
            if r > max_row_num:
                max_row_num = r
            if item["col"] > max_col_num:
                max_col_num = item["col"]

        elif item["type"] == "border_rect":
            apply_outer_border(
                ws,
                item["row"] + row_offset, item["end_row"] + row_offset,
                item["col"], item["end_col"]
            )
            if item["end_row"] + row_offset > max_row_num:
                max_row_num = item["end_row"] + row_offset
            if item["end_col"] > max_col_num:
                max_col_num = item["end_col"]

    # Page break for subsequent pages
    if page["page_number"] > 1:
        row_offset += 50 # Assuming 50 rows per page for grid calculation
        ws.row_breaks.append(Break(id=(page["page_number"] - 1) * 50))

# Set print settings
ws.page_setup.paperSize = 9  # A4
ws.page_setup.orientation = 'portrait'
ws.page_margins.left = 0.47
ws.page_margins.right = 0.47
ws.page_margins.top = 0.41
ws.page_margins.bottom = 0.41

# Determine print range
# The problem statement specifies to use the maximum used row/column from the elements
# not rely on total pages * 50 for print range.
if max_row_num > 0 and max_col_num > 0:
    max_col_letter = get_column_letter(max_col_num)
    ws.print_area = f"A1:{max_col_letter}{max_row_num}"
else:
    # Fallback if no elements are processed (should not happen with provided data)
    ws.print_area = "A1:Z100" # Default reasonable range

# Save the workbook
wb.save("output.xlsx")