import openpyxl
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill

def create_excel():
    # データの定義 (STEP 5 の出力結果)
    input_data = [
        {
            "page_number": 1,
            "width": 595.0,
            "height": 842.0,
            "print_range": "B2:Z38",
            "data": [
                {"action": "merge_and_set", "start_row": 2, "start_column": 11, "end_row": 3, "end_column": 16, "value": "御見積書", "border": False},
                {"action": "merge_and_set", "start_row": 4, "start_column": 18, "end_row": 4, "end_column": 26, "value": "見積書Ｎｏ．20121119_00001", "border": False},
                {"action": "merge_and_set", "start_row": 5, "start_column": 2, "end_row": 5, "end_column": 12, "value": "△△株式会社 御中", "border": False},
                {"action": "merge_and_set", "start_row": 6, "start_column": 18, "end_row": 6, "end_column": 26, "value": "作成日: 2012年11月19日", "border": False},
                {"action": "merge_and_set", "start_row": 8, "start_column": 4, "end_row": 9, "end_column": 15, "value": "下記のとおり御見積申し上げます。\n何卒ご用命の程、お願い申し上げます。", "border": False},
                {"action": "merge_and_set", "start_row": 10, "start_column": 2, "end_row": 10, "end_column": 8, "value": "受渡期日:別途御打合せ", "border": False},
                {"action": "merge_and_set", "start_row": 10, "start_column": 18, "end_row": 10, "end_column": 26, "value": "ワークフロー商事株式会社", "border": False},
                {"action": "merge_and_set", "start_row": 11, "start_column": 2, "end_row": 11, "end_column": 8, "value": "取引方法:別途御打合せ", "border": False},
                {"action": "merge_and_set", "start_row": 11, "start_column": 18, "end_row": 12, "end_column": 26, "value": "東京都新宿区千代田９−９−９\nWSビル", "border": False},
                {"action": "merge_and_set", "start_row": 12, "start_column": 2, "end_row": 12, "end_column": 9, "value": "有効期限:発行日から30日", "border": False},
                {"action": "merge_and_set", "start_row": 13, "start_column": 18, "end_row": 14, "end_column": 26, "value": "TEL ０３-××××−９９９９\nFAX ０３-９９９９−××××", "border": False},
                {"action": "merge_and_set", "start_row": 14, "start_column": 2, "end_row": 14, "end_column": 6, "value": "貴社管理番号：", "border": False},
                {"action": "merge_and_set", "start_row": 14, "start_column": 21, "end_row": 14, "end_column": 23, "value": "承認", "border": True},
                {"action": "merge_and_set", "start_row": 14, "start_column": 24, "end_row": 14, "end_column": 26, "value": "担当営業", "border": True},
                {"action": "merge_and_set", "start_row": 15, "start_column": 21, "end_row": 16, "end_column": 23, "value": "", "border": True},
                {"action": "merge_and_set", "start_row": 15, "start_column": 24, "end_row": 16, "end_column": 26, "value": "", "border": True},
                {"action": "merge_and_set", "start_row": 17, "start_column": 6, "end_row": 18, "end_column": 18, "value": "合計金額： \\3,700,000\n(消費税別)", "border": False},
                {"action": "merge_and_set", "start_row": 19, "start_column": 2, "end_row": 19, "end_column": 3, "value": "No.", "border": True},
                {"action": "merge_and_set", "start_row": 19, "start_column": 4, "end_row": 19, "end_column": 15, "value": "摘 要", "border": True},
                {"action": "merge_and_set", "start_row": 19, "start_column": 16, "end_row": 19, "end_column": 17, "value": "数量", "border": True},
                {"action": "merge_and_set", "start_row": 19, "start_column": 18, "end_row": 19, "end_column": 20, "value": "標準価格", "border": True},
                {"action": "merge_and_set", "start_row": 19, "start_column": 21, "end_row": 19, "end_column": 23, "value": "見積価格", "border": True},
                {"action": "merge_and_set", "start_row": 19, "start_column": 24, "end_row": 19, "end_column": 26, "value": "合計金額", "border": True},
                {"action": "merge_and_set", "start_row": 20, "start_column": 2, "end_row": 20, "end_column": 3, "value": "1", "border": True},
                {"action": "merge_and_set", "start_row": 20, "start_column": 4, "end_row": 20, "end_column": 15, "value": "ワークフローシステム 30ユーザーライセンス", "border": True},
                {"action": "merge_and_set", "start_row": 20, "start_column": 16, "end_row": 20, "end_column": 17, "value": "1", "border": True},
                {"action": "merge_and_set", "start_row": 20, "start_column": 18, "end_row": 20, "end_column": 20, "value": "2,700,000", "border": True},
                {"action": "merge_and_set", "start_row": 20, "start_column": 21, "end_row": 20, "end_column": 23, "value": "2,700,000", "border": True},
                {"action": "merge_and_set", "start_row": 20, "start_column": 24, "end_row": 20, "end_column": 26, "value": "2,700,000", "border": True},
                {"action": "merge_and_set", "start_row": 21, "start_column": 2, "end_row": 21, "end_column": 3, "value": "2", "border": True},
                {"action": "merge_and_set", "start_row": 21, "start_column": 4, "end_row": 21, "end_column": 15, "value": "初期設定費用", "border": True},
                {"action": "merge_and_set", "start_row": 21, "start_column": 16, "end_row": 21, "end_column": 17, "value": "1", "border": True},
                {"action": "merge_and_set", "start_row": 21, "start_column": 18, "end_row": 21, "end_column": 20, "value": "500,000", "border": True},
                {"action": "merge_and_set", "start_row": 21, "start_column": 21, "end_row": 21, "end_column": 23, "value": "500,000", "border": True},
                {"action": "merge_and_set", "start_row": 21, "start_column": 24, "end_row": 21, "end_column": 26, "value": "500,000", "border": True},
                {"action": "merge_and_set", "start_row": 22, "start_column": 2, "end_row": 22, "end_column": 3, "value": "3", "border": True},
                {"action": "merge_and_set", "start_row": 22, "start_column": 4, "end_row": 22, "end_column": 15, "value": "管理者費用", "border": True},
                {"action": "merge_and_set", "start_row": 22, "start_column": 16, "end_row": 22, "end_column": 17, "value": "1", "border": True},
                {"action": "merge_and_set", "start_row": 22, "start_column": 18, "end_row": 22, "end_column": 20, "value": "500,000", "border": True},
                {"action": "merge_and_set", "start_row": 22, "start_column": 21, "end_row": 22, "end_column": 23, "value": "500,000", "border": True},
                {"action": "merge_and_set", "start_row": 22, "start_column": 24, "end_row": 22, "end_column": 26, "value": "500,000", "border": True},
                {"action": "merge_and_set", "start_row": 34, "start_column": 4, "end_row": 34, "end_column": 23, "value": "合計", "border": True},
                {"action": "merge_and_set", "start_row": 34, "start_column": 24, "end_row": 34, "end_column": 26, "value": "3,700,000", "border": True},
                {"action": "merge_and_set", "start_row": 35, "start_column": 2, "end_row": 35, "end_column": 26, "value": "備 考", "border": True},
                {"action": "merge_and_set", "start_row": 36, "start_column": 2, "end_row": 38, "end_column": 26, "value": "・消費税は別途計上させていただきます。\n・製品の瑕疵、無償保証期間は御購入後3ヶ月間です。", "border": True},
            ]
        }
    ]

    wb = openpyxl.Workbook()
    ws = wb.active

    # スタイル設定
    thin_side = Side(style='thin', color='000000')
    border_style = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    header_fill = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
    header_font = Font(bold=True)

    # 方眼紙サイズ設定 (1-26列, 1-38行)
    for col in range(1, 27):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 3.61
    for row in range(1, 39):
        ws.row_dimensions[row].height = 22.68

    page_info = input_data[0]
    
    # 描画処理
    for item in page_info["data"]:
        s_row = item["start_row"]
        s_col = item["start_column"]
        e_row = item["end_row"]
        e_col = item["end_column"]
        val = item["value"]
        has_border = item.get("border", False)

        try:
            cell = ws.cell(row=s_row, column=s_col)
            cell.value = val
            
            # 結合
            if s_row != e_row or s_col != e_col:
                ws.merge_cells(start_row=s_row, start_column=s_col, end_row=e_row, end_column=e_col)
            
            # 基本アラインメント
            alignment = Alignment(wrap_text=True, vertical='center', horizontal='left')
            if "御見積書" in str(val):
                alignment.horizontal = 'center'
                cell.font = Font(size=16, bold=True)
            
            # ヘッダー判定 (例: 明細ヘッダーや備考タイトル)
            if val in ["No.", "摘 要", "数量", "標準価格", "見積価格", "合計金額", "承認", "担当営業", "備 考"]:
                cell.fill = header_fill
                cell.font = header_font
                alignment.horizontal = 'center'

            # 罫線とアラインメントの適用
            for r in range(s_row, e_row + 1):
                for c in range(s_col, e_col + 1):
                    target_cell = ws.cell(row=r, column=c)
                    target_cell.alignment = alignment
                    if has_border:
                        target_cell.border = border_style

        except AttributeError:
            pass

    # 印刷設定
    ws.page_setup.paperSize = 9  # A4
    ws.page_setup.orientation = 'portrait'
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = 0.51
    ws.page_margins.right = 0.51
    ws.page_margins.top = 0.49
    ws.page_margins.bottom = 0.49
    
    if "print_range" in page_info:
        ws.print_area = page_info["print_range"]

    wb.save("output.xlsx")
    print("Excel file 'output.xlsx' has been generated.")

if __name__ == "__main__":
    create_excel()