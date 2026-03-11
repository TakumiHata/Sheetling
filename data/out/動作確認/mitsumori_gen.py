import openpyxl
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
from openpyxl.utils import get_column_letter

def create_mitsumori_excel():
    # 1. ブックとシートの作成
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Page 1"

    # 入力データ (STEP 5 出力)
    print_range = "B2:AI49"
    data_list = [
        {"action": "merge_and_set", "start_row": 2, "start_column": 14, "end_row": 3, "end_column": 22, "value": "御見積書", "border": False},
        {"action": "merge_and_set", "start_row": 6, "start_column": 24, "end_row": 6, "end_column": 35, "value": "見積書Ｎｏ．20121119_00001", "border": False},
        {"action": "merge_and_set", "start_row": 7, "start_column": 2, "end_row": 8, "end_column": 15, "value": "△△株式会社 御中", "border": False},
        {"action": "merge_and_set", "start_row": 8, "start_column": 24, "end_row": 8, "end_column": 35, "value": "作成日: 2012年11月19日", "border": False},
        {"action": "merge_and_set", "start_row": 10, "start_column": 4, "end_row": 11, "end_column": 20, "value": "下記のとおり御見積申し上げます。\n何卒ご用命の程、お願い申し上げます。", "border": False},
        {"action": "merge_and_set", "start_row": 13, "start_column": 2, "end_row": 13, "end_column": 12, "value": "受渡期日:別途御打合せ", "border": False},
        {"action": "merge_and_set", "start_row": 13, "start_column": 24, "end_row": 13, "end_column": 35, "value": "ワークフロー商事株式会社", "border": False},
        {"action": "merge_and_set", "start_row": 14, "start_column": 24, "end_row": 15, "end_column": 35, "value": "東京都新宿区千代田９－９－９\nWSビル", "border": False},
        {"action": "merge_and_set", "start_row": 16, "start_column": 24, "end_row": 16, "end_column": 35, "value": "TEL ０３-××××－９９９９", "border": False},
        {"action": "merge_and_set", "start_row": 17, "start_column": 24, "end_row": 17, "end_column": 35, "value": "FAX ０３-９９９９－××××", "border": False},
        {"action": "merge_and_set", "start_row": 15, "start_column": 2, "end_row": 15, "end_column": 12, "value": "取引方法:別途御打合せ", "border": False},
        {"action": "merge_and_set", "start_row": 17, "start_column": 2, "end_row": 17, "end_column": 12, "value": "有効期限:発行日から30日", "border": False},
        {"action": "merge_and_set", "start_row": 19, "start_column": 2, "end_row": 19, "end_column": 12, "value": "貴社管理番号：", "border": False},
        {"action": "merge_and_set", "start_row": 19, "start_column": 28, "end_row": 19, "end_column": 35, "value": "| 承認 | 担当営業 |", "border": True},
        {"action": "merge_and_set", "start_row": 20, "start_column": 28, "end_row": 21, "end_column": 35, "value": "| | |", "border": True},
        {"action": "merge_and_set", "start_row": 23, "start_column": 5, "end_row": 23, "end_column": 18, "value": "合計金額： \\3,700,000", "border": False},
        {"action": "merge_and_set", "start_row": 23, "start_column": 19, "end_row": 23, "end_column": 25, "value": "(消費税別)", "border": False},
        {"action": "merge_and_set", "start_row": 25, "start_column": 2, "end_row": 25, "end_column": 35, "value": "| No. | 摘 要 | 数量 | 標準価格 | 見積価格 | 合計金額 |", "border": True},
        {"action": "merge_and_set", "start_row": 26, "start_column": 2, "end_row": 26, "end_column": 35, "value": "| 1 | ワークフローシステム 30ユーザーライセンス | 1 | 2,700,000 | 2,700,000 | 2,700,000 |", "border": True},
        {"action": "merge_and_set", "start_row": 27, "start_column": 2, "end_row": 27, "end_column": 35, "value": "| 2 | 初期設定費用 | 1 | 500,000 | 500,000 | 500,000 |", "border": True},
        {"action": "merge_and_set", "start_row": 28, "start_column": 2, "end_row": 28, "end_column": 35, "value": "| 3 | 管理者費用 | 1 | 500,000 | 500,000 | 500,000 |", "border": True},
        {"action": "merge_and_set", "start_row": 29, "start_column": 2, "end_row": 29, "end_column": 35, "value": "| 4 | | | | | |", "border": True},
        {"action": "merge_and_set", "start_row": 30, "start_column": 2, "end_row": 30, "end_column": 35, "value": "| 5 | | | | | |", "border": True},
        {"action": "merge_and_set", "start_row": 31, "start_column": 2, "end_row": 31, "end_column": 35, "value": "| 6 | | | | | |", "border": True},
        {"action": "merge_and_set", "start_row": 32, "start_column": 2, "end_row": 32, "end_column": 35, "value": "| 7 | | | | | |", "border": True},
        {"action": "merge_and_set", "start_row": 33, "start_column": 2, "end_row": 33, "end_column": 35, "value": "| 8 | | | | | |", "border": True},
        {"action": "merge_and_set", "start_row": 34, "start_column": 2, "end_row": 34, "end_column": 35, "value": "| 9 | | | | | |", "border": True},
        {"action": "merge_and_set", "start_row": 35, "start_column": 2, "end_row": 35, "end_column": 35, "value": "| 10 | | | | | |", "border": True},
        {"action": "merge_and_set", "start_row": 36, "start_column": 2, "end_row": 36, "end_column": 35, "value": "| 11 | | | | | |", "border": True},
        {"action": "merge_and_set", "start_row": 37, "start_column": 2, "end_row": 37, "end_column": 35, "value": "| 12 | | | | | |", "border": True},
        {"action": "merge_and_set", "start_row": 38, "start_column": 2, "end_row": 38, "end_column": 35, "value": "| 13 | | | | | |", "border": True},
        {"action": "merge_and_set", "start_row": 39, "start_column": 2, "end_row": 39, "end_column": 35, "value": "| 14 | | | | | |", "border": True},
        {"action": "merge_and_set", "start_row": 40, "start_column": 2, "end_row": 40, "end_column": 35, "value": "| 15 | | | | | |", "border": True},
        {"action": "merge_and_set", "start_row": 41, "start_column": 2, "end_row": 41, "end_column": 35, "value": "| 16 | | | | | |", "border": True},
        {"action": "merge_and_set", "start_row": 42, "start_column": 2, "end_row": 42, "end_column": 35, "value": "| 17 | | | | | |", "border": True},
        {"action": "merge_and_set", "start_row": 43, "start_column": 2, "end_row": 43, "end_column": 35, "value": "| 18 | | | | | |", "border": True},
        {"action": "merge_and_set", "start_row": 44, "start_column": 2, "end_row": 44, "end_column": 35, "value": "| 19 | | | | | |", "border": True},
        {"action": "merge_and_set", "start_row": 45, "start_column": 2, "end_row": 45, "end_column": 35, "value": "| 20 | | | | | |", "border": True},
        {"action": "merge_and_set", "start_row": 46, "start_column": 2, "end_row": 46, "end_column": 35, "value": "| | 合 計 | | | | 3,700,000 |", "border": True},
        {"action": "merge_and_set", "start_row": 47, "start_column": 2, "end_row": 47, "end_column": 35, "value": "| 備 考 |", "border": True},
        {"action": "merge_and_set", "start_row": 48, "start_column": 2, "end_row": 48, "end_column": 35, "value": "| ・消費税は別途計上させていただきます。 |", "border": True},
        {"action": "merge_and_set", "start_row": 49, "start_column": 2, "end_row": 49, "end_column": 35, "value": "| ・製品の瑕疵、無償保証期間は御購入後3ヶ月間です。 |", "border": True},
    ]

    # 2. レイアウト設定 (1列=6mm, 1行=6mm)
    for col in range(1, 40):
        ws.column_dimensions[get_column_letter(col)].width = 2.53
    for row in range(1, 60):
        ws.row_dimensions[row].height = 17.01

    # スタイル定義
    thin_border = Border(
        left=Side(style='thin'), 
        right=Side(style='thin'), 
        top=Side(style='thin'), 
        bottom=Side(style='thin')
    )
    header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    bold_font = Font(bold=True)

    # 3. データ描画
    for item in data_list:
        sr, sc = item["start_row"], item["start_column"]
        er, ec = item["end_row"], item["end_column"]
        val = item["value"]

        # 結合処理
        if sr != er or sc != ec:
            ws.merge_cells(start_row=sr, start_column=sc, end_row=er, end_column=ec)
        
        # 値のセット
        cell = ws.cell(row=sr, column=sc)
        cell.value = val
        
        # 基本配置スタイル
        cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='left')
        
        # 見出し用スタイル（「御見積書」など）
        if "heading" in str(item.get("type", "")):
            cell.font = Font(size=16, bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center')

        # 罫線の適用
        if item.get("border"):
            for r in range(sr, er + 1):
                for c in range(sc, ec + 1):
                    ws.cell(row=r, column=c).border = thin_border
            
            # テーブルヘッダーの強調 (No. 摘要... など)
            if "No." in val or "承認" in val or "備 考" in val:
                cell.fill = header_fill
                cell.font = bold_font
                cell.alignment = Alignment(horizontal='center', vertical='center')

    # 4. 印刷設定 (絶対等倍・中央配置)
    ws.page_setup.paperSize = 9  # A4
    ws.page_setup.orientation = 'portrait' # 縦
    
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = 0.47
    ws.page_margins.right = 0.47
    ws.page_margins.top = 0.41
    ws.page_margins.bottom = 0.41
    
    if print_range:
        ws.print_area = print_range

    # 保存
    wb.save("output.xlsx")
    print("Excel file 'output.xlsx' has been created successfully.")

if __name__ == "__main__":
    create_mitsumori_excel()