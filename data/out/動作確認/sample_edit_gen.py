import json
from openpyxl import Workbook
from openpyxl.styles import Border, Side, Alignment, Font, PatternFill
from openpyxl.worksheet.pagebreak import Break

def create_excel_from_json(json_data):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # 1. レイアウトと方眼紙サイズの厳密な設定
    max_page = max(page['page_number'] for page in json_data)
    total_rows = max_page * 50

    # 列幅の設定 (1列＝約6.0mm)
    for col in range(1, 37):
        ws.column_dimensions[ws.cell(row=1, column=col).column_letter].width = 2.53

    # 行の高さの設定 (1行＝約6.0mm)
    for row in range(1, total_rows + 1):
        ws.row_dimensions[row].height = 17.01

    # スタイルの定義
    border_side = Side(style='thin')
    border_style = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)
    
    header_fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
    header_font = Font(bold=True)
    
    alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')

    # 4. 複数ページへの対応とデータ出力
    for page in json_data:
        page_num = page['page_number']
        offset = (page_num - 1) * 50
        
        # 改ページ設定
        if page_num > 1:
            ws.row_breaks.append(Break(id=offset))

        for item in page['data']:
            s_row = item['start_row'] + offset
            e_row = item['end_row'] + offset
            s_col = item['start_column']
            e_col = item['end_column']
            val = item['value']
            has_border = item.get('border', False)

            # 5. 技術的制約の遵守 (AttributeError 回避)
            try:
                cell = ws.cell(row=s_row, column=s_col)
                cell.value = val
                
                # 結合処理
                if s_row != e_row or s_col != e_col:
                    ws.merge_cells(start_row=s_row, start_column=s_col, end_row=e_row, end_column=e_col)
                
                # スタイリング
                # 結合セル全体に罫線を適用するための処理
                if has_border:
                    for r in range(s_row, e_row + 1):
                        for c in range(s_col, e_col + 1):
                            ws.cell(row=r, column=c).border = border_style
                    
                    # テーブルヘッダーの装飾 (1ページ目の特定の行など)
                    if page_num == 1 and s_row - offset == 9:
                        cell.fill = header_fill
                        cell.font = header_font
                
                cell.alignment = alignment

            except AttributeError:
                pass

    # 3. 印刷設定 (絶対等倍)
    ws.page_setup.paperSize = 9  # A4
    ws.page_setup.orientation = 'portrait'
    
    # 余白設定 (数学的余白)
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = 0.47
    ws.page_margins.right = 0.47
    ws.page_margins.top = 0.41
    ws.page_margins.bottom = 0.41

    # 印刷範囲の最終決定 (全ページ分、36列制限)
    ws.print_area = f'B1:AJ{total_rows}'

    # 保存
    wb.save("output.xlsx")

# 入力データ
data_json = [
  {
    "page_number": 1,
    "width": 595.27557,
    "height": 841.88977,
    "print_range": "B3:AJ49",
    "data": [
      {"action": "merge_and_set", "start_row": 3, "start_column": 15, "end_row": 4, "end_column": 22, "value": "SAMPLE PDF", "border": False},
      {"action": "merge_and_set", "start_row": 6, "start_column": 2, "end_row": 7, "end_column": 36, "value": "ソースネクストの「いきなりPDF」は、販売本数シェアNo.1。高性能・低価格で操作も簡単。PDF作成의 常識を変えたロングセラー製品です。", "border": False},
      {"action": "merge_and_set", "start_row": 9, "start_column": 2, "end_row": 10, "end_column": 10, "value": "", "border": True},
      {"action": "merge_and_set", "start_row": 9, "start_column": 11, "end_row": 10, "end_column": 19, "value": "いきなりPDF／BASIC Edition Ver.2", "border": True},
      {"action": "merge_and_set", "start_row": 9, "start_column": 20, "end_row": 10, "end_column": 28, "value": "いきなりPDF／ STANDARD Edition Ver.2", "border": True},
      {"action": "merge_and_set", "start_row": 9, "start_column": 29, "end_row": 10, "end_column": 36, "value": "いきなりPDF／ COMPLETE Edition Ver.2", "border": True},
      {"action": "merge_and_set", "start_row": 11, "start_column": 2, "end_row": 12, "end_column": 10, "value": "ひとことで言うと", "border": True},
      {"action": "merge_and_set", "start_row": 11, "start_column": 11, "end_row": 12, "end_column": 19, "value": "PDF作成・編集", "border": True},
      {"action": "merge_and_set", "start_row": 11, "start_column": 20, "end_row": 12, "end_column": 28, "value": "PDF作成・データ変換・編集", "border": True},
      {"action": "merge_and_set", "start_row": 11, "start_column": 29, "end_row": 12, "end_column": 36, "value": "PDF作成・データ変換・高度編集", "border": True},
      {"action": "merge_and_set", "start_row": 13, "start_column": 2, "end_row": 13, "end_column": 10, "value": "標準価格（税込）", "border": True},
      {"action": "merge_and_set", "start_row": 13, "start_column": 11, "end_row": 13, "end_column": 19, "value": "2,980円", "border": True},
      {"action": "merge_and_set", "start_row": 13, "start_column": 20, "end_row": 13, "end_column": 28, "value": "3,980円", "border": True},
      {"action": "merge_and_set", "start_row": 13, "start_column": 29, "end_row": 13, "end_column": 36, "value": "9,980円", "border": True},
      {"action": "merge_and_set", "start_row": 14, "start_column": 2, "end_row": 14, "end_column": 10, "value": "Windows 8対応", "border": True},
      {"action": "merge_and_set", "start_row": 14, "start_column": 11, "end_row": 14, "end_column": 19, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 14, "start_column": 20, "end_row": 14, "end_column": 28, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 14, "start_column": 29, "end_row": 14, "end_column": 36, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 16, "start_column": 2, "end_row": 16, "end_column": 10, "value": "PDFの作成", "border": False},
      {"action": "merge_and_set", "start_row": 17, "start_column": 2, "end_row": 17, "end_column": 10, "value": "PDFファイルの作成", "border": True},
      {"action": "merge_and_set", "start_row": 17, "start_column": 11, "end_row": 17, "end_column": 19, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 17, "start_column": 20, "end_row": 17, "end_column": 28, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 17, "start_column": 29, "end_row": 17, "end_column": 36, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 18, "start_column": 2, "end_row": 18, "end_column": 10, "value": "PDFファイルの閲覧、検索", "border": True},
      {"action": "merge_and_set", "start_row": 18, "start_column": 11, "end_row": 18, "end_column": 19, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 18, "start_column": 20, "end_row": 18, "end_column": 28, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 18, "start_column": 29, "end_row": 18, "end_column": 36, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 19, "start_column": 2, "end_row": 19, "end_column": 10, "value": "Office製品へのプラグイン", "border": True},
      {"action": "merge_and_set", "start_row": 19, "start_column": 11, "end_row": 19, "end_column": 19, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 19, "start_column": 20, "end_row": 19, "end_column": 28, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 19, "start_column": 29, "end_row": 19, "end_column": 36, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 20, "start_column": 2, "end_row": 20, "end_column": 10, "value": "フォントの埋め込み", "border": True},
      {"action": "merge_and_set", "start_row": 20, "start_column": 11, "end_row": 20, "end_column": 19, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 20, "start_column": 20, "end_row": 20, "end_column": 28, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 20, "start_column": 29, "end_row": 20, "end_column": 36, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 21, "start_column": 2, "end_row": 21, "end_column": 10, "value": "複数文書の一括作成", "border": True},
      {"action": "merge_and_set", "start_row": 21, "start_column": 11, "end_row": 21, "end_column": 19, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 21, "start_column": 20, "end_row": 21, "end_column": 28, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 21, "start_column": 29, "end_row": 21, "end_column": 36, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 23, "start_column": 2, "end_row": 23, "end_column": 10, "value": "PDFの組み換え", "border": False},
      {"action": "merge_and_set", "start_row": 24, "start_column": 2, "end_row": 24, "end_column": 10, "value": "ページの分割、抽出、結合", "border": True},
      {"action": "merge_and_set", "start_row": 24, "start_column": 11, "end_row": 24, "end_column": 19, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 24, "start_column": 20, "end_row": 24, "end_column": 28, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 24, "start_column": 29, "end_row": 24, "end_column": 36, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 26, "start_column": 2, "end_row": 26, "end_column": 10, "value": "PDFの編集", "border": False},
      {"action": "merge_and_set", "start_row": 27, "start_column": 2, "end_row": 27, "end_column": 10, "value": "ノート注釈の追加", "border": True},
      {"action": "merge_and_set", "start_row": 27, "start_column": 11, "end_row": 27, "end_column": 19, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 27, "start_column": 20, "end_row": 27, "end_column": 28, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 27, "start_column": 29, "end_row": 27, "end_column": 36, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 28, "start_column": 2, "end_row": 28, "end_column": 10, "value": "テキストボックスの追加", "border": True},
      {"action": "merge_and_set", "start_row": 28, "start_column": 11, "end_row": 28, "end_column": 19, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 28, "start_column": 20, "end_row": 28, "end_column": 28, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 28, "start_column": 29, "end_row": 28, "end_column": 36, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 29, "start_column": 2, "end_row": 29, "end_column": 10, "value": "添付ファイルの追加", "border": True},
      {"action": "merge_and_set", "start_row": 29, "start_column": 11, "end_row": 29, "end_column": 19, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 29, "start_column": 20, "end_row": 29, "end_column": 28, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 29, "start_column": 29, "end_row": 29, "end_column": 36, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 30, "start_column": 2, "end_row": 30, "end_column": 10, "value": "しおりの作成、編集", "border": True},
      {"action": "merge_and_set", "start_row": 30, "start_column": 11, "end_row": 30, "end_column": 19, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 30, "start_column": 20, "end_row": 30, "end_column": 28, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 30, "start_column": 29, "end_row": 30, "end_column": 36, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 31, "start_column": 2, "end_row": 31, "end_column": 10, "value": "ハイパーリンクの挿入", "border": True},
      {"action": "merge_and_set", "start_row": 31, "start_column": 11, "end_row": 31, "end_column": 19, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 31, "start_column": 20, "end_row": 31, "end_column": 28, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 31, "start_column": 29, "end_row": 31, "end_column": 36, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 33, "start_column": 2, "end_row": 33, "end_column": 10, "value": "PDFを変換", "border": False},
      {"action": "merge_and_set", "start_row": 34, "start_column": 2, "end_row": 34, "end_column": 10, "value": "PDFをWordに変換", "border": True},
      {"action": "merge_and_set", "start_row": 34, "start_column": 11, "end_row": 34, "end_column": 19, "value": "×", "border": True},
      {"action": "merge_and_set", "start_row": 34, "start_column": 20, "end_row": 34, "end_column": 28, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 34, "start_column": 29, "end_row": 34, "end_column": 36, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 35, "start_column": 2, "end_row": 35, "end_column": 10, "value": "PDFをExcelに変換", "border": True},
      {"action": "merge_and_set", "start_row": 35, "start_column": 11, "end_row": 35, "end_column": 19, "value": "×", "border": True},
      {"action": "merge_and_set", "start_row": 35, "start_column": 20, "end_row": 35, "end_column": 28, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 35, "start_column": 29, "end_row": 35, "end_column": 36, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 36, "start_column": 2, "end_row": 36, "end_column": 10, "value": "PDFをPowePointに変換", "border": True},
      {"action": "merge_and_set", "start_row": 36, "start_column": 11, "end_row": 36, "end_column": 19, "value": "×", "border": True},
      {"action": "merge_and_set", "start_row": 36, "start_column": 20, "end_row": 36, "end_column": 28, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 36, "start_column": 29, "end_row": 36, "end_column": 36, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 37, "start_column": 2, "end_row": 37, "end_column": 10, "value": "PDFをJPEGに変換", "border": True},
      {"action": "merge_and_set", "start_row": 37, "start_column": 11, "end_row": 37, "end_column": 19, "value": "×", "border": True},
      {"action": "merge_and_set", "start_row": 37, "start_column": 20, "end_row": 37, "end_column": 28, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 37, "start_column": 29, "end_row": 37, "end_column": 36, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 38, "start_column": 2, "end_row": 38, "end_column": 10, "value": "PDFをBMPに変換", "border": True},
      {"action": "merge_and_set", "start_row": 38, "start_column": 11, "end_row": 38, "end_column": 19, "value": "×", "border": True},
      {"action": "merge_and_set", "start_row": 38, "start_column": 20, "end_row": 38, "end_column": 28, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 38, "start_column": 29, "end_row": 38, "end_column": 36, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 39, "start_column": 2, "end_row": 39, "end_column": 10, "value": "透明テキスト付きPDFに変換", "border": True},
      {"action": "merge_and_set", "start_row": 39, "start_column": 11, "end_row": 39, "end_column": 19, "value": "×", "border": True},
      {"action": "merge_and_set", "start_row": 39, "start_column": 20, "end_row": 39, "end_column": 28, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 39, "start_column": 29, "end_row": 39, "end_column": 36, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 41, "start_column": 2, "end_row": 41, "end_column": 10, "value": "PDFの直接編集", "border": False},
      {"action": "merge_and_set", "start_row": 42, "start_column": 2, "end_row": 42, "end_column": 10, "value": "すかしの挿入", "border": True},
      {"action": "merge_and_set", "start_row": 42, "start_column": 11, "end_row": 42, "end_column": 19, "value": "×", "border": True},
      {"action": "merge_and_set", "start_row": 42, "start_column": 20, "end_row": 42, "end_column": 28, "value": "×", "border": True},
      {"action": "merge_and_set", "start_row": 42, "start_column": 29, "end_row": 42, "end_column": 36, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 43, "start_column": 2, "end_row": 43, "end_column": 10, "value": "クリップアートの挿入", "border": True},
      {"action": "merge_and_set", "start_row": 43, "start_column": 11, "end_row": 43, "end_column": 19, "value": "×", "border": True},
      {"action": "merge_and_set", "start_row": 43, "start_column": 20, "end_row": 43, "end_column": 28, "value": "×", "border": True},
      {"action": "merge_and_set", "start_row": 43, "start_column": 29, "end_row": 43, "end_column": 36, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 44, "start_column": 2, "end_row": 44, "end_column": 10, "value": "スタンプの追加", "border": True},
      {"action": "merge_and_set", "start_row": 44, "start_column": 11, "end_row": 44, "end_column": 19, "value": "×", "border": True},
      {"action": "merge_and_set", "start_row": 44, "start_column": 20, "end_row": 44, "end_column": 28, "value": "×", "border": True},
      {"action": "merge_and_set", "start_row": 44, "start_column": 29, "end_row": 44, "end_column": 36, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 45, "start_column": 2, "end_row": 45, "end_column": 10, "value": "ページのトリミング編集", "border": True},
      {"action": "merge_and_set", "start_row": 45, "start_column": 11, "end_row": 45, "end_column": 19, "value": "×", "border": True},
      {"action": "merge_and_set", "start_row": 45, "start_column": 20, "end_row": 45, "end_column": 28, "value": "×", "border": True},
      {"action": "merge_and_set", "start_row": 45, "start_column": 29, "end_row": 45, "end_column": 36, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 46, "start_column": 2, "end_row": 46, "end_column": 10, "value": "フォームオブジェクトの追加", "border": True},
      {"action": "merge_and_set", "start_row": 46, "start_column": 11, "end_row": 46, "end_column": 19, "value": "×", "border": True},
      {"action": "merge_and_set", "start_row": 46, "start_column": 20, "end_row": 46, "end_column": 28, "value": "×", "border": True},
      {"action": "merge_and_set", "start_row": 46, "start_column": 29, "end_row": 46, "end_column": 36, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 47, "start_column": 2, "end_row": 47, "end_column": 10, "value": "テキストの直接編集", "border": True},
      {"action": "merge_and_set", "start_row": 47, "start_column": 11, "end_row": 47, "end_column": 19, "value": "×", "border": True},
      {"action": "merge_and_set", "start_row": 47, "start_column": 20, "end_row": 47, "end_column": 28, "value": "×", "border": True},
      {"action": "merge_and_set", "start_row": 47, "start_column": 29, "end_row": 47, "end_column": 36, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 48, "start_column": 2, "end_row": 48, "end_column": 10, "value": "オブジェクトの編集", "border": True},
      {"action": "merge_and_set", "start_row": 48, "start_column": 11, "end_row": 48, "end_column": 19, "value": "×", "border": True},
      {"action": "merge_and_set", "start_row": 48, "start_column": 20, "end_row": 48, "end_column": 28, "value": "×", "border": True},
      {"action": "merge_and_set", "start_row": 48, "start_column": 29, "end_row": 48, "end_column": 36, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 49, "start_column": 2, "end_row": 49, "end_column": 10, "value": "ページの回転編集", "border": True},
      {"action": "merge_and_set", "start_row": 49, "start_column": 11, "end_row": 49, "end_column": 19, "value": "×", "border": True},
      {"action": "merge_and_set", "start_row": 49, "start_column": 20, "end_row": 49, "end_column": 28, "value": "×", "border": True},
      {"action": "merge_and_set", "start_row": 49, "start_column": 29, "end_row": 49, "end_column": 36, "value": "○", "border": True}
    ]
  },
  {
    "page_number": 2,
    "width": 595.27557,
    "height": 841.88977,
    "print_range": "B3:AJ9",
    "data": [
      {"action": "merge_and_set", "start_row": 3, "start_column": 2, "end_row": 3, "end_column": 10, "value": "セキュリティ", "border": False},
      {"action": "merge_and_set", "start_row": 4, "start_column": 2, "end_row": 4, "end_column": 10, "value": "暗号化", "border": True},
      {"action": "merge_and_set", "start_row": 4, "start_column": 11, "end_row": 4, "end_column": 19, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 4, "start_column": 20, "end_row": 4, "end_column": 28, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 4, "start_column": 29, "end_row": 4, "end_column": 36, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 5, "start_column": 2, "end_row": 5, "end_column": 10, "value": "閲覧制限", "border": True},
      {"action": "merge_and_set", "start_row": 5, "start_column": 11, "end_row": 5, "end_column": 19, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 5, "start_column": 20, "end_row": 5, "end_column": 28, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 5, "start_column": 29, "end_row": 5, "end_column": 36, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 6, "start_column": 2, "end_row": 6, "end_column": 10, "value": "印刷制限", "border": True},
      {"action": "merge_and_set", "start_row": 6, "start_column": 11, "end_row": 6, "end_column": 19, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 6, "start_column": 20, "end_row": 6, "end_column": 28, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 6, "start_column": 29, "end_row": 6, "end_column": 36, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 7, "start_column": 2, "end_row": 7, "end_column": 10, "value": "修正制限", "border": True},
      {"action": "merge_and_set", "start_row": 7, "start_column": 11, "end_row": 7, "end_column": 19, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 7, "start_column": 20, "end_row": 7, "end_column": 28, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 7, "start_column": 29, "end_row": 7, "end_column": 36, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 8, "start_column": 2, "end_row": 8, "end_column": 10, "value": "コピー制限", "border": True},
      {"action": "merge_and_set", "start_row": 8, "start_column": 11, "end_row": 8, "end_column": 19, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 8, "start_column": 20, "end_row": 8, "end_column": 28, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 8, "start_column": 29, "end_row": 8, "end_column": 36, "value": "○", "border": True},
      {"action": "merge_and_set", "start_row": 9, "start_column": 2, "end_row": 9, "end_column": 10, "value": "電子署名の添付", "border": True},
      {"action": "merge_and_set", "start_row": 9, "start_column": 11, "end_row": 9, "end_column": 19, "value": "×", "border": True},
      {"action": "merge_and_set", "start_row": 9, "start_column": 20, "end_row": 9, "end_column": 28, "value": "×", "border": True},
      {"action": "merge_and_set", "start_row": 9, "start_column": 29, "end_row": 9, "end_column": 36, "value": "○", "border": True}
    ]
  }
]

if __name__ == "__main__":
    create_excel_from_json(data_json)