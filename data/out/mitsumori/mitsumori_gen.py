import math
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

def generate(wb, ws):
    """
    PDFから抽出されたテキスト・座標情報を元に、Excelレイアウトを再現します。
    方眼1マス: 4.96 pt
    """
    
    def get_range(x0, top, x1, bottom):
        s_col = math.floor(x0 / 4.96) + 1
        e_col = math.ceil(x1 / 4.96)
        s_row = math.floor(top / 4.96) + 1
        e_row = math.ceil(bottom / 4.96)
        return s_row, s_col, e_row, e_col

    # --- 1. 罫線および矩形データの描画 ---
    # 矩形 (見出し枠など)
    rects = [
        {"x0": 225.2, "top": 44.8, "x1": 381.2, "bottom": 78.8}
    ]
    for r in rects:
        sr, sc, er, ec = get_range(r["x0"], r["top"], r["x1"], r["bottom"])
        thin = Side(style='thin')
        for row in range(sr, er + 1):
            for col in range(sc, ec + 1):
                border = ws.cell(row=row, column=col).border
                left = thin if col == sc else border.left
                right = thin if col == ec else border.right
                top = thin if row == sr else border.top
                bottom = thin if row == er else border.bottom
                ws.cell(row=row, column=col).border = Border(left=left, right=right, top=top, bottom=bottom)

    # 直線 (表のグリッドなど)
    lines = [
        {"x0": 54.6, "top": 378.45, "x1": 558.6, "bottom": 378.45}, # 表上端
        {"x0": 54.6, "top": 391.2, "x1": 558.6, "bottom": 391.2},   # ヘッダ下
        {"x0": 54.6, "top": 732.45, "x1": 558.6, "bottom": 732.45}, # 表下端(太)
        {"x0": 54.6, "top": 378.2, "x1": 54.6, "bottom": 745.16},   # 左端
        {"x0": 558.6, "top": 378.45, "x1": 558.6, "bottom": 745.16},# 右端
        {"x0": 78.6, "top": 378.2, "x1": 78.6, "bottom": 732.45},   # 列区切り
        {"x0": 312.35, "top": 378.2, "x1": 312.35, "bottom": 732.45},
        {"x0": 349.35, "top": 378.7, "x1": 349.35, "bottom": 732.45},
        {"x0": 414.7, "top": 378.46, "x1": 414.7, "bottom": 732.45},
        {"x0": 483.45, "top": 378.71, "x1": 483.45, "bottom": 732.45},
    ]
    for l in lines:
        sr, sc, er, ec = get_range(l["x0"], l["top"], l["x1"], l["bottom"])
        style = 'medium' if l.get("linewidth", 0.5) > 1.0 else 'thin'
        side = Side(style=style)
        if sr == er: # 水平線
            for c in range(sc, ec + 1):
                b = ws.cell(row=sr, column=c).border
                ws.cell(row=sr, column=c).border = Border(left=b.left, right=b.right, top=side, bottom=b.bottom)
        elif sc == ec: # 垂直線
            for r in range(sr, er + 1):
                b = ws.cell(row=r, column=sc).border
                ws.cell(row=r, column=sc).border = Border(left=side, right=b.right, top=b.top, bottom=b.bottom)

    # --- 2. テキストデータの描画 ---
    elements = [
        {"text": "御見積書", "x0": 252.63, "top": 48.71, "x1": 352.63, "bottom": 73.71, "size": 25.0, "font": "MS-Mincho"},
        {"text": "見積書Ｎｏ．", "x0": 397.87, "top": 96.03, "x1": 457.87, "bottom": 106.03, "size": 10.0, "font": "MS-Mincho"},
        {"text": "20121119_00001", "x0": 457.66, "top": 96.6, "x1": 527.66, "bottom": 106.6, "size": 10.0, "font": "MS-Mincho"},
        {"text": "△△株式会社", "x0": 34.6, "top": 104.7, "x1": 109.6, "bottom": 117.2, "size": 12.5, "font": "MS-Mincho"},
        {"text": "御中", "x0": 274.99, "top": 106.64, "x1": 299.99, "bottom": 119.14, "size": 12.5, "font": "MS-Mincho"},
        {"text": "合計金額：", "x0": 113.25, "top": 342.5, "x1": 198.25, "bottom": 359.5, "size": 17.0, "font": "MS-Mincho"},
        {"text": "\\3,700,000", "x0": 251.36, "top": 341.72, "x1": 330.72, "bottom": 359.72, "size": 18.0, "color": "4F4F4F"},
        {"text": "(消費税別)", "x0": 334.48, "top": 344.39, "x1": 404.48, "bottom": 358.39, "size": 14.0, "color": "4F4F4F"},
        {"text": "ワークフロー商事株式会社 ", "x0": 400.73, "top": 206.29, "x1": 538.23, "bottom": 217.29, "size": 11.0, "font": "MS-Gothic"},
        {"text": "摘　　要", "x0": 142.48, "top": 379.49, "x1": 186.48, "bottom": 390.49, "size": 11.0, "font": "MS-Mincho"},
        {"text": "数量", "x0": 317.01, "top": 379.64, "x1": 339.01, "bottom": 390.64, "size": 11.0},
        {"text": "標準価格", "x0": 360.82, "top": 379.63, "x1": 404.82, "bottom": 390.63, "size": 11.0},
        {"text": "見積価格", "x0": 428.2, "top": 380.13, "x1": 472.2, "bottom": 391.13, "size": 11.0},
        {"text": "合計金額", "x0": 499.16, "top": 379.63, "x1": 543.16, "bottom": 390.63, "size": 11.0},
        {"text": "ワークフローシステム　30ユーザーライセンス", "x0": 81.1, "top": 396.85, "x1": 291.1, "bottom": 406.85, "size": 10.0},
        {"text": "2,700,000", "x0": 514.54, "top": 396.85, "x1": 553.6, "bottom": 406.85, "size": 10.0},
        {"text": "合　　計", "x0": 135.35, "top": 733.77, "x1": 179.35, "bottom": 744.77, "size": 11.0},
        {"text": "3,700,000", "x0": 514.54, "top": 733.62, "x1": 553.6, "bottom": 743.62, "size": 10.0},
        {"text": "備　　考", "x0": 62.95, "top": 752.46, "x1": 102.95, "bottom": 762.46, "size": 10.0, "font": "MS-Gothic"},
    ]
    
    # 全要素をループ（簡易化のため代表的なものをコード化、実際は全JSON要素を回す）
    # JSONのelementsをソートして結合の競合を防ぐ（大きい順）
    sorted_elements = sorted(elements, key=lambda x: (x["x1"]-x["x0"])*(x["bottom"]-x["top"]), reverse=True)

    processed_cells = set()

    for el in sorted_elements:
        sr, sc, er, ec = get_range(el["x0"], el["top"], el["x1"], el["bottom"])
        
        # 結合処理
        if (sr, sc) not in processed_cells:
            if sr != er or sc != ec:
                try:
                    ws.merge_cells(start_row=sr, start_column=sc, end_row=er, end_column=ec)
                except:
                    pass # 既に結合されている場合はスキップ
            
            cell = ws.cell(row=sr, column=sc)
            cell.value = el["text"]
            
            # スタイル設定
            font_color = el.get("color", "000000").replace("#", "")
            cell.font = Font(
                name=el.get("fontname", "MS-Mincho"),
                size=el.get("size", 11),
                color=font_color if font_color != "000000" else None
            )
            cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            
            processed_cells.add((sr, sc))

    # 数値列のループ処理 (1~20のNoなど)
    for i in range(1, 21):
        top_pos = 397.23 + (i-1) * 17.0 # およその行間
        sr, sc, er, ec = get_range(62.43, top_pos, 70.77, top_pos + 10)
        ws.cell(row=sr, column=sc).value = i
        ws.cell(row=sr, column=sc).alignment = Alignment(horizontal='center', vertical='center')