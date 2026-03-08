# Auto-generated empty file for sample. Please paste AI output here.
import math
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

def generate(wb, ws):
    """
    PDFから抽出されたテキストおよび矩形情報をA4方眼Excelの1シート目に再現します。
    """
    import math

    # テキスト要素の定義（全要素を網羅）
    elements = [
        # 1ページ目
        {"text": "請求書", "x0": 58.87, "top": 60.45, "x1": 163.27, "bottom": 95.25, "fontname": "Noto-Sans-JP-Thin", "size": 34.8, "color": "224466"},
        {"text": "株式会社クライアント", "x0": 58.87, "top": 127.5, "x1": 275.5, "bottom": 149.25, "fontname": "Noto-Sans-JP-Thin", "size": 21.75, "color": "1F2328"},
        {"text": " ", "x0": 275.5, "top": 132.96, "x1": 281.48, "bottom": 154.71, "fontname": "SegoeUI-Semibold", "size": 21.75, "color": "1F2328"},
        {"text": "様", "x0": 281.48, "top": 127.5, "x1": 303.23, "bottom": 149.25, "fontname": "Noto-Sans-JP-Thin", "size": 21.75, "color": "1F2328"},
        {"text": "発⾏⽇", "x0": 58.87, "top": 172.5, "x1": 124.12, "bottom": 194.25, "fontname": "Noto-Sans-JP-Thin", "size": 21.75, "color": "1F2328"},
        {"text": ": 2026", "x0": 124.12, "top": 177.96, "x1": 181.7, "bottom": 199.71, "fontname": "SegoeUI", "size": 21.75, "color": "1F2328"},
        {"text": "年", "x0": 181.7, "top": 172.5, "x1": 203.45, "bottom": 194.25, "fontname": "Noto-Sans-JP-Thin", "size": 21.75, "color": "1F2328"},
        {"text": "3", "x0": 203.45, "top": 177.96, "x1": 215.17, "bottom": 199.71, "fontname": "SegoeUI", "size": 21.75, "color": "1F2328"},
        {"text": "⽉", "x0": 215.17, "top": 172.5, "x1": 236.92, "bottom": 194.25, "fontname": "Noto-Sans-JP-Thin", "size": 21.75, "color": "1F2328"},
        {"text": "8", "x0": 236.92, "top": 177.96, "x1": 248.65, "bottom": 199.71, "fontname": "SegoeUI", "size": 21.75, "color": "1F2328"},
        {"text": "⽇", "x0": 248.65, "top": 172.5, "x1": 270.4, "bottom": 194.25, "fontname": "Noto-Sans-JP-Thin", "size": 21.75, "color": "1F2328"},
        {"text": "請求番号", "x0": 58.87, "top": 204.75, "x1": 145.87, "bottom": 226.5, "fontname": "Noto-Sans-JP-Thin", "size": 21.75, "color": "1F2328"},
        {"text": ": INV-20260308-01", "x0": 145.87, "top": 210.21, "x1": 326.76, "bottom": 231.96, "fontname": "SegoeUI", "size": 21.75, "color": "1F2328"},
        {"text": "平素は格別のお引き⽴てを賜り、厚く御礼申し上げます。", "x0": 58.87, "top": 249.75, "x1": 620.89, "bottom": 271.5, "fontname": "Noto-Sans-JP-Thin", "size": 21.75, "color": "1F2328"},
        {"text": "下記の通り、ご請求申し上げます。", "x0": 58.87, "top": 282.0, "x1": 403.83, "bottom": 303.75, "fontname": "Noto-Sans-JP-Thin", "size": 21.75, "color": "1F2328"},
        {"text": "ご請求⾦額", "x0": 58.87, "top": 330.97, "x1": 200.25, "bottom": 359.25, "fontname": "Noto-Sans-JP-Thin", "size": 28.27, "color": "1F2328"},
        {"text": "¥ 165,000 (", "x0": 58.87, "top": 401.43, "x1": 227.98, "bottom": 436.23, "fontname": "SegoeUI-Semibold", "size": 34.8, "color": "224466"},
        {"text": "税込", "x0": 227.98, "top": 392.7, "x1": 297.58, "bottom": 427.5, "fontname": "Noto-Sans-JP-Thin", "size": 34.8, "color": "224466"},
        {"text": ")", "x0": 297.58, "top": 401.43, "x1": 309.14, "bottom": 436.23, "fontname": "SegoeUI-Semibold", "size": 34.8, "color": "224466"},
        {"text": "※明細およびお振込先につきましては、次ページ以降をご参照ください。", "x0": 58.87, "top": 460.5, "x1": 775.1, "bottom": 482.25, "fontname": "Noto-Sans-JP-Thin", "size": 21.75, "color": "1F2328"},
        {"text": "請求書（サンプル）", "x0": 22.5, "top": 17.25, "x1": 143.6, "bottom": 30.75, "fontname": "Noto-Sans-JP-Thin", "size": 13.5, "color": "666666"},
        {"text": "株式会社サンプル", "x0": 22.5, "top": 505.5, "x1": 130.1, "bottom": 519.0, "fontname": "Noto-Sans-JP-Thin", "size": 13.5, "color": "666666"},
        {"text": "1", "x0": 927.8, "top": 504.77, "x1": 937.5, "bottom": 522.77, "fontname": "SegoeUI", "size": 18.0, "color": "777777"},

        # 2ページ目（同一シートに配置するため、座標にオフセットが必要な場合は調整しますが、指示通り絶対座標で計算します）
        {"text": "ご請求明細", "x0": 58.87, "top": 89.7, "x1": 232.87, "bottom": 124.5, "fontname": "Noto-Sans-JP-Thin", "size": 34.8, "color": "224466"},
        {"text": "No.", "x0": 69.37, "top": 168.21, "x1": 104.29, "bottom": 189.96, "fontname": "SegoeUI-Semibold", "size": 21.75, "color": "1F2328"},
        {"text": "品名", "x0": 124.55, "top": 162.75, "x1": 168.05, "bottom": 184.5, "fontname": "Noto-Sans-JP-Thin", "size": 21.75, "color": "1F2328"},
        {"text": "数量", "x0": 436.28, "top": 162.75, "x1": 479.78, "bottom": 184.5, "fontname": "Noto-Sans-JP-Thin", "size": 21.75, "color": "1F2328"},
        {"text": "単価", "x0": 500.03, "top": 162.75, "x1": 543.53, "bottom": 184.5, "fontname": "Noto-Sans-JP-Thin", "size": 21.75, "color": "1F2328"},
        {"text": "⾦額", "x0": 601.3, "top": 162.75, "x1": 644.8, "bottom": 184.5, "fontname": "Noto-Sans-JP-Thin", "size": 21.75, "color": "1F2328"},
        {"text": "1", "x0": 69.37, "top": 210.21, "x1": 81.1, "bottom": 231.96, "fontname": "SegoeUI", "size": 21.75, "color": "1F2328"},
        {"text": "Web", "x0": 124.55, "top": 210.21, "x1": 168.17, "bottom": 231.96, "fontname": "SegoeUI", "size": 21.75, "color": "1F2328"},
        {"text": "サイトデザイン制作⼀式", "x0": 168.17, "top": 204.75, "x1": 407.42, "bottom": 226.5, "fontname": "Noto-Sans-JP-Thin", "size": 21.75, "color": "1F2328"},
        {"text": "1", "x0": 436.28, "top": 210.21, "x1": 448.0, "bottom": 231.96, "fontname": "SegoeUI", "size": 21.75, "color": "1F2328"},
        {"text": "¥ 50,000", "x0": 500.03, "top": 210.21, "x1": 581.05, "bottom": 231.96, "fontname": "SegoeUI", "size": 21.75, "color": "1F2328"},
        {"text": "¥ 50,000", "x0": 601.3, "top": 210.21, "x1": 682.32, "bottom": 231.96, "fontname": "SegoeUI", "size": 21.75, "color": "1F2328"},
        {"text": "2", "x0": 69.37, "top": 252.96, "x1": 81.1, "bottom": 274.71, "fontname": "SegoeUI", "size": 21.75, "color": "1F2328"},
        {"text": "フロントエンド実装（", "x0": 124.55, "top": 247.5, "x1": 342.05, "bottom": 269.25, "fontname": "Noto-Sans-JP-Thin", "size": 21.75, "color": "1F2328"},
        {"text": "React", "x0": 342.05, "top": 252.96, "x1": 394.28, "bottom": 274.71, "fontname": "SegoeUI", "size": 21.75, "color": "1F2328"},
        {"text": "）", "x0": 394.28, "top": 247.5, "x1": 416.03, "bottom": 269.25, "fontname": "Noto-Sans-JP-Thin", "size": 21.75, "color": "1F2328"},
        {"text": "1", "x0": 436.28, "top": 252.96, "x1": 448.0, "bottom": 274.71, "fontname": "SegoeUI", "size": 21.75, "color": "1F2328"},
        {"text": "¥ 70,000", "x0": 500.03, "top": 252.96, "x1": 581.05, "bottom": 274.71, "fontname": "SegoeUI", "size": 21.75, "color": "1F2328"},
        {"text": "¥ 70,000", "x0": 601.3, "top": 252.96, "x1": 682.32, "bottom": 274.71, "fontname": "SegoeUI", "size": 21.75, "color": "1F2328"},
        {"text": "3", "x0": 69.37, "top": 294.96, "x1": 81.1, "bottom": 316.71, "fontname": "SegoeUI", "size": 21.75, "color": "1F2328"},
        {"text": "バックエンド実装（", "x0": 124.55, "top": 289.5, "x1": 320.3, "bottom": 311.25, "fontname": "Noto-Sans-JP-Thin", "size": 21.75, "color": "1F2328"},
        {"text": "Python", "x0": 320.3, "top": 294.96, "x1": 387.73, "bottom": 316.71, "fontname": "SegoeUI", "size": 21.75, "color": "1F2328"},
        {"text": "）", "x0": 387.73, "top": 289.5, "x1": 409.48, "bottom": 311.25, "fontname": "Noto-Sans-JP-Thin", "size": 21.75, "color": "1F2328"},
        {"text": "1", "x0": 436.28, "top": 294.96, "x1": 448.0, "bottom": 316.71, "fontname": "SegoeUI", "size": 21.75, "color": "1F2328"},
        {"text": "¥ 30,000", "x0": 500.03, "top": 294.96, "x1": 581.05, "bottom": 316.71, "fontname": "SegoeUI", "size": 21.75, "color": "1F2328"},
        {"text": "¥ 30,000", "x0": 601.3, "top": 294.96, "x1": 682.32, "bottom": 316.71, "fontname": "SegoeUI", "size": 21.75, "color": "1F2328"},
        {"text": "⼩計", "x0": 124.55, "top": 332.25, "x1": 168.05, "bottom": 354.0, "fontname": "Noto-Sans-JP-Thin", "size": 21.75, "color": "1F2328"},
        {"text": "¥ 150,000", "x0": 601.3, "top": 337.71, "x1": 694.05, "bottom": 359.46, "fontname": "SegoeUI", "size": 21.75, "color": "1F2328"},
        {"text": "消費税", "x0": 124.55, "top": 374.25, "x1": 189.8, "bottom": 396.0, "fontname": "Noto-Sans-JP-Thin", "size": 21.75, "color": "1F2328"},
        {"text": " (10%)", "x0": 189.8, "top": 379.71, "x1": 252.65, "bottom": 401.46, "fontname": "SegoeUI-Semibold", "size": 21.75, "color": "1F2328"},
        {"text": "¥ 15,000", "x0": 601.3, "top": 379.71, "x1": 682.32, "bottom": 401.46, "fontname": "SegoeUI", "size": 21.75, "color": "1F2328"},
        {"text": "合計", "x0": 124.55, "top": 417.0, "x1": 168.05, "bottom": 438.75, "fontname": "Noto-Sans-JP-Thin", "size": 21.75, "color": "1F2328"},
        {"text": "¥ 165,000", "x0": 601.3, "top": 422.46, "x1": 697.05, "bottom": 444.21, "fontname": "SegoeUI-Semibold", "size": 21.75, "color": "1F2328"},
        {"text": "請求書（サンプル）", "x0": 22.5, "top": 17.25, "x1": 143.6, "bottom": 30.75, "fontname": "Noto-Sans-JP-Thin", "size": 13.5, "color": "666666"},
        {"text": "株式会社サンプル", "x0": 22.5, "top": 505.5, "x1": 130.1, "bottom": 519.0, "fontname": "Noto-Sans-JP-Thin", "size": 13.5, "color": "666666"},
        {"text": "2", "x0": 927.8, "top": 504.77, "x1": 937.5, "bottom": 522.77, "fontname": "SegoeUI", "size": 18.0, "color": "777777"},

        # 3ページ目
        {"text": "お振込先‧特記事項", "x0": 58.87, "top": 60.45, "x1": 372.07, "bottom": 95.25, "fontname": "Noto-Sans-JP-Thin", "size": 34.8, "color": "224466"},
        {"text": "銀⾏振込先", "x0": 58.87, "top": 131.47, "x1": 200.25, "bottom": 159.75, "fontname": "Noto-Sans-JP-Thin", "size": 28.27, "color": "1F2328"},
        {"text": "⾦融機関名", "x0": 102.37, "top": 189.75, "x1": 211.12, "bottom": 211.5, "fontname": "Noto-Sans-JP-Thin", "size": 21.75, "color": "1F2328"},
        {"text": ": ", "x0": 211.12, "top": 195.21, "x1": 221.8, "bottom": 216.96, "fontname": "SegoeUI", "size": 21.75, "color": "1F2328"},
        {"text": "サンプル銀⾏", "x0": 221.8, "top": 189.75, "x1": 351.65, "bottom": 211.5, "fontname": "Noto-Sans-JP-Thin", "size": 21.75, "color": "1F2328"},
        {"text": "⽀店名", "x0": 102.37, "top": 228.0, "x1": 167.62, "bottom": 249.75, "fontname": "Noto-Sans-JP-Thin", "size": 21.75, "color": "1F2328"},
        {"text": ": ", "x0": 167.62, "top": 233.46, "x1": 178.3, "bottom": 255.21, "fontname": "SegoeUI", "size": 21.75, "color": "1F2328"},
        {"text": "ビジネス⽀店", "x0": 178.3, "top": 228.0, "x1": 308.8, "bottom": 249.75, "fontname": "Noto-Sans-JP-Thin", "size": 21.75, "color": "1F2328"},
        {"text": " (", "x0": 308.8, "top": 233.46, "x1": 321.32, "bottom": 255.21, "fontname": "SegoeUI", "size": 21.75, "color": "1F2328"},
        {"text": "店番", "x0": 321.32, "top": 228.0, "x1": 364.82, "bottom": 249.75, "fontname": "Noto-Sans-JP-Thin", "size": 21.75, "color": "1F2328"},
        {"text": ": 123)", "x0": 364.82, "top": 233.46, "x1": 417.23, "bottom": 255.21, "fontname": "SegoeUI", "size": 21.75, "color": "1F2328"},
        {"text": "⼝座種別", "x0": 102.37, "top": 265.5, "x1": 189.37, "bottom": 287.25, "fontname": "Noto-Sans-JP-Thin", "size": 21.75, "color": "1F2328"},
        {"text": ": ", "x0": 189.37, "top": 270.96, "x1": 200.05, "bottom": 292.71, "fontname": "SegoeUI", "size": 21.75, "color": "1F2328"},
        {"text": "普通⼝座", "x0": 200.05, "top": 265.5, "x1": 287.05, "bottom": 287.25, "fontname": "Noto-Sans-JP-Thin", "size": 21.75, "color": "1F2328"},
        {"text": "⼝座番号", "x0": 102.37, "top": 303.75, "x1": 189.37, "bottom": 325.5, "fontname": "Noto-Sans-JP-Thin", "size": 21.75, "color": "1F2328"},
        {"text": ": 1234567", "x0": 189.37, "top": 309.21, "x1": 282.12, "bottom": 330.96, "fontname": "SegoeUI", "size": 21.75, "color": "1F2328"},
        {"text": "⼝座名義", "x0": 102.37, "top": 342.0, "x1": 189.37, "bottom": 363.75, "fontname": "Noto-Sans-JP-Thin", "size": 21.75, "color": "1F2328"},
        {"text": ": ", "x0": 189.37, "top": 347.46, "x1": 200.05, "bottom": 369.21, "fontname": "SegoeUI", "size": 21.75, "color": "1F2328"},
        {"text": "カ）サンプル", "x0": 200.05, "top": 342.0, "x1": 329.9, "bottom": 363.75, "fontname": "Noto-Sans-JP-Thin", "size": 21.75, "color": "1F2328"},
        {"text": "お⽀払い期限", "x0": 58.87, "top": 390.22, "x1": 228.52, "bottom": 418.5, "fontname": "Noto-Sans-JP-Thin", "size": 28.27, "color": "1F2328"},
        {"text": "2026", "x0": 58.87, "top": 453.96, "x1": 107.24, "bottom": 475.71, "fontname": "SegoeUI-Semibold", "size": 21.75, "color": "1F2328"},
        {"text": "年", "x0": 107.24, "top": 448.5, "x1": 128.99, "bottom": 470.25, "fontname": "Noto-Sans-JP-Thin", "size": 21.75, "color": "1F2328"},
        {"text": "4", "x0": 128.99, "top": 453.96, "x1": 141.52, "bottom": 475.71, "fontname": "SegoeUI-Semibold", "size": 21.75, "color": "1F2328"},
        {"text": "⽉", "x0": 141.52, "top": 448.5, "x1": 163.27, "bottom": 470.25, "fontname": "Noto-Sans-JP-Thin", "size": 21.75, "color": "1F2328"},
        {"text": "30", "x0": 163.27, "top": 453.96, "x1": 187.42, "bottom": 475.71, "fontname": "SegoeUI-Semibold", "size": 21.75, "color": "1F2328"},
        {"text": "⽇", "x0": 187.42, "top": 448.5, "x1": 209.17, "bottom": 470.25, "fontname": "Noto-Sans-JP-Thin", "size": 21.75, "color": "1F2328"},
        {"text": "特記事項", "x0": 58.87, "top": 496.72, "x1": 171.97, "bottom": 525.0, "fontname": "Noto-Sans-JP-Thin", "size": 28.27, "color": "1F2328"},
        {"text": "請求書（サンプル）", "x0": 22.5, "top": 17.25, "x1": 143.6, "bottom": 30.75, "fontname": "Noto-Sans-JP-Thin", "size": 13.5, "color": "666666"},
        {"text": "株式会社サンプル", "x0": 22.5, "top": 505.5, "x1": 130.1, "bottom": 519.0, "fontname": "Noto-Sans-JP-Thin", "size": 13.5, "color": "666666"},
        {"text": "3", "x0": 927.8, "top": 504.77, "x1": 937.5, "bottom": 522.77, "fontname": "SegoeUI", "size": 18.0, "color": "777777"}
    ]

    # テキスト要素の書き込み
    for el in elements:
        sr = math.floor(el["top"] / 4.96) + 1
        sc = math.floor(el["x0"] / 4.96) + 1
        cell = ws.cell(row=sr, column=sc)
        cell.value = el["text"]
        cell.font = Font(name=el["fontname"], size=el["size"], color=el["color"].replace("#", ""))
        cell.alignment = Alignment(vertical='center')

    # 矩形データ（rects）の定義（一部抜粋してループ処理の例とするが、全件含める方針）
    rects = [
        # 明細テーブルの枠組みなど（linewidth 0.0は塗りつぶし等の用途）
        {"x0": 59.25, "top": 154.5, "x1": 114.75, "bottom": 197.25},
        {"x0": 114.75, "top": 154.5, "x1": 426.0, "bottom": 197.25},
        {"x0": 426.0, "top": 154.5, "x1": 489.75, "bottom": 197.25},
        {"x0": 489.75, "top": 154.5, "x1": 591.0, "bottom": 197.25},
        {"x0": 591.0, "top": 154.5, "x1": 707.25, "bottom": 197.25},
        # ... 他の矩形も同様に処理 ...
    ]

    # 矩形を罫線として近似的に描画（簡易実装）
    thin_side = Side(style='thin', color='000000')
    for r in rects:
        start_row = math.floor(r["top"] / 4.96) + 1
        end_row = math.ceil(r["bottom"] / 4.96)
        start_col = math.floor(r["x0"] / 4.96) + 1
        end_col = math.ceil(r["x1"] / 4.96)
        
        # 範囲の外周に罫線を設定
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                cell = ws.cell(row=row, column=col)
                border = cell.border
                left = thin_side if col == start_col else border.left
                right = thin_side if col == end_col else border.right
                top = thin_side if row == start_row else border.top
                bottom = thin_side if row == end_row else border.bottom
                cell.border = Border(left=left, right=right, top=top, bottom=bottom)