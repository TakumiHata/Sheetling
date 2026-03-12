import pdfplumber
from typing import Dict, Any, Optional


def _to_hex_color(color) -> Optional[str]:
    """pdfplumber のカラー値を Excel 用 RRGGBB 16進文字列に変換する。None の場合は None を返す。"""
    if color is None:
        return None
    if isinstance(color, (int, float)):
        # グレースケール (0.0=黒 〜 1.0=白)
        v = int(round(float(color) * 255))
        return f"{v:02X}{v:02X}{v:02X}"
    if isinstance(color, (list, tuple)):
        if len(color) == 3:
            # RGB
            r, g, b = [int(round(c * 255)) for c in color]
            return f"{r:02X}{g:02X}{b:02X}"
        if len(color) == 4:
            # CMYK
            c, m, y, k = color
            r = int(round((1 - c) * (1 - k) * 255))
            g = int(round((1 - m) * (1 - k) * 255))
            b = int(round((1 - y) * (1 - k) * 255))
            return f"{r:02X}{g:02X}{b:02X}"
    return None


def extract_pdf_data(pdf_path: str) -> Dict[str, Any]:
    """
    """
    extracted_pages = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            # テキストの抽出 (x0, top, x1, bottom, text + フォント情報・色情報)
            all_words = page.extract_words(
                extra_attrs=['fontname', 'size', 'non_stroking_color']
            )
            
            # 表データの抽出
            tables = page.find_tables()
            table_bboxes = [table.bbox for table in tables]
            # 各テーブルの列左端X座標リスト（列アンカー計算用）
            table_col_x_positions = []
            # 各テーブルの全セルbbox一覧（セル単位の枠線描画用）
            table_cells = []
            for table in tables:
                try:
                    valid_cells = [c for c in table.cells if c is not None]
                    col_xs = sorted(set(float(c[0]) for c in valid_cells))
                    table_col_x_positions.append(col_xs)
                    table_cells.append([
                        {'x0': float(c[0]), 'top': float(c[1]),
                         'x1': float(c[2]), 'bottom': float(c[3])}
                        for c in valid_cells
                    ])
                except Exception:
                    table_col_x_positions.append([])
                    table_cells.append([])
            # font_color フィールドを hex 文字列に変換（None の場合は省略）
            words = []
            for w in all_words:
                word = dict(w)
                raw_color = word.pop('non_stroking_color', None)
                hex_color = _to_hex_color(raw_color)
                if hex_color is not None:
                    word['font_color'] = hex_color
                # font_size は小数点以下1桁に丸める
                raw_size = word.pop('size', None)
                if raw_size is not None:
                    word['font_size'] = round(float(raw_size), 1)
                words.append(word)

            # 表の内部構造（2次元配列）の取得
            table_data = page.extract_tables()
            # 扱いやすくするため、改行文字等が含まれていたら除去
            cleaned_tables = []
            for table in table_data:
                cleaned_table = []
                for row in table:
                    cleaned_row = [cell.replace('\n', ' ') if isinstance(cell, str) else cell for cell in row]
                    cleaned_table.append(cleaned_row)
                cleaned_tables.append(cleaned_table)
            
            # ページサイズの取得
            width = page.width
            height = page.height
            page_area = float(width) * float(height)

            # 矩形枠の抽出（フォームフィールド・罫線ボックス等）
            # ページ全体を覆う矩形（ページ境界・背景）は除外する
            rects = []
            for r in page.rects:
                rect_area = (r['x1'] - r['x0']) * (r['bottom'] - r['top'])
                if rect_area < 0.85 * page_area:
                    rect_entry = {
                        'x0': float(r['x0']),
                        'top': float(r['top']),
                        'x1': float(r['x1']),
                        'bottom': float(r['bottom'])
                    }
                    fill_hex = _to_hex_color(r.get('non_stroking_color'))
                    if fill_hex is not None:
                        rect_entry['fill_color'] = fill_hex
                    rects.append(rect_entry)

            page_data = {
                "page_number": page_number,
                "width": float(width),
                "height": float(height),
                "words": words,
                "table_bboxes": table_bboxes,
                "table_col_x_positions": table_col_x_positions,
                "table_cells": table_cells,
                "table_data": cleaned_tables,
                "rects": rects
            }
            extracted_pages.append(page_data)
            
    return {"pages": extracted_pages}
