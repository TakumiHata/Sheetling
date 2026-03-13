import pdfplumber
from typing import Dict, Any, Optional


def _remove_containing_rects(rects: list) -> list:
    """
    他の矩形を完全に内包している（より大きな外枠）矩形を除去する。
    内包される側（小さい矩形＝実際のセル境界）を保持し、
    内包する側（大きい外枠・行枠）を除去することで罫線の重複描画を防ぐ。
    """
    tol = 1.0  # 座標誤差許容範囲（pt）
    to_remove = set()
    for i, a in enumerate(rects):
        if i in to_remove:
            continue
        for j, b in enumerate(rects):
            if i == j or j in to_remove:
                continue
            # a が b を完全に含み、かつ同一矩形でない
            a_contains_b = (
                a['x0'] - tol <= b['x0'] and
                a['x1'] + tol >= b['x1'] and
                a['top'] - tol <= b['top'] and
                a['bottom'] + tol >= b['bottom']
            )
            is_same = (
                abs(a['x0'] - b['x0']) < tol and
                abs(a['x1'] - b['x1']) < tol and
                abs(a['top'] - b['top']) < tol and
                abs(a['bottom'] - b['bottom']) < tol
            )
            if a_contains_b and not is_same:
                to_remove.add(i)
                break
    return [r for i, r in enumerate(rects) if i not in to_remove]


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
            # 各テーブルの列境界X座標リスト・行境界Y座標リスト（グリッド生成用）
            table_col_x_positions = []
            table_row_y_positions = []
            # 各テーブルの全セルbbox一覧（グリッド座標計算用・LLMには渡さない）
            table_cells = []
            for table in tables:
                try:
                    valid_cells = [c for c in table.cells if c is not None]
                    col_xs = sorted(set(
                        [float(c[0]) for c in valid_cells] +
                        [float(c[2]) for c in valid_cells]
                    ))
                    row_ys = sorted(set(
                        [float(c[1]) for c in valid_cells] + [float(table.bbox[3])]
                    ))
                    table_col_x_positions.append(col_xs)
                    table_row_y_positions.append(row_ys)
                    table_cells.append([
                        {'x0': float(c[0]), 'top': float(c[1]),
                         'x1': float(c[2]), 'bottom': float(c[3])}
                        for c in valid_cells
                    ])
                except Exception:
                    table_col_x_positions.append([])
                    table_row_y_positions.append([])
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

            rects = _remove_containing_rects(rects)

            page_data = {
                "page_number": page_number,
                "width": float(width),
                "height": float(height),
                "words": words,
                "table_bboxes": table_bboxes,
                "table_col_x_positions": table_col_x_positions,
                "table_row_y_positions": table_row_y_positions,
                "table_cells": table_cells,
                "table_data": cleaned_tables,
                "rects": rects
            }
            extracted_pages.append(page_data)
            
    return {"pages": extracted_pages}
