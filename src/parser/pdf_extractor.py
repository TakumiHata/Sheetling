import pdfplumber
from typing import Dict, Any, List
from markitdown import MarkItDown

def extract_pdf_data(pdf_path: str) -> Dict[str, Any]:
    """
    指定されたPDFファイルから、次の2つを抽出して返却します。
    1. markitdownによるPDF全体のMarkdownテキスト
    2. pdfplumberによる各ページのテキスト情報、表のバウンディングボックス等の詳細データ
    
    Args:
        pdf_path (str): 読み込むPDFファイルのパス
        
    Returns:
        Dict[str, Any]: 
            - markdown_content: (str) markitdownで抽出したMarkdownテキスト
            - pages: (List[Dict]) pdfplumberで抽出した各ページの詳細データ
    """
    # 1. markitdown による Markdown テキストの抽出
    md = MarkItDown()
    result = md.convert(pdf_path)
    markdown_content = result.text_content

    # 2. pdfplumber による詳細データの抽出
    extracted_pages = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            # テキストの抽出 (x0, top, x1, bottom, text などが含まれる)
            all_words = page.extract_words()
            
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
            words = all_words

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
                    rects.append({
                        'x0': float(r['x0']),
                        'top': float(r['top']),
                        'x1': float(r['x1']),
                        'bottom': float(r['bottom'])
                    })

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
            
    return {
        "markdown_content": markdown_content,
        "pages": extracted_pages
    }
