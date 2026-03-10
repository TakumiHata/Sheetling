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
            words = page.extract_words()
            
            # 表データの抽出 (バウンディングボックスのみ取得)
            tables = page.find_tables()
            table_bboxes = [table.bbox for table in tables]
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
            
            page_data = {
                "page_number": page_number,
                "width": float(width),
                "height": float(height),
                "words": words,
                "table_bboxes": table_bboxes,
                "table_data": cleaned_tables
            }
            extracted_pages.append(page_data)
            
    return {
        "markdown_content": markdown_content,
        "pages": extracted_pages
    }
