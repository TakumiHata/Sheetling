import os
import json
from pathlib import Path
from docling.document_converter import DocumentConverter
from src.utils.logger import get_logger

logger = get_logger(__name__)

class DoclingParser:
    def __init__(self, output_md_dir: str, output_json_dir: str):
        self.output_md_dir = Path(output_md_dir)
        self.output_json_dir = Path(output_json_dir)
        self.output_md_dir.mkdir(parents=True, exist_ok=True)
        self.output_json_dir.mkdir(parents=True, exist_ok=True)
        self.converter = DocumentConverter()

    def parse(self, pdf_path: str):
        """
        PDFからMarkdownを抽出する処理。
        """
        pdf_name = Path(pdf_path).stem
        logger.info(f"Parsing PDF with Docling: {pdf_path}")
        
        # Doclingによる変換実行
        result = self.converter.convert(pdf_path)
        
        # Markdownとしてエクスポート
        md_path = self.output_md_dir / f"{pdf_name}.md"
        with open(md_path, "w", encoding="utf-8") as f:
            f.write(result.document.export_to_markdown())
            
        # JSONとしてエクスポート (Doclingのドキュメント構造をそのまま出力)
        json_path = self.output_json_dir / f"{pdf_name}.json"
        with open(json_path, "w", encoding="utf-8") as f:
            # result.document.export_to_dict() を使用
            json.dump(result.document.export_to_dict(), f, indent=2, ensure_ascii=False)
            
        logger.info(f"Extracted data saved to {md_path} and {json_path}")
        return str(md_path), str(json_path)
