import os
import json
from pathlib import Path
from src.utils.logger import get_logger

logger = get_logger(__name__)

class DoclingParser:
    def __init__(self, output_md_dir: str, output_json_dir: str):
        self.output_md_dir = Path(output_md_dir)
        self.output_json_dir = Path(output_json_dir)
        self.output_md_dir.mkdir(parents=True, exist_ok=True)
        self.output_json_dir.mkdir(parents=True, exist_ok=True)

    def parse(self, pdf_path: str):
        """
        PDFからMarkdownとJSONを抽出するモック処理。
        実際には Docling 外部ライブラリを呼び出す。
        """
        pdf_name = Path(pdf_path).stem
        logger.info(f"Parsing PDF: {pdf_path}")
        
        # モック出力ファイルのパス
        md_path = self.output_md_dir / f"{pdf_name}.md"
        json_path = self.output_json_dir / f"{pdf_name}.json"
        
        # モック内容の書き込み
        md_content = f"# Parsed Content of {pdf_name}\n\nThis is a mock markdown output."
        json_content = {
            "source": pdf_path,
            "metadata": {"title": pdf_name, "author": "Mock"},
            "elements": []
        }
        
        with open(md_path, "w", encoding="utf-8") as f:
            f.write(md_content)
        
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(json_content, f, indent=2, ensure_ascii=False)
            
        logger.info(f"Extracted data saved to {md_path} and {json_path}")
        return str(md_path), str(json_path)
