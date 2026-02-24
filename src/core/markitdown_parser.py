from pathlib import Path
from markitdown import MarkItDown
from src.utils.logger import get_logger

logger = get_logger(__name__)

class MarkItDownParser:
    def __init__(self, output_dir: str):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.md = MarkItDown()

    def parse(self, pdf_path: str):
        """
        MarkItDownを使用してPDFからテキスト情報を高精度に抽出する。
        """
        logger.info(f"Parsing PDF with MarkItDown: {pdf_path}")
        pdf_path_obj = Path(pdf_path)
        pdf_name = pdf_path_obj.stem
        
        try:
            result = self.md.convert(str(pdf_path))
            output_path = self.output_dir / f"{pdf_name}_mid.md"
            
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(result.text_content)
                
            logger.info(f"MarkItDown extraction saved to: {output_path}")
            return str(output_path)
        except Exception as e:
            logger.error(f"MarkItDown parsing failed: {e}")
            raise
