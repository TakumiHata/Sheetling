import re
from pathlib import Path
from markitdown import MarkItDown
from src.utils.logger import get_logger

logger = get_logger(__name__)


class MarkItDownParser:
    def __init__(self, output_dir: str):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.md = MarkItDown()

    def parse(self, pdf_path: str) -> str:
        """
        MarkItDownを使用してPDFからテキスト情報を高精度に抽出する。
        NaN/Unnamed等のノイズを除去した「クリーンなMD」を出力する。
        """
        logger.info(f"Parsing PDF with MarkItDown: {pdf_path}")
        pdf_name = Path(pdf_path).stem

        try:
            result = self.md.convert(str(pdf_path))
            cleaned = self._clean_markdown(result.text_content)

            output_path = self.output_dir / f"{pdf_name}.md"
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(cleaned)

            logger.info(f"MarkItDown extraction saved to: {output_path}")
            return str(output_path)
        except Exception as e:
            logger.error(f"MarkItDown parsing failed: {e}")
            raise

    def _clean_markdown(self, text: str) -> str:
        """
        MarkItDown出力からノイズを除去する。
        - NaN, Unnamed 系のプレースホルダーを除去
        - 過剰な空行を整理
        """
        # NaN / Unnamed / nan を除去（テーブルセル内）
        text = re.sub(r'\bNaN\b', '', text)
        text = re.sub(r'\bnan\b', '', text)
        text = re.sub(r'Unnamed:\s*\d+', '', text)
        text = re.sub(r'Unnamed', '', text)

        # テーブル行の中身が全て空（| | | |）になった行を除去
        text = re.sub(r'^\|[\s|]*\|$', '', text, flags=re.MULTILINE)

        # 3行以上の連続空行を2行に圧縮
        text = re.sub(r'\n{3,}', '\n\n', text)

        return text.strip() + '\n'
