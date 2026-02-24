import json
from pathlib import Path
import pdfplumber
from src.utils.logger import get_logger

logger = get_logger(__name__)

class PdfLayoutExtractor:
    def __init__(self, output_dir: str):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)

    def extract(self, pdf_path: str):
        """
        pdfplumberを使用してPDFから詳細なレイアウト情報を抽出する。
        """
        logger.info(f"Extracting detailed layout from: {pdf_path}")
        pdf_path_obj = Path(pdf_path)
        pdf_name = pdf_path_obj.stem
        
        layout_data = {
            "source": str(pdf_path),
            "pages": []
        }

        with pdfplumber.open(pdf_path) as pdf:
            for i, page in enumerate(pdf.pages):
                logger.info(f"Processing page {i+1}/{len(pdf.pages)}")
                
                page_data = {
                    "page_number": i + 1,
                    "width": float(page.width),
                    "height": float(page.height),
                    "chars": [],
                    "lines": [],
                    "rects": [],
                    "words": []
                }

                # 単語情報の抽出 (座標維持のため重要)
                for word in page.extract_words():
                    page_data["words"].append({
                        "text": word["text"],
                        "x0": float(word["x0"]),
                        "top": float(word["top"]),
                        "x1": float(word["x1"]),
                        "bottom": float(word["bottom"]),
                    })

                # 文字情報の抽出
                for char in page.chars:
                    page_data["chars"].append({
                        "text": char["text"],
                        "x0": float(char["x0"]),
                        "top": float(char["top"]),
                        "x1": float(char["x1"]),
                        "bottom": float(char["bottom"]),
                        "size": float(char["size"]),
                        "fontname": char["fontname"],
                        "stroking_color": char.get("stroking_color"),
                        "non_stroking_color": char.get("non_stroking_color")
                    })

                # 直線情報の抽出
                for line in page.lines:
                    page_data["lines"].append({
                        "x0": float(line["x0"]),
                        "top": float(line["top"]),
                        "x1": float(line["x1"]),
                        "bottom": float(line["bottom"]),
                        "width": float(line["width"]),
                        "stroking_color": line.get("stroking_color"),
                        "non_stroking_color": line.get("non_stroking_color")
                    })

                # 長方形情報の抽出
                for rect in page.rects:
                    page_data["rects"].append({
                        "x0": float(rect["x0"]),
                        "top": float(rect["top"]),
                        "x1": float(rect["x1"]),
                        "bottom": float(rect["bottom"]),
                        "width": float(rect["width"]),
                        "stroking_color": rect.get("stroking_color"),
                        "non_stroking_color": rect.get("non_stroking_color")
                    })

                layout_data["pages"].append(page_data)

        output_path = self.output_dir / f"{pdf_name}_layout.json"
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(layout_data, f, indent=2, ensure_ascii=False)

        logger.info(f"Detailed layout saved to: {output_path}")
        return str(output_path)
