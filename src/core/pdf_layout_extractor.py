import json
from pathlib import Path
import pdfplumber
from src.utils.logger import get_logger

logger = get_logger(__name__)


class PdfLayoutExtractor:
    def __init__(self, output_dir: str):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)

    def extract(self, pdf_path: str) -> dict:
        """
        pdfplumberを使用してPDFから詳細なレイアウト情報を抽出する。
        色情報はHEX形式に正規化して辞書として返す。
        """
        logger.info(f"Extracting detailed layout from: {pdf_path}")

        layout_data = {
            "source": str(pdf_path),
            "pages": []
        }

        with pdfplumber.open(pdf_path) as pdf:
            current_y_offset = 0.0

            for i, page in enumerate(pdf.pages):
                logger.info(f"Processing page {i+1}/{len(pdf.pages)}")

                page_data = {
                    "page_number": i + 1,
                    "width": float(page.width),
                    "height": float(page.height),
                    "words": [],
                    "rects": [],
                    "lines": [],
                    "chars": []
                }

                # 単語情報の抽出（座標維持のため重要）
                for word in page.extract_words():
                    page_data["words"].append({
                        "text": word["text"],
                        "x0": float(word["x0"]),
                        "top": float(word["top"]) + current_y_offset,
                        "x1": float(word["x1"]),
                        "bottom": float(word["bottom"]) + current_y_offset,
                    })

                # 長方形情報の抽出（背景色・ボーダー）
                for rect in page.rects:
                    page_data["rects"].append({
                        "x0": float(rect["x0"]),
                        "top": float(rect["top"]) + current_y_offset,
                        "x1": float(rect["x1"]),
                        "bottom": float(rect["bottom"]) + current_y_offset,
                        "stroke_width": float(rect.get("width", 0) or 0),
                    })

                # 直線情報の抽出（罫線）
                for line in page.lines:
                    page_data["lines"].append({
                        "x0": float(line["x0"]),
                        "top": float(line["top"]) + current_y_offset,
                        "x1": float(line["x1"]),
                        "bottom": float(line["bottom"]) + current_y_offset,
                        "stroke_width": float(line.get("width", 0) or 0),
                    })

                # 文字情報の抽出（フォントサイズ・色）
                for char in page.chars:
                    page_data["chars"].append({
                        "text": char["text"],
                        "x0": float(char["x0"]),
                        "top": float(char["top"]) + current_y_offset,
                        "x1": float(char["x1"]),
                        "bottom": float(char["bottom"]) + current_y_offset,
                        "size": float(char["size"]),
                        "fontname": char["fontname"],
                    })

                layout_data["pages"].append(page_data)
                
                # 次のページ用に現在のページの高さ分を加算
                current_y_offset += float(page.height)

        logger.info(f"Layout extraction complete: {len(layout_data['pages'])} page(s)")
        return layout_data
