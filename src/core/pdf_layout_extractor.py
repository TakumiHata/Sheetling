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
                        "top": float(word["top"]),
                        "x1": float(word["x1"]),
                        "bottom": float(word["bottom"]),
                    })

                # 長方形情報の抽出（背景色・ボーダー）
                for rect in page.rects:
                    page_data["rects"].append({
                        "x0": float(rect["x0"]),
                        "top": float(rect["top"]),
                        "x1": float(rect["x1"]),
                        "bottom": float(rect["bottom"]),
                        "fill_color": None,
                        "stroke_color": None,
                        "stroke_width": float(rect.get("width", 0) or 0),
                    })

                # 直線情報の抽出（罫線）
                for line in page.lines:
                    page_data["lines"].append({
                        "x0": float(line["x0"]),
                        "top": float(line["top"]),
                        "x1": float(line["x1"]),
                        "bottom": float(line["bottom"]),
                        "stroke_width": float(line.get("width", 0) or 0),
                        "stroke_color": None,
                    })

                # 文字情報の抽出（フォントサイズ・色）
                for char in page.chars:
                    page_data["chars"].append({
                        "text": char["text"],
                        "x0": float(char["x0"]),
                        "top": float(char["top"]),
                        "x1": float(char["x1"]),
                        "bottom": float(char["bottom"]),
                        "size": float(char["size"]),
                        "fontname": char["fontname"],
                    })

                layout_data["pages"].append(page_data)

        logger.info(f"Layout extraction complete: {len(layout_data['pages'])} page(s)")
        return layout_data

    def _to_hex(self, color) -> str | None:
        """
        pdfplumberの色値（[R, G, B] 0.0-1.0 or 0-255、またはグレースケール）を
        '#RRGGBB' 形式のHEX文字列に変換する。
        """
        if color is None:
            return None

        try:
            # グレースケール（単一値）
            if isinstance(color, (int, float)):
                v = int(color * 255) if color <= 1.0 else int(color)
                return "#{:02X}{:02X}{:02X}".format(v, v, v)

            # RGB配列
            if isinstance(color, (list, tuple)):
                if len(color) == 1:
                    # グレースケール配列 [0.5]
                    v = int(color[0] * 255) if color[0] <= 1.0 else int(color[0])
                    return "#{:02X}{:02X}{:02X}".format(v, v, v)
                elif len(color) == 3:
                    if all(isinstance(x, (float, int)) and 0 <= x <= 1.0 for x in color):
                        rgb = [int(x * 255) for x in color]
                    else:
                        rgb = [int(x) for x in color]
                    return "#{:02X}{:02X}{:02X}".format(rgb[0], rgb[1], rgb[2])
                elif len(color) == 4:
                    # CMYK → RGB（簡易変換）
                    c, m, y, k = color
                    r = int(255 * (1 - c) * (1 - k))
                    g = int(255 * (1 - m) * (1 - k))
                    b = int(255 * (1 - y) * (1 - k))
                    return "#{:02X}{:02X}{:02X}".format(r, g, b)

            return None
        except (ValueError, TypeError):
            return None
