"""
PDFからテキスト・座標・フォント・色情報を抽出するモジュール。
pdfplumber を使用し、A4方眼Excel変換に必要なレイアウト情報をJSON/Markdownで出力する。
"""

import json
from pathlib import Path
from collections import OrderedDict

import pdfplumber

from src.core.config import config
from src.utils.logger import get_logger

logger = get_logger(__name__)


class PdfExtractor:
    """pdfplumberを使用してPDFからレイアウト情報を抽出する"""

    def __init__(self):
        # 設定ファイルから基準となる方眼のサイズ(pt)を取得
        self.grid_unit = config.grid.unit_pt

    def extract(self, pdf_path: str, out_dir: Path) -> dict:
        """
        PDFを解析し、テキスト・座標・フォント・色・罫線情報を抽出する。

        Returns:
            dict with keys: json_path, md_path, fonts, colors
        """
        logger.info(f"Starting PDF extraction for: {pdf_path}")
        pdf_name = Path(pdf_path).stem

        all_pages = []
        # フォントとカラーの情報を重複なしで保持、挿入順序を維持するためにOrderedDictを使用
        all_fonts = OrderedDict()
        all_colors = OrderedDict()

        with pdfplumber.open(pdf_path) as pdf:
            # ページ累計高さを追跡 (実測値ベースのオフセット)
            cumulative_y = 0.0
            for page_idx, page in enumerate(pdf.pages):
                # ページあたりのオフセット = 前ページまでの実際の累積高さ(pt)
                # ※ target_rows * grid_unit (818.4pt) ではなく page.height (842pt) を使う
                y_offset = cumulative_y
                # 各ページ内のテキスト要素と罫線を抽出
                page_data = self._extract_page(page, page_idx + 1, y_offset)
                all_pages.append(page_data)
                cumulative_y += float(page.height)

                # 抽出した要素からフォント名とサイズの組み合わせ、カラーコードを抽出し一覧化する
                for elem in page_data["elements"]:
                    font_key = f"{elem.get('fontname', 'unknown')}_{elem.get('size', 0)}"
                    if font_key not in all_fonts:
                        all_fonts[font_key] = {
                            "fontname": elem.get("fontname", "unknown"),
                            "size": elem.get("size", 0),
                        }
                    color = elem.get("color")
                    if color and str(color) not in all_colors:
                        all_colors[str(color)] = color

        extracted = {
            "source": pdf_name,
            "pages": all_pages,
            # 各ページの実際の高さ(pt) - executor.pyがここから正確な改ページ行を計算する
            "page_heights": [float(p["height"]) for p in all_pages],
        }

        # JSON出力
        json_path = out_dir / f"{pdf_name}.json"
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(extracted, f, indent=2, ensure_ascii=False)
        logger.info(f"Extraction JSON saved to: {json_path}")

        # Markdown出力
        md_content = self._to_markdown(extracted)
        md_path = out_dir / f"{pdf_name}.md"
        with open(md_path, "w", encoding="utf-8") as f:
            f.write(md_content)
        logger.info(f"Extraction MD saved to: {md_path}")

        return {
            "json_path": str(json_path),
            "md_path": str(md_path),
            "fonts": list(all_fonts.values()),
            "colors": list(all_colors.values()),
        }

    def _extract_page(self, page, page_number: int, y_offset: float) -> dict:
        """1ページ分のテキスト・罫線情報を抽出する"""
        page_data = {
            "page_number": page_number,
            "width": float(page.width),
            "height": float(page.height),
            "elements": [],
            "lines": [],
            "rects": [],
        }

        # テキスト要素（文字単位 → ワード単位で集約）
        words = page.extract_words(
            keep_blank_chars=True,
            extra_attrs=["fontname", "size", "stroking_color", "non_stroking_color"],
        )
        page_width = float(page.width)
        page_height = float(page.height)
        for word in words:
            elem = {
                "text": word["text"],
                "x0": round(float(word["x0"]), 2),
                "top": round(float(word["top"]) + y_offset, 2),
                "x1": round(float(word["x1"]), 2),
                "bottom": round(float(word["bottom"]) + y_offset, 2),
                # PDFサブセット接頭辞（AAAAAA+）を除去してExcelで認識できるフォント名にする
                "fontname": self._clean_fontname(word.get("fontname", "unknown")),
                "size": round(float(word.get("size", 0)), 2),
                # 座標正規化用：LLMがスケーリング変換に使う
                "page_width": round(page_width, 2),
                "page_height": round(page_height, 2),
            }
            # 色情報の取得（non_stroking_color = テキスト塗り色）
            color = word.get("non_stroking_color")
            if color is not None:
                elem["color"] = self._normalize_color(color)
            page_data["elements"].append(elem)

        # 罫線情報
        if page.lines:
            for line in page.lines:
                page_data["lines"].append({
                    "x0": round(float(line["x0"]), 2),
                    "top": round(float(line["top"]) + y_offset, 2),
                    "x1": round(float(line["x1"]), 2),
                    "bottom": round(float(line["bottom"]) + y_offset, 2),
                    "linewidth": round(float(line.get("linewidth", 0)), 2),
                })

        # 矩形情報
        if page.rects:
            for rect in page.rects:
                page_data["rects"].append({
                    "x0": round(float(rect["x0"]), 2),
                    "top": round(float(rect["top"]) + y_offset, 2),
                    "x1": round(float(rect["x1"]), 2),
                    "bottom": round(float(rect["bottom"]) + y_offset, 2),
                    "linewidth": round(float(rect.get("linewidth", 0)), 2),
                })

        return page_data

    def _clean_fontname(self, fontname: str) -> str:
        """PDFサブセット接頭辞（6文字 + '+' 形式）を除去し、Excelで利用可能なフォント名を返す"""
        if "+" in fontname:
            cleaned = fontname.split("+", 1)[1]
            # ハイフン区切りでウェイト指定が含まれる場合も残す（例: Noto-Sans-JP-Thin）
            return cleaned
        return fontname

    def _normalize_color(self, color) -> str:
        """色情報を統一的な16進カラーコード(#RRGGBB)に変換する"""
        if color is None:
            return "#000000"

        # グレースケールの場合は単一の数値として判定
        if isinstance(color, (int, float)):
            val = int(round(float(color) * 255))
            return f"#{val:02X}{val:02X}{val:02X}"

        if isinstance(color, (list, tuple)):
            if len(color) == 1:
                # グレースケール
                val = int(round(float(color[0]) * 255))
                return f"#{val:02X}{val:02X}{val:02X}"
            elif len(color) == 3:
                # RGB
                r = int(round(float(color[0]) * 255))
                g = int(round(float(color[1]) * 255))
                b = int(round(float(color[2]) * 255))
                return f"#{r:02X}{g:02X}{b:02X}"
            elif len(color) == 4:
                # CMYK → RGB変換
                c, m, y, k = [float(v) for v in color]
                r = int(round(255 * (1 - c) * (1 - k)))
                g = int(round(255 * (1 - m) * (1 - k)))
                b = int(round(255 * (1 - y) * (1 - k)))
                return f"#{r:02X}{g:02X}{b:02X}"

        return "#000000"

    def _to_markdown(self, extracted: dict) -> str:
        """抽出結果を人間可読なMarkdown形式に変換する"""
        lines = [f"# {extracted['source']}\n"]

        for page in extracted["pages"]:
            lines.append(f"## Page {page['page_number']} ({page['width']}pt × {page['height']}pt)\n")
            lines.append("### テキスト要素\n")
            for elem in page["elements"]:
                lines.append(
                    f"- `{elem['text']}` @ ({elem['x0']}, {elem['top']}) - ({elem['x1']}, {elem['bottom']}) "
                    f"font={elem['fontname']} size={elem['size']}"
                )
            if page["lines"]:
                lines.append("\n### 罫線\n")
                for line in page["lines"]:
                    lines.append(
                        f"- ({line['x0']}, {line['top']}) → ({line['x1']}, {line['bottom']}) "
                        f"width={line['linewidth']}"
                    )
            if page["rects"]:
                lines.append("\n### 矩形\n")
                for rect in page["rects"]:
                    lines.append(
                        f"- ({rect['x0']}, {rect['top']}) → ({rect['x1']}, {rect['bottom']}) "
                        f"width={rect['linewidth']}"
                    )
            lines.append("")

        return "\n".join(lines)
