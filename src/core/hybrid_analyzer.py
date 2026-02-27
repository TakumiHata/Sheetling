import json
import math
from pathlib import Path
from src.core.pdf_layout_extractor import PdfLayoutExtractor
from src.core.markitdown_parser import MarkItDownParser
from src.utils.logger import get_logger

logger = get_logger(__name__)


class HybridAnalyzer:
    """
    MarkItDown（テキスト抽出）と pdfplumber（座標・色彩・レイアウト）の結果を統合し、
    指示書のJSONスキーマに準拠した構造化データを出力する。
    """

    def __init__(self, inter_md_dir: str, inter_json_dir: str, grid_size: float = 8.5):
        self.inter_md_dir = Path(inter_md_dir)
        self.inter_json_dir = Path(inter_json_dir)

        self.pdf_extractor = PdfLayoutExtractor(str(self.inter_json_dir))
        self.mid_parser = MarkItDownParser(str(self.inter_md_dir))

        # グリッド設定 (ポイント単位: 1pt = 1/72 inch)
        # 8.5pt ≒ 3mm。セルが小さすぎない適度な方眼サイズ。
        self.grid_size = grid_size

    def analyze(self, pdf_path: str) -> dict:
        """
        MarkItDown + pdfplumber でPDFを解析し、指示書スキーマに準拠したJSONを出力する。

        Returns:
            dict with keys: json_path, md_path
        """
        logger.info(f"Starting hybrid analysis for: {pdf_path}")
        pdf_name = Path(pdf_path).stem

        # Phase 1: MarkItDown でテキスト抽出（最優先テキストソース）
        mid_md_path = self.mid_parser.parse(pdf_path)
        logger.info(f"Phase 1 complete: MarkItDown MD -> {mid_md_path}")

        # Phase 2: pdfplumber で座標・色彩・幾何学情報を抽出
        layout_data = self.pdf_extractor.extract(pdf_path)
        logger.info(f"Phase 2 complete: pdfplumber layout extracted")

        # Phase 3: 統合・正規化 → 指示書スキーマ準拠JSON
        output_data = self._build_schema_json(pdf_name, layout_data)

        # 出力
        output_json_path = self.inter_json_dir / f"{pdf_name}.json"
        with open(output_json_path, "w", encoding="utf-8") as f:
            json.dump(output_data, f, indent=2, ensure_ascii=False)

        logger.info(f"Integrated JSON saved to: {output_json_path}")

        return {
            "json_path": str(output_json_path),
            "md_path": mid_md_path,
        }

    def _build_schema_json(self, pdf_name: str, layout_data: dict) -> dict:
        """
        pdfplumberの生データを指示書のJSONスキーマに変換する。

        出力スキーマ:
        {
          "pdf_name": str,
          "pages": [{
            "page": { "width_pt", "height_pt", "grid_unit_pt", "grid_cols", "grid_rows" },
            "elements": [{
              "type": "rect" | "text" | "line",
              "bbox": { "x0", "y0", "x1", "y1" },
              "grid_bbox": { "col_start", "row_start", "col_end", "row_end" },
              "style": { "stroke_width", "border" },
              "text": str | null,
              "font_size": float | null
            }]
          }]
        }
        """
        output = {
            "pdf_name": pdf_name,
            "pages": []
        }

        for page_data in layout_data["pages"]:
            width = page_data["width"]
            height = page_data["height"]
            grid_cols = math.ceil(width / self.grid_size)
            grid_rows = math.ceil(height / self.grid_size)

            page_output = {
                "page": {
                    "page_number": page_data["page_number"],
                    "width_pt": width,
                    "height_pt": height,
                    "grid_unit_pt": self.grid_size,
                    "grid_cols": grid_cols,
                    "grid_rows": grid_rows,
                },
                "elements": []
            }

            # --- rects（背景ボックス）→ type: "rect" ---
            for rect in page_data.get("rects", []):
                bbox = self._make_bbox(rect)
                grid_bbox = self._to_grid_bbox(bbox)
                border = self._detect_border_from_rect(rect)

                elem = {
                    "type": "rect",
                    "bbox": bbox,
                    "grid_bbox": grid_bbox,
                    "style": {
                        "stroke_width": rect.get("stroke_width", 0),
                        "border": border,
                    },
                    "text": None,
                    "font_size": None,
                }
                page_output["elements"].append(elem)

            # --- lines（罫線）→ type: "line" ---
            for line in page_data.get("lines", []):
                bbox = self._make_bbox(line)
                grid_bbox = self._to_grid_bbox(bbox)

                elem = {
                    "type": "line",
                    "bbox": bbox,
                    "grid_bbox": grid_bbox,
                    "style": {
                        "stroke_width": line.get("stroke_width", 0),
                        "border": self._detect_border_from_line(line),
                    },
                    "text": None,
                    "font_size": None,
                }
                page_output["elements"].append(elem)

            # --- words（テキスト）→ グループ化して type: "text" ---
            text_elements = self._group_words_to_text_elements(
                page_data.get("words", []),
                page_data.get("chars", []),
            )
            page_output["elements"].extend(text_elements)

            output["pages"].append(page_output)

        return output

    def _make_bbox(self, obj: dict) -> dict:
        """pdfplumberオブジェクトからbbox辞書を生成する。"""
        return {
            "x0": round(obj["x0"], 2),
            "y0": round(obj["top"], 2),
            "x1": round(obj["x1"], 2),
            "y1": round(obj["bottom"], 2),
        }

    def _to_grid_bbox(self, bbox: dict) -> dict:
        """bbox (points) をグリッドインデックスに変換（スナップ）。"""
        return {
            "col_start": int(math.floor(bbox["x0"] / self.grid_size)),
            "row_start": int(math.floor(bbox["y0"] / self.grid_size)),
            "col_end": int(math.ceil(bbox["x1"] / self.grid_size)),
            "row_end": int(math.ceil(bbox["y1"] / self.grid_size)),
        }

    def _detect_border_from_rect(self, rect: dict) -> dict:
        """rectのstroke情報からborderの有無を推定する。"""
        has_stroke = rect.get("stroke_width", 0) > 0
        return {
            "top": has_stroke,
            "right": has_stroke,
            "bottom": has_stroke,
            "left": has_stroke,
        }

    def _detect_border_from_line(self, line: dict) -> dict:
        """lineの方向（水平/垂直）からborderの向きを推定する。"""
        x0, y0 = line["x0"], line["top"]
        x1, y1 = line["x1"], line["bottom"]

        is_horizontal = abs(y1 - y0) < 1.0
        is_vertical = abs(x1 - x0) < 1.0

        return {
            "top": is_horizontal,
            "right": is_vertical,
            "bottom": is_horizontal,
            "left": is_vertical,
        }

    def _group_words_to_text_elements(self, words: list, chars: list) -> list:
        """
        pdfplumberの単語リストを、行単位でグループ化してtext要素に変換する。

        1. Y座標（top）で行グループ化（誤差3pt以内を同一行）
        2. 各行内でX座標でソートし、近接する単語を統合
        3. chars情報からフォントサイズを推定
        """
        if not words:
            return []

        # charsからフォントサイズマップを構築（座標 → サイズ）
        font_size_map = self._build_font_size_map(chars)

        elements = []

        # 1. topでソート → 行グループ化
        sorted_words = sorted(words, key=lambda w: w["top"])
        lines = []
        current_line = [sorted_words[0]]

        for i in range(1, len(sorted_words)):
            if abs(sorted_words[i]["top"] - current_line[-1]["top"]) < 3.0:
                current_line.append(sorted_words[i])
            else:
                lines.append(current_line)
                current_line = [sorted_words[i]]
        lines.append(current_line)

        # 2. 各行内でx0ソート → 近接テキスト統合
        for line_words in lines:
            sorted_line = sorted(line_words, key=lambda w: w["x0"])
            if not sorted_line:
                continue

            current_group = [sorted_line[0]]
            for i in range(1, len(sorted_line)):
                prev = sorted_line[i - 1]
                curr = sorted_line[i]

                # 5pt以内の間隔 → 同一テキストブロック
                if (curr["x0"] - prev["x1"]) < 5.0:
                    current_group.append(curr)
                else:
                    elements.append(self._words_to_element(current_group, font_size_map))
                    current_group = [curr]
            elements.append(self._words_to_element(current_group, font_size_map))

        return elements

    def _words_to_element(self, group: list, font_size_map: dict) -> dict:
        """単語グループをtext要素に変換する。"""
        text = "".join(w["text"] for w in group)
        bbox = {
            "x0": round(min(w["x0"] for w in group), 2),
            "y0": round(min(w["top"] for w in group), 2),
            "x1": round(max(w["x1"] for w in group), 2),
            "y1": round(max(w["bottom"] for w in group), 2),
        }
        grid_bbox = self._to_grid_bbox(bbox)

        # フォントサイズの推定（最頻値）
        font_size = self._estimate_font_size(bbox, font_size_map)

        return {
            "type": "text",
            "bbox": bbox,
            "grid_bbox": grid_bbox,
            "style": {
                "stroke_width": 0,
                "border": {"top": False, "right": False, "bottom": False, "left": False},
            },
            "text": text,
            "font_size": font_size,
        }

    def _build_font_size_map(self, chars: list) -> dict:
        """
        charsリストからY座標帯ごとのフォントサイズマップを構築する。
        key: (Y帯の中心をgrid_sizeでスナップした値)
        value: 最頻出フォントサイズ
        """
        if not chars:
            return {}

        band_sizes: dict[int, list[float]] = {}
        for char in chars:
            band_key = int(char["top"] / self.grid_size)
            if band_key not in band_sizes:
                band_sizes[band_key] = []
            band_sizes[band_key].append(char["size"])

        result = {}
        for band_key, sizes in band_sizes.items():
            # 最頻値を使用
            from collections import Counter
            counter = Counter(round(s, 1) for s in sizes)
            most_common = counter.most_common(1)[0][0]
            result[band_key] = most_common

        return result

    def _estimate_font_size(self, bbox: dict, font_size_map: dict) -> float | None:
        """bboxの位置から対応するフォントサイズを推定する。"""
        if not font_size_map:
            return None

        band_key = int(bbox["y0"] / self.grid_size)
        # 近傍のバンドも探索
        for offset in [0, -1, 1]:
            if (band_key + offset) in font_size_map:
                return font_size_map[band_key + offset]
        return None
