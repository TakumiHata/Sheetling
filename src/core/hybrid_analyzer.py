import json
from pathlib import Path
from src.core.docling_parser import DoclingParser
from src.core.pdf_layout_extractor import PdfLayoutExtractor
from src.core.markitdown_parser import MarkItDownParser
from src.utils.logger import get_logger

logger = get_logger(__name__)

class HybridAnalyzer:
    def __init__(self, inter_md_dir: str, inter_json_dir: str):
        self.inter_md_dir = Path(inter_md_dir)
        self.inter_json_dir = Path(inter_json_dir)
        
        self.docling = DoclingParser(str(self.inter_md_dir), str(self.inter_json_dir))
        self.pdf_extractor = PdfLayoutExtractor(str(self.inter_json_dir))
        self.mid_parser = MarkItDownParser(str(self.inter_md_dir))
        
        # グリッド設定 (ポイント単位: 1pt = 1/72 inch)
        # 5pt ≒ 1.76mm。方眼の1マスに対応。
        self.grid_size = 5.0 

    def analyze(self, pdf_path: str):
        """
        全ツールを総動員してマスタ解析JSONを生成する。
        """
        logger.info(f"Starting hybrid analysis for: {pdf_path}")
        pdf_name = Path(pdf_path).stem
        
        # 各ツールでの抽出
        dl_md, dl_json = self.docling.parse(pdf_path)
        layout_json = self.pdf_extractor.extract(pdf_path)
        mid_md = self.mid_parser.parse(pdf_path)
        
        # データのロード
        with open(dl_json, "r", encoding="utf-8") as f:
            docling_data = json.load(f)
        with open(layout_json, "r", encoding="utf-8") as f:
            visual_data = json.load(f)
            
        # マスター構造の構築
        master_data = {
            "pdf_name": pdf_name,
            "grid_size": self.grid_size,
            "pages": []
        }
        
        for v_page in visual_data["pages"]:
            page_master = {
                "page_number": v_page["page_number"],
                "width": v_page["width"],
                "height": v_page["height"],
                "grid_cols": int(v_page["width"] / self.grid_size) + 1,
                "grid_rows": int(v_page["height"] / self.grid_size) + 1,
                "elements": []
            }
            
            # ビジュアル要素（長方形、背景色）の正規化
            for rect in v_page["rects"]:
                page_master["elements"].append({
                    "type": "box",
                    "bbox": rect,
                    "grid_range": self._to_grid_range(rect),
                    "color": rect.get("non_stroking_color")
                })
            
            # テキスト要素（行単位での統合）
            # 1. まずY座標（top）でグループ化（誤差3pt以内を同じ行とする）
            raw_words = v_page["words"]
            if raw_words:
                temp_lines = []
                # topでソート
                sorted_by_top = sorted(raw_words, key=lambda x: x["top"])
                
                if sorted_by_top:
                    current_raw_line = [sorted_by_top[0]]
                    for i in range(1, len(sorted_by_top)):
                        if abs(sorted_by_top[i]["top"] - current_raw_line[-1]["top"]) < 3.0:
                            current_raw_line.append(sorted_by_top[i])
                        else:
                            temp_lines.append(current_raw_line)
                            current_raw_line = [sorted_by_top[i]]
                    temp_lines.append(current_raw_line)

                # 2. 各行内でX座標（x0）でソートし、水平方向に近い単語を統合
                for raw_line in temp_lines:
                    line_words = sorted(raw_line, key=lambda x: x["x0"])
                    
                    if not line_words:
                        continue
                        
                    current_group = [line_words[0]]
                    for i in range(1, len(line_words)):
                        prev = line_words[i-1]
                        curr = line_words[i]
                        
                        # 水平方向に近い（5.0pt以内）かつ重複していない場合は統合
                        if (curr["x0"] - prev["x1"]) < 5.0:
                            current_group.append(curr)
                        else:
                            self._append_text_element(page_master, current_group)
                            current_group = [curr]
                    self._append_text_element(page_master, current_group)
            
            master_data["pages"].append(page_master)
            
        master_output_path = self.inter_json_dir / f"{pdf_name}_master.json"
        with open(master_output_path, "w", encoding="utf-8") as f:
            json.dump(master_data, f, indent=2, ensure_ascii=False)
            
        logger.info(f"Master hybrid analysis saved to: {master_output_path}")
        return str(master_output_path), dl_md, mid_md

    def _append_text_element(self, page_master: dict, group: list):
        if not group: return
        text = "".join([w["text"] for w in group])
        bbox = {
            "x0": min(w["x0"] for w in group),
            "top": min(w["top"] for w in group),
            "x1": max(w["x1"] for w in group),
            "bottom": max(w["bottom"] for w in group)
        }
        page_master["elements"].append({
            "type": "text",
            "content": text,
            "bbox": bbox,
            "grid_range": self._to_grid_range(bbox),
            "style": {"size": 10, "font": "Default"}
        })

    def _to_grid_range(self, bbox: dict):
        """
        座標(points)をグリッド(方眼セル)のインデックスに変換・スナップ。
        """
        import math
        return {
            "start_col": int(math.floor(bbox["x0"] / self.grid_size)),
            "start_row": int(math.floor(bbox["top"] / self.grid_size)),
            "end_col": int(math.ceil(bbox["x1"] / self.grid_size)),
            "end_row": int(math.ceil(bbox["bottom"] / self.grid_size))
        }
