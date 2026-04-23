"""ビジョンLLM が出力した修正指示を layout JSON に適用するサービス。

action ごとの処理はモジュールレベルの関数に分解し、_ACTIONS で
ディスパッチする。各関数は (elements, correction, content_bounds) を
受け取り、適用した件数を返す。
"""

import json
from pathlib import Path

from src.renderer.excel import render_layout_to_xlsx
from src.utils.logger import get_logger

logger = get_logger(__name__)


def _apply_add_text(elements: list, c: dict, _bounds: dict) -> int:
    elements.append({
        "type": "text", "content": c["content"],
        "row": c["row"], "col": c["col"], "end_col": c["col"] + len(c["content"]),
    })
    return 1


def _apply_fix_text(elements: list, c: dict, _bounds: dict) -> int:
    for elem in elements:
        if elem.get("type") == "text" and elem["row"] == c["row"] and elem["col"] == c["col"]:
            elem["row"] = c["new_row"]
            elem["col"] = c["new_col"]
            return 1
    return 0


def _apply_add_border(elements: list, c: dict, bounds: dict) -> int:
    end_row = min(c.get("end_row") or c.get("row_end", c["row"]), bounds.get("max_row", 9999))
    end_col = min(c.get("end_col") or c.get("col_end", c["col"]), bounds.get("max_col", 9999))
    elements.append({
        "type": "border_rect", "row": c["row"], "end_row": end_row,
        "col": c["col"], "end_col": end_col,
        "borders": c.get("borders", {"top": True, "bottom": True, "left": True, "right": True}),
    })
    return 1


def _apply_remove_border(elements: list, c: dict, _bounds: dict) -> int:
    before = len(elements)
    r, er = c["row"], c.get("end_row") or c.get("row_end", c["row"])
    co, ec = c["col"], c.get("end_col") or c.get("col_end", c["col"])
    elements[:] = [
        e for e in elements
        if not (e.get("type") == "border_rect"
                and e["row"] >= r and e["end_row"] <= er
                and e["col"] >= co and e["end_col"] <= ec)
    ]
    return before - len(elements)


_ACTIONS = {
    "add_text":      _apply_add_text,
    "fix_text":      _apply_fix_text,
    "add_border":    _apply_add_border,
    "remove_border": _apply_remove_border,
}


class CorrectionService:
    """修正 JSON の適用とリレンダーを担当。"""

    def __init__(self, output_base_dir: str):
        self.output_base_dir = Path(output_base_dir)

    def apply(self, pdf_name: str, corrections_json: str,
              specific_out_dir: str = None, layout_json_name: str = None) -> None:
        out_dir = Path(specific_out_dir) if specific_out_dir else self.output_base_dir / pdf_name
        json_name = layout_json_name or f"{pdf_name}_layout.json"
        output_json_path = out_dir / json_name

        if not output_json_path.exists():
            raise FileNotFoundError(f"_layout.json が見つかりません: {output_json_path}")

        layout = json.loads(output_json_path.read_text(encoding="utf-8"))
        corrections = self._parse_corrections(corrections_json)
        page_map = {p["page_number"]: p["elements"] for p in layout}
        content_bounds = self._compute_content_bounds(layout)

        applied = 0
        for c in corrections:
            applied += self._dispatch(c, page_map, content_bounds)

        output_json_path.write_text(json.dumps(layout, ensure_ascii=False), encoding="utf-8")
        logger.info(f"[correct] {applied} 件の修正を適用しました: {output_json_path}")

    def rerender(self, pdf_name: str, grid_size: str,
                 specific_out_dir: str = None) -> str:
        logger.info(f"--- [correct/rerender] Excel 再生成: {pdf_name} ({grid_size}) ---")
        out_dir = Path(specific_out_dir) if specific_out_dir else self.output_base_dir / pdf_name

        layout_path = out_dir / f"{pdf_name}_{grid_size}_layout.json"
        grid_params_path = out_dir / f"{pdf_name}_{grid_size}_grid_params.json"
        xlsx_suffix = f"_{grid_size}" if grid_size in ("1pt", "2pt") else ""
        xlsx_path = out_dir / f"{pdf_name}_Python版{xlsx_suffix}.xlsx"

        if not layout_path.exists():
            raise FileNotFoundError(f"layout JSON が見つかりません: {layout_path}")
        if not grid_params_path.exists():
            raise FileNotFoundError(f"grid_params JSON が見つかりません: {grid_params_path}")

        layout = json.loads(layout_path.read_text(encoding="utf-8"))
        grid_params = json.loads(grid_params_path.read_text(encoding="utf-8"))

        render_layout_to_xlsx(layout, grid_params, str(xlsx_path))
        logger.info(f"✅ correct/rerender 完了: {xlsx_path.name}")
        return str(xlsx_path)

    def _parse_corrections(self, corrections_json: str) -> list:
        try:
            data = json.loads(corrections_json)
            return data.get("corrections", [])
        except json.JSONDecodeError as e:
            raise ValueError(f"corrections JSON のパースに失敗しました: {e}")

    def _compute_content_bounds(self, layout: list) -> dict:
        bounds = {}
        for p in layout:
            pn = p["page_number"]
            border_elems = [e for e in p["elements"] if e.get("type") == "border_rect"]
            if border_elems:
                bounds[pn] = {
                    "max_row": max(e.get("end_row", e["row"]) for e in border_elems),
                    "max_col": max(e.get("end_col", e["col"]) for e in border_elems),
                }
            else:
                bounds[pn] = {"max_row": 9999, "max_col": 9999}
        return bounds

    def _dispatch(self, c: dict, page_map: dict, content_bounds: dict) -> int:
        action = c.get("action")
        page_no = c.get("page", 1)
        elements = page_map.get(page_no)
        if elements is None:
            logger.warning(f"[correct] ページ {page_no} が見つかりません。スキップします。")
            return 0
        handler = _ACTIONS.get(action)
        if handler is None:
            return 0
        return handler(elements, c, content_bounds.get(page_no, {}))
