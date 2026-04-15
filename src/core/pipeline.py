"""
Sheetling パイプライン。

  auto: PDF解析 → レイアウトJSON自動生成 → Excel描画
  correct: ビジョンLLM修正指示を適用して Excel を再生成
"""

import json
import shutil
from pathlib import Path

from src.core.grid import compute_grid_coords, setup_grid_params
from src.core.layout import generate_layout
from src.parser.pdf_extractor import extract_pdf_data
from src.renderer.excel import render_layout_to_xlsx
from src.renderer.preview import generate_border_preview
from src.templates.prompts import VISUAL_REVIEW_PROMPT
from src.utils.logger import get_logger

logger = get_logger(__name__)


def _collect_content_bounds(extracted_data: dict, grid_params: dict) -> dict:
    bounds = {}
    for page in extracted_data['pages']:
        pn = page.get('page_number', 1)
        bounds[pn] = {
            'min_x': page.get('_content_min_x', 0.0),
            'min_y': page.get('_content_min_y', 0.0),
            'grid_w': page.get('_content_grid_w', float(page['width']) / grid_params['max_cols']),
            'grid_h': page.get('_content_grid_h', float(page['height']) / grid_params['max_rows']),
            'page_width': float(page['width']),
            'page_height': float(page['height']),
        }
    return bounds


def _cleanup_extracted_data(extracted_data: dict) -> None:
    for page in extracted_data['pages']:
        for key in ('table_data', 'table_data_raw', 'table_row_y_positions',
                     'table_cells', '_content_min_x', '_content_min_y',
                     '_content_grid_w', '_content_grid_h'):
            page.pop(key, None)


def _generate_pdf_page_images(pdf_path: str, pdf_name: str, prompts_dir: Path) -> None:
    import pdfplumber as _pdfplumber
    with _pdfplumber.open(pdf_path) as pdf:
        for pg in pdf.pages:
            pn = pg.page_number
            pdir = prompts_dir / f"page_{pn}"
            pdir.mkdir(parents=True, exist_ok=True)
            img = pg.to_image(resolution=144)
            img.save(str(pdir / f"{pdf_name}_page{pn}.png"))
    logger.info(f"  PDF ページ画像を生成しました: {prompts_dir}/page_N/")


def _generate_review_materials(layout_data, grid_params, pdf_name, prompts_dir, content_bounds):
    for page_layout in layout_data:
        pn = page_layout.get('page_number', 1)
        pdir = prompts_dir / f"page_{pn}"
        pdir.mkdir(parents=True, exist_ok=True)

        pdf_img = pdir / f"{pdf_name}_page{pn}.png"
        preview = pdir / f"{pdf_name}_excel_page{pn}.png"
        try:
            generate_border_preview(page_layout, grid_params, str(preview),
                                    pdf_image_path=str(pdf_img),
                                    content_bounds=content_bounds.get(pn, {}))
        except Exception as e:
            logger.warning(f"  ページ {pn}: 罫線プレビュー生成に失敗しました: {e}")

        _write_prompt_and_corrections(page_layout, grid_params, pdf_name, pn, pdir)


def _write_prompt_and_corrections(page_layout, grid_params, pdf_name, pn, pdir):
    gp = dict(grid_params)
    gp.setdefault('position_tolerance_cells', '1〜2')
    elems = page_layout.get('elements', [])
    end_rows = [e.get('end_row', e.get('row', 1)) for e in elems if e.get('type') == 'border_rect']
    end_cols = [e.get('end_col', e.get('col', 1)) for e in elems if e.get('type') == 'border_rect']
    gp['content_max_row'] = max(end_rows) if end_rows else grid_params['max_rows']
    gp['content_max_col'] = max(end_cols) if end_cols else grid_params['max_cols']
    prompt_text = VISUAL_REVIEW_PROMPT.format(page_number=pn, **gp)
    (pdir / f"{pdf_name}_visual_review_page{pn}.txt").write_text(prompt_text, encoding="utf-8")

    corr_path = pdir / f"{pdf_name}_visual_corrections_page{pn}.json"
    if not corr_path.exists():
        corr_path.write_text('{"corrections": []}', encoding="utf-8")


class SheetlingPipeline:
    """PDF から Excel 方眼紙を自動生成するパイプライン。"""

    def __init__(self, output_base_dir: str):
        self.output_base_dir = Path(output_base_dir)

    def _resolve_out_dir(self, pdf_path: str, in_base_dir: str):
        path_obj = Path(pdf_path)
        try:
            rel_path = path_obj.parent.relative_to(Path(in_base_dir))
            return self.output_base_dir / rel_path
        except ValueError:
            return self.output_base_dir / path_obj.stem

    def _extract_and_build_layout(self, pdf_path, pdf_name, out_dir, grid_size):
        extracted_data = extract_pdf_data(pdf_path)
        with open(out_dir / f"{pdf_name}_extracted.json", "w", encoding="utf-8") as f:
            json.dump(extracted_data, f, indent=2, ensure_ascii=False)

        grid_params = setup_grid_params(extracted_data['pages'][0], grid_size)
        for page in extracted_data['pages']:
            compute_grid_coords(page, grid_params['max_rows'], grid_params['max_cols'])

        with open(out_dir / f"{pdf_name}_{grid_size}_grid_params.json", "w", encoding="utf-8") as f:
            json.dump(grid_params, f, ensure_ascii=False)

        layout_json_str = generate_layout(extracted_data, grid_params)
        layout_data = json.loads(layout_json_str)
        content_bounds = _collect_content_bounds(extracted_data, grid_params)
        _cleanup_extracted_data(extracted_data)

        output_json_path = out_dir / f"{pdf_name}_{grid_size}_layout.json"
        with open(output_json_path, "w", encoding="utf-8") as f:
            f.write(layout_json_str)

        return grid_params, layout_data, content_bounds, output_json_path

    def auto_layout(self, pdf_path: str, in_base_dir: str = "data/in", grid_size: str = "small") -> dict:
        logger.info(f"--- [auto] PDF → Excel 高精度自動生成: {Path(pdf_path).name} ---")
        path_obj = Path(pdf_path)
        pdf_name = path_obj.stem
        out_dir = self._resolve_out_dir(pdf_path, in_base_dir)
        out_dir.mkdir(parents=True, exist_ok=True)
        prompts_dir = out_dir / "prompts" / grid_size
        prompts_dir.mkdir(parents=True, exist_ok=True)

        grid_params, layout_data, content_bounds, output_json_path = \
            self._extract_and_build_layout(pdf_path, pdf_name, out_dir, grid_size)

        xlsx_suffix = f"_{grid_size}" if grid_size in ("1pt", "2pt") else ""
        xlsx_path = out_dir / f"{pdf_name}_Python版{xlsx_suffix}.xlsx"
        render_layout_to_xlsx(layout_data, grid_params, str(xlsx_path))
        logger.info(f"✅ Excel 生成完了: {xlsx_path.name}")

        shutil.copy(str(path_obj), str(out_dir / path_obj.name))
        logger.info(f"📄 元PDF コピー完了: {path_obj.name}")

        try:
            _generate_pdf_page_images(pdf_path, pdf_name, prompts_dir)
        except Exception as e:
            logger.warning(f"PDF ページ画像の生成に失敗しました: {e}")

        _generate_review_materials(layout_data, grid_params, pdf_name, prompts_dir, content_bounds)
        logger.info(
            f"  [review 素材] prompts/{grid_size}/page_N/ に出力しました\n"
            f"  次のステップ:\n"
            f"    1. PDF 画像と罫線プレビューを AI に渡し比較させる\n"
            f"    2. AI の出力 JSON を visual_corrections_page{{N}}.json に保存\n"
            f"    3. python -m src.main correct --pdf {pdf_name} --grid-size {grid_size}"
        )
        return {"xlsx_path": str(xlsx_path), "layout_json": str(output_json_path), "grid_params": grid_params}

    def apply_corrections(self, pdf_name: str, corrections_json: str,
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
            applied += self._apply_single_correction(c, page_map, content_bounds)

        output_json_path.write_text(json.dumps(layout, ensure_ascii=False), encoding="utf-8")
        logger.info(f"[correct] {applied} 件の修正を適用しました: {output_json_path}")

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

    def _apply_single_correction(self, c: dict, page_map: dict, content_bounds: dict) -> int:
        action = c.get("action")
        page_no = c.get("page", 1)
        elements = page_map.get(page_no)
        if elements is None:
            logger.warning(f"[correct] ページ {page_no} が見つかりません。スキップします。")
            return 0

        if action == "add_text":
            elements.append({
                "type": "text", "content": c["content"],
                "row": c["row"], "col": c["col"], "end_col": c["col"] + len(c["content"]),
            })
            return 1
        if action == "fix_text":
            for elem in elements:
                if elem.get("type") == "text" and elem["row"] == c["row"] and elem["col"] == c["col"]:
                    elem["row"] = c["new_row"]
                    elem["col"] = c["new_col"]
                    return 1
            return 0
        if action == "add_border":
            bounds = content_bounds.get(page_no, {})
            end_row = min(c.get("end_row") or c.get("row_end", c["row"]), bounds.get("max_row", 9999))
            end_col = min(c.get("end_col") or c.get("col_end", c["col"]), bounds.get("max_col", 9999))
            elements.append({
                "type": "border_rect", "row": c["row"], "end_row": end_row,
                "col": c["col"], "end_col": end_col,
                "borders": c.get("borders", {"top": True, "bottom": True, "left": True, "right": True}),
            })
            return 1
        if action == "remove_border":
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
        return 0

    def rerender_after_corrections(self, pdf_name: str, grid_size: str,
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
