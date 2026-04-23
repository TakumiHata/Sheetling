"""PDF → Excel 自動生成サービス。

処理の流れ:
  1. PDF 抽出 (pdf_extractor)
  2. グリッド座標計算 (grid)
  3. レイアウト JSON 生成 (layout)
  4. Excel 描画 (renderer.excel)
  5. LLM レビュー用の画像/プロンプト/罫線プレビュー生成
"""

import json
import shutil
from pathlib import Path

from src.core.edges import enumerate_runs_with_ids
from src.core.grid import compute_grid_coords, setup_grid_params
from src.core.layout import generate_layout
from src.parser.pdf_extractor import extract_pdf_data
from src.renderer.excel import render_layout_to_xlsx
from src.renderer.preview import generate_border_preview, generate_diff_overlay
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

        runs_with_ids = enumerate_runs_with_ids(page_layout.get('elements', []))
        _write_edges_json(runs_with_ids, pdf_name, pn, pdir)
        _write_diff_overlay(pdf_img, runs_with_ids, grid_params,
                            pdf_name, pn, pdir, content_bounds.get(pn, {}))
        _write_prompt_and_corrections(page_layout, grid_params, pdf_name, pn, pdir)


def _write_edges_json(runs_with_ids: list, pdf_name: str, pn: int, pdir) -> None:
    edges_path = pdir / f"{pdf_name}_edges_page{pn}.json"
    edges_path.write_text(
        json.dumps({"edges": runs_with_ids}, ensure_ascii=False, indent=2),
        encoding="utf-8")


def _write_diff_overlay(pdf_img, runs_with_ids, grid_params,
                         pdf_name, pn, pdir, content_bounds):
    if not pdf_img.exists():
        return
    diff_path = pdir / f"{pdf_name}_diff_page{pn}.png"
    try:
        generate_diff_overlay(str(pdf_img), runs_with_ids, grid_params,
                              str(diff_path), content_bounds=content_bounds)
    except Exception as e:
        logger.warning(f"  ページ {pn}: 差分オーバーレイ画像の生成に失敗しました: {e}")


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


def _resolve_out_dir(output_base_dir: Path, pdf_path: str, in_base_dir: str) -> Path:
    path_obj = Path(pdf_path)
    try:
        rel_path = path_obj.parent.relative_to(Path(in_base_dir))
        return output_base_dir / rel_path
    except ValueError:
        return output_base_dir / path_obj.stem


class AutoLayoutService:
    """PDF を読み込み、レイアウト JSON と Excel を生成する。"""

    def __init__(self, output_base_dir: str):
        self.output_base_dir = Path(output_base_dir)

    def run(self, pdf_path: str, in_base_dir: str = "data/in",
            grid_size: str = "small") -> dict:
        logger.info(f"--- [auto] PDF → Excel 高精度自動生成: {Path(pdf_path).name} ---")
        path_obj = Path(pdf_path)
        pdf_name = path_obj.stem
        out_dir = _resolve_out_dir(self.output_base_dir, pdf_path, in_base_dir)
        out_dir.mkdir(parents=True, exist_ok=True)
        prompts_dir = out_dir / "prompts" / grid_size
        prompts_dir.mkdir(parents=True, exist_ok=True)

        grid_params, layout_data, content_bounds, output_json_path = \
            self._extract_and_build_layout(pdf_path, pdf_name, out_dir, grid_size)

        xlsx_path = self._render_excel(layout_data, grid_params, out_dir, pdf_name, grid_size)

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
            f"    1. {{pdf_name}}_diff_page{{N}}.png と {{pdf_name}}_edges_page{{N}}.json と\n"
            f"       {{pdf_name}}_visual_review_page{{N}}.txt(プロンプト) を AI に渡す\n"
            f"    2. AI の出力 JSON を visual_corrections_page{{N}}.json に保存\n"
            f"    3. python -m src.main correct --pdf {pdf_name}"
        )
        return {"xlsx_path": str(xlsx_path), "layout_json": str(output_json_path),
                "grid_params": grid_params}

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

    def _render_excel(self, layout_data, grid_params, out_dir, pdf_name, grid_size):
        xlsx_suffix = f"_{grid_size}" if grid_size in ("1pt", "2pt") else ""
        xlsx_path = out_dir / f"{pdf_name}_Python版{xlsx_suffix}.xlsx"
        render_layout_to_xlsx(layout_data, grid_params, str(xlsx_path))
        logger.info(f"✅ Excel 生成完了: {xlsx_path.name}")
        return xlsx_path
