"""PDF → Excel 自動生成サービス。

処理の流れ:
  1. PDF 抽出 (pdf_extractor)
  2. グリッド座標計算 (grid)
  3. レイアウト JSON 生成 (layout) + 短スパンエッジフィルタ
  4. Excel 描画 (renderer.excel)
"""

import json
import shutil
from pathlib import Path

from src.core.constants import EDGE_MIN_H_SPAN, EDGE_MIN_V_SPAN
from src.core.edges import filter_short_runs
from src.core.grid import compute_grid_coords, setup_grid_params
from src.core.layout import generate_layout
from src.parser.pdf_extractor import extract_pdf_data
from src.renderer.excel import render_layout_to_xlsx
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

        grid_params, layout_data, output_json_path = \
            self._extract_and_build_layout(pdf_path, pdf_name, out_dir, grid_size)

        try:
            xlsx_path = self._render_excel(layout_data, grid_params, out_dir, pdf_name, grid_size)
        except Exception as e:
            logger.warning(f"⚠️ Excel 生成に失敗しました: {e}")
            xlsx_path = None

        shutil.copy(str(path_obj), str(out_dir / path_obj.name))
        logger.info(f"📄 元PDF コピー完了: {path_obj.name}")

        return {"xlsx_path": str(xlsx_path) if xlsx_path else None,
                "layout_json": str(output_json_path),
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
        for page in layout_data:
            filter_short_runs(page['elements'], EDGE_MIN_H_SPAN, EDGE_MIN_V_SPAN)
        _cleanup_extracted_data(extracted_data)

        output_json_path = out_dir / f"{pdf_name}_{grid_size}_layout.json"
        with open(output_json_path, "w", encoding="utf-8") as f:
            json.dump(layout_data, f, ensure_ascii=False)

        return grid_params, layout_data, output_json_path

    def _render_excel(self, layout_data, grid_params, out_dir, pdf_name, grid_size):
        xlsx_suffix = f"_{grid_size}" if grid_size in ("1pt", "2pt") else ""
        xlsx_path = out_dir / f"{pdf_name}_Python版{xlsx_suffix}.xlsx"
        render_layout_to_xlsx(layout_data, grid_params, str(xlsx_path))
        logger.info(f"✅ Excel 生成完了: {xlsx_path.name}")
        return xlsx_path
