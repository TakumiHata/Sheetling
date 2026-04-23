"""Sheetling パイプラインのファサード。

  auto: PDF解析 → レイアウトJSON自動生成 → Excel描画（AutoLayoutService）
  correct: ビジョンLLM修正指示を適用して Excel を再生成（CorrectionService）

実処理は auto_layout_service / correction_service に委譲する。
"""

from src.core.auto_layout_service import (
    AutoLayoutService,
    _cleanup_extracted_data,
    _collect_content_bounds,
)
from src.core.correction_service import CorrectionService

__all__ = [
    'SheetlingPipeline',
    '_collect_content_bounds',
    '_cleanup_extracted_data',
]


class SheetlingPipeline:
    """PDF から Excel 方眼紙を自動生成するパイプライン。"""

    def __init__(self, output_base_dir: str):
        self._auto = AutoLayoutService(output_base_dir)
        self._correct = CorrectionService(output_base_dir)

    @property
    def output_base_dir(self):
        return self._auto.output_base_dir

    def auto_layout(self, pdf_path: str, in_base_dir: str = "data/in",
                    grid_size: str = "small") -> dict:
        return self._auto.run(pdf_path, in_base_dir=in_base_dir, grid_size=grid_size)

    def apply_corrections(self, pdf_name: str, corrections_json: str,
                          specific_out_dir: str = None, layout_json_name: str = None) -> None:
        return self._correct.apply(pdf_name, corrections_json,
                                   specific_out_dir=specific_out_dir,
                                   layout_json_name=layout_json_name)

    def rerender_after_corrections(self, pdf_name: str, grid_size: str,
                                   specific_out_dir: str = None) -> str:
        return self._correct.rerender(pdf_name, grid_size=grid_size,
                                      specific_out_dir=specific_out_dir)
