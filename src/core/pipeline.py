"""Sheetling パイプラインのファサード。

  auto: PDF解析 → レイアウトJSON自動生成 → Excel描画（AutoLayoutService）

実処理は auto_layout_service に委譲する。
"""

from src.core.auto_layout_service import AutoLayoutService

__all__ = ['SheetlingPipeline']


class SheetlingPipeline:
    """PDF から Excel 方眼紙を自動生成するパイプライン。"""

    def __init__(self, output_base_dir: str):
        self._auto = AutoLayoutService(output_base_dir)

    @property
    def output_base_dir(self):
        return self._auto.output_base_dir

    def auto_layout(self, pdf_path: str, in_base_dir: str = "data/in",
                    grid_size: str = "1pt") -> dict:
        return self._auto.run(pdf_path, in_base_dir=in_base_dir, grid_size=grid_size)
