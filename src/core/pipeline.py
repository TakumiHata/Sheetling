import json
from pathlib import Path
from src.core.hybrid_analyzer import HybridAnalyzer
from src.core.placement_generator import PlacementGenerator, format_table_structure_summary
from src.core.code_generator import CodeGenerator
from src.core.prompt_builder import PromptBuilder
from src.core.config import config
from src.utils.logger import get_logger

logger = get_logger(__name__)


class SheetlingPipeline:
    """
    PDFからExcel生成スクリプト（およびプロンプト）を作成する
    一連のパイプライン処理を統括するクラス。
    """

    def __init__(self, md_dir: str, json_dir: str, prompt_dir: str):
        self.md_dir = Path(md_dir)
        self.json_dir = Path(json_dir)
        self.prompt_dir = Path(prompt_dir)

        # Core components
        self.analyzer = HybridAnalyzer(str(self.md_dir), str(self.json_dir))
        self.placement_gen = PlacementGenerator()
        self.code_gen = CodeGenerator()
        self.prompt_builder = PromptBuilder(str(self.prompt_dir))

    def run(self, pdf_path: str) -> dict:
        """
        単一のPDFファイルに対するパイプライン処理を実行する。
        """
        logger.info(f"--- Processing: {Path(pdf_path).name} ---")
        pdf_name = Path(pdf_path).stem

        # Phase 1-2: ハイブリッド解析（MD + JSON）
        analyze_result = self.analyzer.analyze(pdf_path)
        md_path = analyze_result["md_path"]
        json_path = analyze_result["json_path"]

        # 抽出データの読み込み
        md_content = Path(md_path).read_text(encoding="utf-8")
        with open(json_path, "r", encoding="utf-8") as f:
            json_data = json.load(f)

        # 圧縮（余分な空白等の除去）
        compressed_json = self._compress_json(json_data)

        # グリッドおよびページ情報の導出
        # JSON側の grid_unit_pt があれば取得、なければ config のデフォルト値を使用
        grid_unit_pt = config.grid.unit_pt
        grid_cols = config.grid.target_cols
        grid_rows = config.grid.target_rows
        page_breaks = json_data.get("page_breaks", [])

        if json_data.get("pages") and json_data["pages"][0].get("page"):
            page_info = json_data["pages"][0]["page"]
            grid_unit_pt = page_info.get("grid_unit_pt", grid_unit_pt)
            grid_cols = page_info.get("grid_cols", grid_cols)
            if not page_breaks:
                grid_rows = page_info.get("grid_rows", grid_rows)

        if page_breaks:
            grid_rows = page_breaks[-1]

        page_count = len(json_data.get("pages", []))
        output_filename = f"{pdf_name}.xlsx"
        
        # Excelセルの固定サイズ（正方形への調整など）
        row_height = config.excel.row_height_pt
        col_width = config.excel.col_width_chars
        scale_factor = row_height / grid_unit_pt if grid_unit_pt else 1.0

        # Phase 3: 配置命令リストの生成
        placement_result = self.placement_gen.generate(compressed_json)
        if placement_result.warnings:
            for w in placement_result.warnings:
                logger.warning(f"配置命令: {w}")

        # テーブル構造サマリーのテキスト化
        table_summary = format_table_structure_summary(placement_result)

        # Phase 4: コードの生成
        generated_code = self.code_gen.generate(
            placement_result=placement_result,
            grid_cols=grid_cols,
            grid_rows=grid_rows,
            col_width=col_width,
            row_height=row_height,
            page_count=page_count,
            output_filename=output_filename,
            pdf_name=pdf_name,
            scale_factor=scale_factor,
            page_breaks=page_breaks,
        )

        # Phase 5: プロンプト生成（すべてを合体させる）
        prompt_path = self.prompt_builder.build(
            md_content=md_content,
            generated_code=generated_code,
            table_summary=table_summary,
            pdf_name=pdf_name,
            output_filename=output_filename,
            page_count=page_count,
        )

        logger.info(f"✅ Successfully processed: {Path(pdf_path).name}")
        return {
            "md_path": md_path,
            "json_path": json_path,
            "prompt_path": prompt_path
        }

    def _compress_json(self, json_data: dict) -> dict:
        """
        トークン節約と処理効率化のため、JSONから冗長な情報を除外する。
        - テキストが空のtext要素を除外
        - 空白のみのテキストを除外
        - style情報が全てnull/Falseの要素のstyleを省略
        """
        compressed = {
            "pdf_name": json_data.get("pdf_name"),
            "page_breaks": json_data.get("page_breaks", []),
            "pages": []
        }

        for page_data in json_data.get("pages", []):
            compressed_page = {
                "page": page_data.get("page"),
                "elements": []
            }

            for elem in page_data.get("elements", []):
                # テキスト要素で空の場合はスキップ
                if elem["type"] == "text":
                    text = elem.get("text", "")
                    if not text or not text.strip():
                        continue

                # スタイル情報の圧縮
                style = elem.get("style", {})
                compressed_style = {}
                if style.get("stroke_width") and style["stroke_width"] > 0:
                    compressed_style["stroke_width"] = style["stroke_width"]

                border = style.get("border", {})
                if any(border.values()):
                    compressed_style["border"] = border

                compressed_elem = {
                    "type": elem["type"],
                    "grid_bbox": elem.get("grid_bbox"),
                }

                if compressed_style:
                    compressed_elem["style"] = compressed_style

                if elem.get("text"):
                    compressed_elem["text"] = elem["text"]

                if elem.get("font_size"):
                    compressed_elem["font_size"] = elem["font_size"]

                compressed_page["elements"].append(compressed_elem)

            compressed["pages"].append(compressed_page)

        return compressed
