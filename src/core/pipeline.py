"""
Sheetling パイプライン。
Phase1: PDF解析 → プロンプト生成
Phase3: AI出力Pythonソース実行 → 3シートExcel生成
"""

import json
from pathlib import Path

from src.core.extractor import PdfExtractor
from src.core.image_converter import ImageConverter
from src.core.executor import Executor
from src.core.prompts import get_system_prompt
from src.core.config import config
from src.utils.logger import get_logger

logger = get_logger(__name__)


class SheetlingPipeline:
    """
    1. PDF を解析してプロンプトを出力する (Phase 1)。
    2. ユーザーがLLMから得たPythonソースを実行し、3シートExcelを生成する (Phase 3)。
    """

    def __init__(self, output_base_dir: str):
        self.output_base_dir = Path(output_base_dir)
        self.extractor = PdfExtractor()
        self.image_converter = ImageConverter()
        self.executor = Executor()

    def generate_prompts(self, pdf_path: str) -> dict:
        """
        Phase 1: PDFを解析し、LLMに渡すためのプロンプトを data/out/ に出力する。
        """
        logger.info(f"--- [Phase 1] PDF解析 & プロンプト生成: {Path(pdf_path).name} ---")
        pdf_name = Path(pdf_path).stem

        out_dir = self.output_base_dir / pdf_name
        out_dir.mkdir(parents=True, exist_ok=True)

        # pdfplumber でテキスト・座標・フォント・色情報を抽出
        extract_result = self.extractor.extract(pdf_path, out_dir)

        # PDF → 画像変換（3シート目用）
        try:
            image_paths = self.image_converter.convert(pdf_path, out_dir)
        except Exception as e:
            logger.warning(f"PDF→画像変換に失敗（3シート目は空になります）: {e}")
            image_paths = []

        # 画像パスを記録しておく（Phase3で使用）
        meta_path = out_dir / f"{pdf_name}_meta.json"
        meta = {
            "fonts": extract_result["fonts"],
            "colors": extract_result["colors"],
            "image_paths": image_paths,
        }
        with open(meta_path, "w", encoding="utf-8") as f:
            json.dump(meta, f, indent=2, ensure_ascii=False)

        # 抽出JSONを読み込み
        with open(extract_result["json_path"], "r", encoding="utf-8") as f:
            extracted_json = json.load(f)

        # プロンプトを作成して保存
        system_prompt = get_system_prompt()
        prompt_text = (
            f"{system_prompt}\n\n"
            f"=== 以下は {pdf_name} から抽出されたレイアウトデータです ===\n"
            f"```json\n"
            f"{json.dumps(extracted_json, indent=2, ensure_ascii=False)}\n"
            f"```\n"
        )

        prompt_path = out_dir / f"{pdf_name}_prompt.txt"
        with open(prompt_path, "w", encoding="utf-8") as f:
            f.write(prompt_text)

        logger.info(f"✅ Phase 1 完了: {pdf_name}")
        logger.info(f"  プロンプト: {prompt_path}")
        logger.info(f"  ※ プロンプトをLLMに投入し、返されたPythonコードを")
        logger.info(f"    {out_dir / f'{pdf_name}_gen.py'} として保存してください。")

        return {
            "md_path": extract_result["md_path"],
            "json_path": extract_result["json_path"],
            "prompt_path": str(prompt_path),
            "meta_path": str(meta_path),
        }

    def render_excel(self, pdf_name: str, gen_py_path: str) -> str:
        """
        Phase 3: AI出力のPythonソースを実行し、3シートExcelを生成する。
        """
        logger.info(f"--- [Phase 3] Excel生成: {pdf_name} ---")
        out_dir = self.output_base_dir / pdf_name
        out_dir.mkdir(parents=True, exist_ok=True)

        output_xlsx_path = out_dir / f"{pdf_name}.xlsx"

        # メタデータを読み込み（Phase1で保存したフォント・色・画像パス情報）
        meta_path = out_dir / f"{pdf_name}_meta.json"
        if meta_path.exists():
            with open(meta_path, "r", encoding="utf-8") as f:
                meta = json.load(f)
            fonts = meta.get("fonts", [])
            colors = meta.get("colors", [])
            image_paths = meta.get("image_paths", [])
        else:
            logger.warning(f"メタデータが見つかりません: {meta_path}")
            fonts = []
            colors = []
            image_paths = []

        # Executor で3シートExcel生成
        result_path = self.executor.execute(
            gen_py_path=gen_py_path,
            output_xlsx_path=str(output_xlsx_path),
            fonts=fonts,
            colors=colors,
            image_paths=image_paths,
        )

        logger.info(f"✅ Phase 3 完了: {result_path}")
        return result_path
