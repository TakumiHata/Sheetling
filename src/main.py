"""
Sheetling エントリポイント。
Phase 1: PDF → 解析 → プロンプト生成（常に実行）
Phase 3: AI出力Pythonソース → 実行 → 3シートExcel生成
"""

from pathlib import Path
from src.core.pipeline import SheetlingPipeline
from src.utils.logger import get_logger

logger = get_logger(__name__)


def run_pipeline():
    """
    Sheetling パイプライン:
    - Phase 1 は常に実行（最新の抽出データ・画像を保証）
    - *_gen.py が存在する場合 → Phase 3 も実行（Excel生成）
    """
    logger.info("=" * 60)
    logger.info("Starting Sheetling pipeline...")
    logger.info("=" * 60)

    input_dir = Path("data/in")
    output_base_dir = Path("data/out")

    pipeline = SheetlingPipeline(str(output_base_dir))

    # 入力PDFの走査
    pdf_files = list(input_dir.glob("*.pdf"))
    if not pdf_files:
        logger.warning(f"No PDF files found in {input_dir}. Please place PDF files to process.")
        return

    for pdf_path in pdf_files:
        pdf_name = pdf_path.stem
        out_dir = output_base_dir / pdf_name
        gen_py_path = out_dir / f"{pdf_name}_gen.py"

        # Phase 1: 常に実行（抽出データ・画像を最新化）
        try:
            pipeline.generate_prompts(str(pdf_path))
        except Exception as e:
            logger.error(f"❌ Phase 1 failed for {pdf_path.name}: {e}", exc_info=True)

        # Phase 3: AI出力のPythonソースが存在する場合 → Excel生成
        if gen_py_path.exists():
            try:
                pipeline.render_excel(pdf_name, str(gen_py_path))
            except Exception as e:
                logger.error(f"❌ Phase 3 failed for {pdf_name}: {e}", exc_info=True)

    logger.info("=" * 60)
    logger.info("Pipeline complete.")
    logger.info("=" * 60)


if __name__ == "__main__":
    run_pipeline()
