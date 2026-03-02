import os
from pathlib import Path
from src.core.pipeline import SheetlingPipeline
from src.utils.logger import get_logger

logger = get_logger(__name__)


def run_pipeline():
    """
    Sheetling パイプライン:
    PDF → [MarkItDown + pdfplumber + Docling] → 統合 → MD + JSON + Prompt 出力
    """
    logger.info("=" * 60)
    logger.info("Starting Sheetling pipeline...")
    logger.info("=" * 60)

    # パス設定
    input_dir = Path("data/in")
    output_base_dir = "data/out"

    # インスタンス生成
    pipeline = SheetlingPipeline(output_base_dir)

    # PDFファイルの走査
    pdf_files = list(input_dir.glob("*.pdf"))
    if not pdf_files:
        logger.warning(f"No PDF files found in {input_dir}. Please place PDF files to process.")
        return

    logger.info(f"Found {len(pdf_files)} PDF file(s) to process.")

    for pdf_path in pdf_files:
        try:
            result = pipeline.run(str(pdf_path))
            logger.info("")

        except Exception as e:
            logger.error(f"❌ Failed to process {pdf_path.name}: {e}", exc_info=True)

    logger.info("=" * 60)
    logger.info("Pipeline complete.")
    logger.info("=" * 60)


if __name__ == "__main__":
    run_pipeline()
