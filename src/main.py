import os
from pathlib import Path
from dotenv import load_dotenv
from src.core.pipeline import SheetlingPipeline
from src.utils.logger import get_logger

# 環境変数の読み込み
load_dotenv()

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
    input_dir = Path("data/01_input_pdf")
    md_dir = "data/02_markdown"
    json_dir = "data/03_layout_json"
    prompt_dir = "data/04_prompt"

    # インスタンス生成
    pipeline = SheetlingPipeline(md_dir, json_dir, prompt_dir)

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
