import os
from pathlib import Path
from dotenv import load_dotenv
from src.core.hybrid_analyzer import HybridAnalyzer
from src.core.prompt_builder import PromptBuilder
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
    inter_md_dir = "data/02_inter_md"
    inter_json_dir = "data/03_inter_json"
    output_dir = "data/04_output_excel"

    # インスタンス生成
    analyzer = HybridAnalyzer(inter_md_dir, inter_json_dir)
    prompt_builder = PromptBuilder(output_dir)

    # PDFファイルの走査
    pdf_files = list(input_dir.glob("*.pdf"))
    if not pdf_files:
        logger.warning(f"No PDF files found in {input_dir}. Please place PDF files to process.")
        return

    logger.info(f"Found {len(pdf_files)} PDF file(s) to process.")

    for pdf_path in pdf_files:
        try:
            logger.info(f"--- Processing: {pdf_path.name} ---")
            pdf_name = pdf_path.stem

            # Phase 1-4: ハイブリッド解析（MarkItDown + pdfplumber + Docling → 統合JSON）
            result = analyzer.analyze(str(pdf_path))

            # Phase 5: プロンプト生成（MD + JSON → 固定テンプレートに埋め込み）
            prompt_path = prompt_builder.build(
                md_path=result["md_path"],
                json_path=result["json_path"],
                pdf_name=pdf_name,
            )

            logger.info(f"✅ Successfully processed: {pdf_path.name}")
            logger.info(f"   MD:     {result['md_path']}")
            logger.info(f"   JSON:   {result['json_path']}")
            logger.info(f"   Prompt: {prompt_path}")
            logger.info("")

        except Exception as e:
            logger.error(f"❌ Failed to process {pdf_path.name}: {e}", exc_info=True)

    logger.info("=" * 60)
    logger.info("Pipeline complete.")
    logger.info("=" * 60)


if __name__ == "__main__":
    run_pipeline()
