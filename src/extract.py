"""
Sheetling エントリポイント。
Phase 1: PDFデータ抽出 → プロンプト出力
"""

from pathlib import Path
from src.core.pipeline import SheetlingPipeline
from src.utils.logger import get_logger

logger = get_logger(__name__)


def run_extract():
    """
    Sheetling Phase 1 (抽出)
    - data/in のPDFを読み込み、pipeline.generate_prompts() を実行する。
    """
    logger.info("=" * 60)
    logger.info("Starting Sheetling Extraction (Phase 1)...")
    logger.info("=" * 60)

    input_dir = Path("data/in")
    output_base_dir = Path("data/out")

    # 指定ディレクトリ内のPDFファイルをサブフォルダ含めてすべて取得
    pdf_files = list(input_dir.rglob("*.pdf"))
    if not pdf_files:
        logger.warning(f"No PDF files found in {input_dir}. Please place PDF files to process.")
        return

    # パイプラインを初期化
    pipeline = SheetlingPipeline(str(output_base_dir))

    for pdf_path in pdf_files:
        try:
            # 入力ディレクトリからの相対パス（サブフォルダ構造）を取得
            # 例: data/in/sub/test.pdf -> sub
            #     data/in/test.pdf -> .
            rel_dir = pdf_path.parent.relative_to(input_dir)
            
            pipeline.generate_prompts(str(pdf_path), rel_dir)
        except Exception as e:
            logger.error(f"❌ Phase 1 failed for {pdf_path.name}: {e}", exc_info=True)

    logger.info("=" * 60)
    logger.info("Extraction complete.")
    logger.info("=" * 60)


if __name__ == "__main__":
    run_extract()
