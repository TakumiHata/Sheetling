"""
Sheetling エントリポイント。
Phase 3: AI出力Pythonソース実行 → Excel生成
"""

from pathlib import Path
from src.core.pipeline import SheetlingPipeline
from src.utils.logger import get_logger

logger = get_logger(__name__)


def run_generate():
    """
    Sheetling Phase 3 (Excel作成)
    - data/out 以下のフォルダを探索し、*_gen.py があれば pipeline.render_excel() を実行する。
    """
    logger.info("=" * 60)
    logger.info("Starting Sheetling Generation (Phase 3)...")
    logger.info("=" * 60)

    output_base_dir = Path("data/out")

    # パイプラインを初期化
    pipeline = SheetlingPipeline(str(output_base_dir))

    # data/out 以下のフォルダをスキャンし、_gen.py があれば実行する
    target_dirs = [d for d in output_base_dir.iterdir() if d.is_dir()]
    if not target_dirs:
        logger.warning(f"No target directories found in {output_base_dir}.")
        return

    generated_count = 0
    for out_dir in target_dirs:
        pdf_name = out_dir.name
        gen_py_path = out_dir / f"{pdf_name}_gen.py"

        if gen_py_path.exists():
            generated_count += 1
            try:
                pipeline.render_excel(pdf_name, str(gen_py_path))
            except Exception as e:
                logger.error(f"❌ Phase 3 failed for {pdf_name}: {e}", exc_info=True)

    if generated_count == 0:
        logger.warning(f"No *_gen.py files found in any subdirectories of {output_base_dir}.")

    logger.info("=" * 60)
    logger.info("Generation complete.")
    logger.info("=" * 60)


if __name__ == "__main__":
    run_generate()
