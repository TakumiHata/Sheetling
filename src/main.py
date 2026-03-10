import argparse
from pathlib import Path
from src.core.pipeline import SheetlingPipeline
from src.utils.logger import get_logger

logger = get_logger(__name__)

def main():
    parser = argparse.ArgumentParser(description="Sheetling: PDF to Excel conversion")
    parser.add_argument("phase", choices=["extract", "generate"], help="Phase to run: extract (Phase 1) or generate (Phase 3)")
    parser.add_argument("--pdf", type=str, help="PDF file path for extraction (Phase 1). If not provided, processes all PDFs in data/in/")
    args = parser.parse_args()

    pipeline = SheetlingPipeline("data/out")

    if args.phase == "extract":
        if args.pdf:
            pdf_files = [Path(args.pdf)]
        else:
            pdf_files = list(Path("data/in").rglob("*.pdf"))
            
        if not pdf_files:
            logger.warning("No PDF files found in data/in. Please place PDF files to process.")
            return

        for pdf_path in pdf_files:
            try:
                pipeline.generate_prompts(str(pdf_path))
            except Exception as e:
                logger.error(f"❌ Phase 1 failed for {pdf_path.name}: {e}", exc_info=True)

    elif args.phase == "generate":
        output_base_dir = Path("data/out")
        target_dirs = [d for d in output_base_dir.iterdir() if d.is_dir()]
        
        generated_count = 0
        for out_dir in target_dirs:
            pdf_name = out_dir.name
            generated_code_path = out_dir / f"{pdf_name}_gen.py"
            
            # 手動で保存された生成コードがある場合のみ実行
            if generated_code_path.exists():
                generated_count += 1
                try:
                    pipeline.render_excel(pdf_name)
                except Exception as e:
                    logger.error(f"❌ Phase 3 failed for {pdf_name}: {e}", exc_info=True)
                    
        if generated_count == 0:
            logger.warning(f"No *_gen.py files found in subdirectories of {output_base_dir}. Please paste AI generated code first.")

if __name__ == "__main__":
    main()
