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
        
        # *_gen.py ファイルを検索
        gen_files = list(output_base_dir.rglob("*_gen.py"))
        
        generated_count = 0
        for gen_file in gen_files:
            out_dir = gen_file.parent
            
            # gen_file.name は "{pdf_name}_gen.py" なので、末尾の "_gen.py" (7文字) を除外して pdf_name を取得
            if gen_file.name.endswith("_gen.py"):
                pdf_name = gen_file.name[:-7]
                generated_count += 1
                try:
                    pipeline.render_excel(pdf_name, specific_out_dir=str(out_dir))
                except Exception as e:
                    logger.error(f"❌ Phase 3 failed for {pdf_name}: {e}", exc_info=True)
                    
        if generated_count == 0:
            logger.warning(f"No *_gen.py files found in subdirectories of {output_base_dir}. Please paste AI generated code first.")

if __name__ == "__main__":
    main()
