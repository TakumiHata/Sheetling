import os
from pathlib import Path
from dotenv import load_dotenv
from src.core.docling_parser import DoclingParser
from src.core.llm_excel_gen import LLMExcelGenerator
from src.utils.logger import get_logger

# 環境変数の読み込み
load_dotenv()

logger = get_logger(__name__)

def run_pipeline():
    logger.info("Starting Sheetling pipeline...")
    
    # パス設定
    input_dir = Path("data/01_input_pdf")
    inter_md_dir = "data/02_inter_md"
    inter_json_dir = "data/03_inter_json"
    output_excel_dir = "data/04_output_excel"
    
    # インスタンス生成
    parser = DoclingParser(inter_md_dir, inter_json_dir)
    generator = LLMExcelGenerator(output_excel_dir)
    
    # PDFファイルの走査
    pdf_files = list(input_dir.glob("*.pdf"))
    if not pdf_files:
        logger.warning(f"No PDF files found in {input_dir}. Please place PDF files to process.")
        return

    for pdf_path in pdf_files:
        try:
            # 1. Doclingによる解析
            md_path, json_path = parser.parse(str(pdf_path))
            
            # 2. LLMによるExcel生成
            excel_path = generator.generate(md_path, json_path)
            
            logger.info(f"Successfully processed: {pdf_path.name} -> {Path(excel_path).name}")
        except Exception as e:
            logger.error(f"Failed to process {pdf_path.name}: {e}", exc_info=True)

if __name__ == "__main__":
    run_pipeline()
