import argparse
from pathlib import Path
from src.core.pipeline import SheetlingPipeline
from src.utils.logger import get_logger

logger = get_logger(__name__)

def main():
    parser = argparse.ArgumentParser(description="Sheetling: PDF to Excel conversion")
    parser.add_argument("phase", choices=["extract", "fill", "generate"],
                        help=(
                            "Phase to run: "
                            "extract (Phase 1: PDF解析 & プロンプト生成), "
                            "fill (Phase 1.5後処理: STEP 1.5出力のテキスト補完 & STEP 2プロンプト更新), "
                            "generate (Phase 3: 生成コードを実行してExcel出力)"
                        ))
    parser.add_argument("--pdf", type=str, help="PDF名（拡張子なし）または PDFファイルパス。fill/generate では出力ディレクトリの特定に使用。")
    parser.add_argument("--grid-size", type=str, choices=["small", "medium", "large"], default="small", help="Grid size for Excel layout (small, medium, large)")
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
                pipeline.generate_prompts(str(pdf_path), grid_size=args.grid_size)
            except Exception as e:
                logger.error(f"❌ Phase 1 failed for {pdf_path.name}: {e}", exc_info=True)

    elif args.phase == "fill":
        # STEP 1.5 出力のテキスト補完 & STEP 2 プロンプト更新
        # ユーザーは STEP 1.5 の LLM 出力を
        #   data/out/{pdf_name}/prompts/{pdf_name}_step1_5_input.json
        # に貼り付けてからこのコマンドを実行する（extract 時に自動生成済み）。
        output_base_dir = Path("data/out")

        if args.pdf:
            pdf_name = Path(args.pdf).stem
            out_dir = output_base_dir / pdf_name
            input_files = list((out_dir / "prompts").glob(f"{pdf_name}_step1_5_input.json"))
        else:
            input_files = list(output_base_dir.rglob("*_step1_5_input.json"))

        if not input_files:
            logger.warning(
                "STEP 1.5 の入力ファイルが見つかりません。\n"
                "  STEP 1.5 の LLM 出力 JSON を以下のパスに貼り付けてから再実行してください:\n"
                "  data/out/{pdf_name}/prompts/{pdf_name}_step1_5_input.json"
            )
            return

        filled_count = 0
        for input_file in input_files:
            pdf_name = input_file.name.replace("_step1_5_input.json", "")
            out_dir = input_file.parent.parent  # prompts/ の親
            try:
                step1_5_json = input_file.read_text(encoding="utf-8").strip()
                pipeline.fill_layout(pdf_name, step1_5_json, specific_out_dir=str(out_dir))
                filled_count += 1
                logger.info(
                    f"✅ fill 完了: {pdf_name}\n"
                    f"  補完済みJSON: {input_file.parent / f'{pdf_name}_step1_5_output.json'}\n"
                    f"  STEP 2 プロンプトを更新しました: {input_file.parent / f'{pdf_name}_prompt_step2.txt'}\n"
                    f"  ※ 次のステップ: STEP 2 プロンプトを AI チャットに貼り付けて Python コードを生成してください"
                )
            except Exception as e:
                logger.error(f"❌ fill failed for {pdf_name}: {e}", exc_info=True)

        if filled_count == 0:
            logger.warning("fill を実行できたファイルがありませんでした。")

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

                # fill コマンドが実行済みか確認（必須）
                fill_output = out_dir / "prompts" / f"{pdf_name}_step1_5_output.json"
                if not fill_output.exists():
                    logger.error(
                        f"❌ generate をスキップ: {pdf_name}\n"
                        f"  fill コマンドが未実行です。先に以下を実行してください:\n"
                        f"  1. STEP 1.5 の LLM 出力を data/out/{pdf_name}/input/input.json に保存\n"
                        f"  2. python -m src.main fill --pdf {pdf_name}"
                    )
                    continue

                generated_count += 1
                try:
                    pipeline.render_excel(pdf_name, specific_out_dir=str(out_dir))
                except Exception as e:
                    logger.error(f"❌ Phase 3 failed for {pdf_name}: {e}", exc_info=True)

        if generated_count == 0:
            logger.warning(f"No *_gen.py files found in subdirectories of {output_base_dir}. Please paste AI generated code first.")

if __name__ == "__main__":
    main()
