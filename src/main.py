import argparse
from pathlib import Path
from src.core.pipeline import SheetlingPipeline
from src.utils.logger import get_logger

logger = get_logger(__name__)

def main():
    parser = argparse.ArgumentParser(description="Sheetling: PDF to Excel conversion")
    parser.add_argument("phase", choices=["extract", "auto", "fill", "correct", "generate"],
                        help=(
                            "Phase to run: "
                            "extract (Phase 1: PDF解析 & プロンプト生成), "
                            "auto (step1+step1.5+fill+step2を自動化: レイアウトJSON生成・_gen.py生成・視覚的検証プロンプト出力), "
                            "fill (Phase 1.5後処理: STEP 1.5出力のテキスト補完 & STEP 2プロンプト更新), "
                            "correct (ビジョンLLMの修正指示を適用して _gen.py を再生成), "
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

    elif args.phase == "auto":
        # step1 + step1.5 + fill + step2 をスクリプトで完全自動化
        if args.pdf:
            pdf_files = [Path(args.pdf)]
        else:
            pdf_files = list(Path("data/in").rglob("*.pdf"))

        if not pdf_files:
            logger.warning("No PDF files found in data/in.")
            return

        for pdf_path in pdf_files:
            try:
                result = pipeline.auto_layout(str(pdf_path), grid_size=args.grid_size)
                logger.info(
                    f"✅ auto 完了: {pdf_path.stem}\n"
                    f"  _gen.py: {result['gen_py_path']}\n"
                    f"  視覚的検証プロンプト:\n"
                    + "\n".join(f"    {p}" for p in result['visual_review_paths'])
                    + "\n  ※ 修正不要なら: python -m src.main generate"
                    + "\n  ※ 修正あり  なら: _visual_corrections.json に保存後 python -m src.main correct --pdf "
                    + pdf_path.stem
                )
            except FileNotFoundError as e:
                logger.error(f"❌ {e}")
            except Exception as e:
                logger.error(f"❌ auto failed for {pdf_path.name}: {e}", exc_info=True)

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

    elif args.phase == "correct":
        # ビジョンLLMの修正指示を適用して _gen.py を再生成
        output_base_dir = Path("data/out")

        if args.pdf:
            pdf_name = Path(args.pdf).stem
            out_dir = output_base_dir / pdf_name
            corrections_files = list((out_dir / "prompts").glob(f"{pdf_name}_visual_corrections.json"))
            if not corrections_files:
                corrections_files = [out_dir / "prompts" / f"{pdf_name}_visual_corrections.json"]
        else:
            corrections_files = list(output_base_dir.rglob("*_visual_corrections.json"))

        if not corrections_files:
            logger.warning(
                "修正ファイルが見つかりません。\n"
                "ビジョンLLMの出力JSONを以下のパスに保存してから再実行してください:\n"
                "  data/out/<pdf_name>/prompts/<pdf_name>_visual_corrections.json"
            )
            return

        for corrections_file in corrections_files:
            if not corrections_file.exists():
                logger.warning(f"⚠️  ファイルが存在しません: {corrections_file}")
                continue
            pdf_name = corrections_file.name.replace("_visual_corrections.json", "")
            out_dir = corrections_file.parent.parent
            try:
                corrections_json = corrections_file.read_text(encoding="utf-8")
                pipeline.apply_corrections(pdf_name, corrections_json, specific_out_dir=str(out_dir))
                logger.info(f"✅ correct 完了: {pdf_name} → 次は `generate` を実行してください")
            except Exception as e:
                logger.error(f"❌ correct failed for {pdf_name}: {e}", exc_info=True)

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
