import argparse
import json
from pathlib import Path
from src.core.pipeline import SheetlingPipeline
from src.utils.logger import get_logger

logger = get_logger(__name__)


def main():
    parser = argparse.ArgumentParser(description="Sheetling: PDF to Excel conversion")
    parser.add_argument(
        "command",
        choices=["auto", "correct"],
        help=(
            "auto: PDF から Excel を自動生成 (PDF解析 → レイアウト生成 → Excel出力), "
            "correct: ビジョンLLMの修正指示を適用して Excel を再生成"
        ),
    )
    parser.add_argument(
        "--pdf",
        type=str,
        help="PDF名（拡張子なし）または PDFファイルパス。correct では出力ディレクトリの特定に使用。",
    )
    parser.add_argument(
        "--grid-size",
        type=str,
        choices=["small", "medium", "large"],
        default="small",
        help="Grid size for Excel layout (small, medium, large)",
    )
    args = parser.parse_args()

    pipeline = SheetlingPipeline("data/out")

    if args.command == "auto":
        if args.pdf:
            pdf_files = [Path(args.pdf)]
        else:
            pdf_files = list(Path("data/in").rglob("*.pdf"))

        if not pdf_files:
            logger.warning("No PDF files found in data/in. Please place PDF files to process.")
            return

        for pdf_path in pdf_files:
            try:
                result = pipeline.auto_layout(str(pdf_path), grid_size=args.grid_size)
                page_imgs = result.get("page_image_paths", [])
                review_paths = result.get("visual_review_paths", [])
                img_lines = "\n".join(f"    {p}" for p in page_imgs)
                prompt_lines = "\n".join(f"    {p}" for p in review_paths)
                logger.info(
                    f"✅ auto 完了: {pdf_path.stem}\n"
                    f"  Excel:         {result['xlsx_path']}\n"
                    f"  PDFページ画像:\n{img_lines}\n"
                    f"  検証プロンプト:\n{prompt_lines}\n"
                    f"  ※ 罫線修正あり なら:\n"
                    f"    1. 各ページの PNG + Excelファイル + プロンプトテキストを社内LLMに投入\n"
                    f"    2. 出力JSONを <pdf_name>_visual_corrections_page{{N}}.json に保存\n"
                    f"    3. python -m src.main correct --pdf {pdf_path.stem}"
                )
            except FileNotFoundError as e:
                logger.error(f"❌ {e}")
            except Exception as e:
                logger.error(f"❌ auto failed for {pdf_path.name}: {e}", exc_info=True)

    elif args.command == "correct":
        output_base_dir = Path("data/out")

        # 処理対象の out_dir 一覧を収集
        if args.pdf:
            out_dirs = [output_base_dir / Path(args.pdf).stem]
        else:
            # corrections ファイルが存在する out_dir をすべて収集
            out_dirs = sorted(set(
                p.parent.parent.parent if p.parent.name.startswith("page_") else p.parent.parent
                for p in output_base_dir.rglob("*_visual_corrections*.json")
            ))

        if not out_dirs:
            logger.warning(
                "修正ファイルが見つかりません。\n"
                "ビジョンLLMの出力JSONを以下のパスに保存してから再実行してください:\n"
                "  data/out/<name>/prompts/page_1/<pdf_name>_visual_corrections_page1.json"
            )
            return

        for out_dir in out_dirs:
            prompts_dir = out_dir / "prompts"
            # _layout.json からPDF名を特定（ディレクトリ名と異なる場合に対応）
            layout_files = list(out_dir.glob("*_layout.json"))
            if not layout_files:
                logger.warning(f"⚠️  _layout.json が見つかりません: {out_dir}")
                continue
            pdf_name = layout_files[0].stem.replace("_layout", "")
            try:
                # ワイルドカードでページ単位ファイルを収集してマージ
                page_files = sorted(prompts_dir.glob("page_*/*_visual_corrections_page*.json"))
                if not page_files:
                    page_files = sorted(prompts_dir.glob("*_visual_corrections_page*.json"))
                single_files = list(prompts_dir.glob("*_visual_corrections.json"))

                if page_files:
                    merged_corrections: list = []
                    for pf in page_files:
                        data = json.loads(pf.read_text(encoding="utf-8"))
                        merged_corrections.extend(data.get("corrections", []))
                    corrections_json = json.dumps({"corrections": merged_corrections}, ensure_ascii=False)
                    logger.info(f"[correct] {len(page_files)} ページ分の修正ファイルをマージしました")
                elif single_files:
                    corrections_json = single_files[0].read_text(encoding="utf-8")
                else:
                    logger.warning(f"⚠️  修正ファイルが見つかりません: {prompts_dir}")
                    continue

                pipeline.apply_corrections(pdf_name, corrections_json, specific_out_dir=str(out_dir))
                pipeline.render_excel(pdf_name, specific_out_dir=str(out_dir), apply_border_post_process=False)
                logger.info(f"✅ correct 完了: {out_dir.name} ({pdf_name})")
            except Exception as e:
                logger.error(f"❌ correct failed for {out_dir.name}: {e}", exc_info=True)


if __name__ == "__main__":
    main()
