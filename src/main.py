import argparse
import json
from pathlib import Path
from src.core.pipeline import SheetlingPipeline
from src.utils.logger import get_logger

logger = get_logger(__name__)


def main():
    parser = argparse.ArgumentParser(description="Sheetling: PDF to Excel conversion (High Precision Auto)")
    parser.add_argument(
        "command",
        choices=["auto", "correct"],
        help=(
            "auto: PDF から Excel を自動生成 (高精度 Pre版ロジック採用), "
            "correct: ビジョンLLMの修正指示を適用して Excel を再生成"
        ),
    )
    parser.add_argument(
        "--pdf",
        type=str,
        help="PDF名またはパス。省略時は data/in 内の全PDFを処理、correct では出力フォルダ特定に使用。",
    )
    parser.add_argument(
        "--grid-size",
        type=str,
        default="small",
        help="Grid size (デフォルト: small)。Pre版ロジックでは small が基準となります。",
    )
    args = parser.parse_args()

    pipeline = SheetlingPipeline("data/out")

    if args.command == "auto":
        if args.pdf:
            if Path(args.pdf).exists():
                pdf_files = [Path(args.pdf)]
            else:
                # 拡張子なしの指定に対応
                p = Path("data/in") / (args.pdf if args.pdf.endswith(".pdf") else f"{args.pdf}.pdf")
                if p.exists():
                    pdf_files = [p]
                else:
                    # フォルダ検索
                    pdf_files = list(Path("data/in").rglob(f"*{args.pdf}*.pdf"))
        else:
            pdf_files = list(Path("data/in").rglob("*.pdf"))

        if not pdf_files:
            logger.warning("処理対象の PDF ファイルが見つかりません。")
            return

        for pdf_path in pdf_files:
            for _gs in ("1pt", "2pt"):
                try:
                    pipeline.auto_layout(str(pdf_path), grid_size=_gs)
                except Exception as e:
                    logger.error(f"❌ auto ({_gs}) failed for {pdf_path.name}: {e}", exc_info=True)

    elif args.command == "correct":
        from src.templates.prompts import GRID_SIZES
        all_grid_sizes = list(GRID_SIZES.keys())

        output_base_dir = Path("data/out")
        in_base_dir = Path("data/in")
        if args.pdf:
            pdf_path_obj = Path(args.pdf)
            pdf_stem = pdf_path_obj.stem
            # auto と同じロジック: data/in/ からの相対パスで out_dir を決定
            try:
                rel = pdf_path_obj.parent.relative_to(in_base_dir)
                out_dirs = [output_base_dir / rel]
            except ValueError:
                # パスが data/in/ 配下でない場合は layout ファイルを持つディレクトリを探索
                candidate_dirs = [
                    d for d in output_base_dir.rglob("*")
                    if d.is_dir() and any(d.glob(f"{pdf_stem}_*_layout.json"))
                ]
                out_dirs = candidate_dirs if candidate_dirs else [output_base_dir / pdf_stem]
        else:
            # 修正ファイルが存在するディレクトリを自動探索
            # 構造: out_dir/prompts/{grid_size}/page_N/corrections.json（新）
            #       out_dir/prompts/page_N/corrections.json（旧）
            out_dir_set = set()
            for p in output_base_dir.rglob("*_visual_corrections*.json"):
                if p.parent.name.startswith("page_"):
                    if p.parent.parent.name == "prompts":
                        # 旧構造: out_dir/prompts/page_N/
                        out_dir_set.add(p.parent.parent.parent)
                    else:
                        # 新構造: out_dir/prompts/{grid_size}/page_N/
                        out_dir_set.add(p.parent.parent.parent.parent)
                else:
                    out_dir_set.add(p.parent.parent)
            out_dirs = sorted(out_dir_set)

        if not out_dirs:
            logger.warning("修正ファイル (*_visual_corrections*.json) が見つかりませんでした。")
            return

        for out_dir in out_dirs:
            if not out_dir.exists(): continue

            layout_files = list(out_dir.glob("*_layout.json"))
            if not layout_files: continue

            # layout ファイルから (pdf_name, grid_size) ペアを検出
            # ファイル名: {pdf_name}_{grid_size}_layout.json → stem: {pdf_name}_{grid_size}_layout
            pairs: list[tuple[str, str]] = []
            for lf in layout_files:
                stem = lf.stem  # e.g. "tirechange_1pt_layout"
                if stem.endswith("_layout"):
                    stem = stem[: -len("_layout")]  # → "tirechange_1pt"
                for gs in all_grid_sizes:
                    if stem.endswith(f"_{gs}"):
                        pairs.append((stem[: -len(f"_{gs}")], gs))
                        break

            if not pairs:
                continue

            for pdf_name, grid_size in pairs:
                layout_json_name = f"{pdf_name}_{grid_size}_layout.json"
                grid_params_name = f"{pdf_name}_{grid_size}_grid_params.json"

                try:
                    # 修正ファイルの収集（新構造: prompts/{grid_size}/page_N/）
                    prompts_dir = out_dir / "prompts" / grid_size
                    page_files = sorted(prompts_dir.glob("page_*/*_visual_corrections_page*.json"))
                    if not page_files:
                        page_files = sorted(prompts_dir.glob("*_visual_corrections_page*.json"))

                    if page_files:
                        merged = []
                        for pf in page_files:
                            data = json.loads(pf.read_text(encoding="utf-8"))
                            merged.extend(data.get("corrections", []))
                        corrections_json = json.dumps({"corrections": merged}, ensure_ascii=False)

                        pipeline.apply_corrections(
                            pdf_name, corrections_json,
                            specific_out_dir=str(out_dir),
                            layout_json_name=layout_json_name,
                            grid_params_name=grid_params_name
                        )
                        pipeline.rerender_after_corrections(
                            pdf_name, grid_size=grid_size,
                            specific_out_dir=str(out_dir),
                        )
                        logger.info(f"✅ correct 完了: {pdf_name} ({out_dir.name}, {grid_size})")
                    else:
                        logger.warning(f"⚠️ {pdf_name} ({grid_size}): 修正ファイルが見つかりません: {prompts_dir}")

                except Exception as e:
                    logger.error(f"❌ correct failed for {pdf_name} ({grid_size}): {e}", exc_info=True)


if __name__ == "__main__":
    main()
