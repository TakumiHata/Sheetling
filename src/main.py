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
        output_base_dir = Path("data/out")
        if args.pdf:
            pdf_stem = Path(args.pdf).stem
            out_dirs = [output_base_dir / pdf_stem]
        else:
            # 修正ファイルが存在するディレクトリを自動探索
            out_dirs = sorted(set(
                p.parent.parent.parent if p.parent.name.startswith("page_") else p.parent.parent
                for p in output_base_dir.rglob("*_visual_corrections*.json")
            ))

        if not out_dirs:
            logger.warning("修正ファイル (*_visual_corrections*.json) が見つかりませんでした。")
            return

        for out_dir in out_dirs:
            if not out_dir.exists(): continue
            
            # ディレクトリ名または layout.json から PDF 名を取得
            layout_files = list(out_dir.glob("*_layout.json"))
            if not layout_files: continue
            
            pdf_name = layout_files[0].stem
            # サフィックス除去
            for s in ["_small", "_medium", "_large", "_pattern_1", "_pattern_2"]:
                if pdf_name.endswith(s):
                    pdf_name = pdf_name[:-len(s)]
                    break
            
            grid_size = args.grid_size
            layout_json_name = f"{pdf_name}_{grid_size}_layout.json"
            grid_params_name = f"{pdf_name}_{grid_size}_grid_params.json"
            
            try:
                # 修正ファイルの収集
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

                    # 再レンダリング（grid_size サフィックス付きファイルを使用）
                    pipeline.rerender_after_corrections(
                        pdf_name, grid_size=grid_size,
                        specific_out_dir=str(out_dir),
                    )
                    logger.info(f"✅ correct 完了: {pdf_name} ({out_dir.name})")
                else:
                    logger.warning(f"⚠️ {pdf_name}: 修正ファイルが見つかりません: {prompts_dir}")
                    
            except Exception as e:
                logger.error(f"❌ correct failed for {pdf_name}: {e}", exc_info=True)


if __name__ == "__main__":
    main()
