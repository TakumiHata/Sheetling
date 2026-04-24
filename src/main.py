import argparse
import csv
import json
from pathlib import Path
from src.core.pipeline import SheetlingPipeline
from src.utils.logger import get_logger

logger = get_logger(__name__)


def _resolve_pdf_files(pdf_arg: str | None) -> list:
    if pdf_arg:
        if Path(pdf_arg).exists():
            return [Path(pdf_arg)]
        p = Path("data/in") / (pdf_arg if pdf_arg.endswith(".pdf") else f"{pdf_arg}.pdf")
        if p.exists():
            return [p]
        return list(Path("data/in").rglob(f"*{pdf_arg}*.pdf"))
    return list(Path("data/in").rglob("*.pdf"))


def _run_auto(args, pipeline):
    pdf_files = _resolve_pdf_files(args.pdf)
    if not pdf_files:
        logger.warning("処理対象の PDF ファイルが見つかりません。")
        return
    for pdf_path in pdf_files:
        for gs in ("1pt", "2pt"):
            try:
                pipeline.auto_layout(str(pdf_path), grid_size=gs)
            except Exception as e:
                logger.error(f"❌ auto ({gs}) failed for {pdf_path.name}: {e}", exc_info=True)


def _find_correction_out_dirs(args):
    from src.core.grid_config import GRID_SIZES
    output_base_dir = Path("data/out")
    in_base_dir = Path("data/in")

    if args.pdf:
        pdf_path_obj = Path(args.pdf)
        pdf_stem = pdf_path_obj.stem
        try:
            rel = pdf_path_obj.parent.relative_to(in_base_dir)
            return [output_base_dir / rel]
        except ValueError:
            candidate_dirs = [
                d for d in output_base_dir.rglob("*")
                if d.is_dir() and any(d.glob(f"{pdf_stem}_*_layout.json"))
            ]
            return candidate_dirs if candidate_dirs else [output_base_dir / pdf_stem]

    out_dir_set = set()
    for p in output_base_dir.rglob("*_visual_corrections*.json"):
        if p.parent.name.startswith("page_"):
            if p.parent.parent.name == "prompts":
                out_dir_set.add(p.parent.parent.parent)
            else:
                out_dir_set.add(p.parent.parent.parent.parent)
        else:
            out_dir_set.add(p.parent.parent)
    return sorted(out_dir_set)


def _detect_layout_pairs(out_dir):
    from src.core.grid_config import GRID_SIZES
    all_grid_sizes = list(GRID_SIZES.keys())
    pairs = []
    for lf in out_dir.glob("*_layout.json"):
        stem = lf.stem
        if stem.endswith("_layout"):
            stem = stem[: -len("_layout")]
        for gs in all_grid_sizes:
            if stem.endswith(f"_{gs}"):
                pairs.append((stem[: -len(f"_{gs}")], gs))
                break
    return pairs


def _apply_corrections_for_pair(pipeline, out_dir, pdf_name, grid_size):
    layout_json_name = f"{pdf_name}_{grid_size}_layout.json"
    prompts_dir = out_dir / "prompts" / grid_size
    page_files = sorted(prompts_dir.glob("page_*/*_visual_corrections_page*.json"))
    if not page_files:
        page_files = sorted(prompts_dir.glob("*_visual_corrections_page*.json"))

    if not page_files:
        logger.warning(f"⚠️ {pdf_name} ({grid_size}): 修正ファイルが見つかりません: {prompts_dir}")
        return

    merged = []
    for pf in page_files:
        data = json.loads(pf.read_text(encoding="utf-8"))
        merged.extend(data.get("corrections", []))

    corrections_json = json.dumps({"corrections": merged}, ensure_ascii=False)
    pipeline.apply_corrections(
        pdf_name, corrections_json,
        specific_out_dir=str(out_dir), layout_json_name=layout_json_name)
    pipeline.rerender_after_corrections(
        pdf_name, grid_size=grid_size, specific_out_dir=str(out_dir))
    logger.info(f"✅ correct 完了: {pdf_name} ({out_dir.name}, {grid_size})")


def _run_correct(args, pipeline):
    out_dirs = _find_correction_out_dirs(args)
    if not out_dirs:
        logger.warning("修正ファイル (*_visual_corrections*.json) が見つかりませんでした。")
        return

    for out_dir in out_dirs:
        if not out_dir.exists():
            continue
        pairs = _detect_layout_pairs(out_dir)
        for pdf_name, grid_size in pairs:
            try:
                _apply_corrections_for_pair(pipeline, out_dir, pdf_name, grid_size)
            except Exception as e:
                logger.error(f"❌ correct failed for {pdf_name} ({grid_size}): {e}", exc_info=True)


def _run_check():
    import pdfplumber

    in_base_dir = Path("data/in")
    doc_dir = Path("data/doc")
    doc_dir.mkdir(parents=True, exist_ok=True)
    output_csv = doc_dir / "pdflist_check.csv"

    pdf_files = sorted(in_base_dir.rglob("*.pdf"))
    logger.info(f"対象PDFファイル数: {len(pdf_files)} 件")

    results = []
    for pdf_path in pdf_files:
        rel_path = pdf_path.relative_to(in_base_dir)
        logger.info(f"チェック中: {rel_path}")
        try:
            with pdfplumber.open(str(pdf_path)) as pdf:
                page_count = len(pdf.pages)
                status = "スキャンPDF（画像）"
                for page in pdf.pages:
                    text = page.extract_text()
                    if text and text.strip():
                        status = "通常PDF（テキストあり）"
                        break
        except Exception as e:
            logger.error(f"  [ERROR] {e}")
            status, page_count = "エラー", 0
        logger.info(f"  => {status}（{page_count}ページ）")
        results.append({"ファイルパス": str(rel_path), "ページ数": page_count, "判定": status})

    _write_check_results(results, output_csv)


def _write_check_results(results, output_csv):
    with open(output_csv, "w", newline="", encoding="utf_8_sig") as f:
        writer = csv.DictWriter(f, fieldnames=["ファイルパス", "ページ数", "判定"])
        writer.writeheader()
        writer.writerows(results)

    logger.info(f"完了。結果を保存しました: {output_csv}")
    scan_count = sum(1 for r in results if r["判定"] == "スキャンPDF（画像）")
    text_count = sum(1 for r in results if r["判定"] == "通常PDF（テキストあり）")
    error_count = sum(1 for r in results if r["判定"] == "エラー")
    logger.info(f"--- サマリー ---")
    logger.info(f"  通常PDF（テキストあり）: {text_count} 件")
    logger.info(f"  スキャンPDF（画像）    : {scan_count} 件")
    logger.info(f"  エラー                 : {error_count} 件")


def main():
    parser = argparse.ArgumentParser(description="Sheetling: PDF to Excel conversion")
    parser.add_argument(
        "command", choices=["auto", "correct", "check"],
        help="auto: PDF→Excel自動生成, correct: LLM修正適用, check: PDF判定CSV出力",
    )
    parser.add_argument("--pdf", type=str, help="PDF名またはパス。省略時は全PDF処理。")
    args = parser.parse_args()

    pipeline = SheetlingPipeline("data/out")

    if args.command == "auto":
        _run_auto(args, pipeline)
    elif args.command == "correct":
        _run_correct(args, pipeline)
    elif args.command == "check":
        _run_check()


if __name__ == "__main__":
    main()
