import argparse
import csv
from pathlib import Path
from src.core.auto_layout_service import AutoLayoutService
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


def _run_auto(args, service):
    pdf_files = _resolve_pdf_files(args.pdf)
    if not pdf_files:
        logger.warning("処理対象の PDF ファイルが見つかりません。")
        return
    scan = getattr(args, 'scan', False)
    for pdf_path in pdf_files:
        for gs in ("1pt", "2pt"):
            try:
                service.run(str(pdf_path), grid_size=gs, scan=scan)
            except Exception as e:
                logger.error(f"❌ auto ({gs}) failed for {pdf_path.name}: {e}", exc_info=True)


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
        "command", choices=["auto", "check"],
        help="auto: PDF→Excel自動生成, check: PDF判定CSV出力",
    )
    parser.add_argument("--pdf", type=str, help="PDF名またはパス。省略時は全PDF処理。")
    parser.add_argument(
        "--scan", action="store_true",
        help="スキャンPDF（画像PDF）をOCRで処理する。pymupdf・pytesseract・Tesseract本体が必要。",
    )
    args = parser.parse_args()

    service = AutoLayoutService("data/out")

    if args.command == "auto":
        _run_auto(args, service)
    elif args.command == "check":
        _run_check()


if __name__ == "__main__":
    main()
