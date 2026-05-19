"""スキャンPDF（画像PDF）のOCRによるデータ抽出。

extract_pdf_data() と同じスキーマで返すことで、
以降のパイプライン（grid.py / layout.py / excel.py）をそのまま流用できる。

必要パッケージ（--scan 使用時のみ）:
    pip install rapidocr Pillow
"""

from src.utils.logger import get_logger

logger = get_logger(__name__)

# RapidOCR の信頼スコア下限（0〜1 スケール）
_OCR_CONF_THRESHOLD = 0.3


def extract_scan_pdf_data(pdf_path: str, dpi: int = 300) -> dict:
    """スキャンPDFをOCRで解析し、extract_pdf_data() と同じ構造で返す。

    Args:
        pdf_path: 対象PDFパス
        dpi: ラスタライズ解像度（300推奨。低いとOCR精度が落ちる）

    Returns:
        extract_pdf_data() と互換のページデータ辞書
    """
    _check_dependencies()

    import pypdfium2 as pdfium  # noqa: PLC0415
    from rapidocr import RapidOCR  # noqa: PLC0415
    from rapidocr.utils.typings import LangRec  # noqa: PLC0415

    ocr = RapidOCR(params={"Rec.lang_type": LangRec.JAPAN})

    pdf = pdfium.PdfDocument(pdf_path)
    pages = []
    scale = dpi / 72  # PDF pt (72pt/inch 基準) → pixel への拡大率

    for page_num in range(len(pdf)):
        page = pdf[page_num]
        logger.info(f"[scan] ページ {page_num + 1}/{len(pdf)} OCR処理中...")

        width = page.get_width()    # PDF pt 単位
        height = page.get_height()  # PDF pt 単位

        bitmap = page.render(scale=scale)
        img = bitmap.to_pil()
        words = _run_ocr(img, scale, ocr)

        pages.append({
            'page_number': page_num + 1,
            'width': width,
            'height': height,
            'words': words,
            # テキストのみモード: 罫線・テーブル情報はなし
            'table_bboxes': [],
            'table_col_x_positions': [],
            'table_row_y_positions': [],
            'table_cells': [],
            'table_data': [],
            'table_data_raw': [],
            'rects': [],
            'h_edges': [],
            'v_edges': [],
        })

    pdf.close()
    total_words = sum(len(p['words']) for p in pages)
    logger.info(f"[scan] {len(pages)} ページ完了（抽出語数: {total_words}）")
    return {'pages': pages}


def _run_ocr(img, scale: float, ocr) -> list:
    """PIL Image を RapidOCR で解析し、pdfplumber 互換の words リストに変換する。"""
    import numpy as np  # noqa: PLC0415

    pt_per_pixel = 1.0 / scale  # pixel → PDF pt

    result = ocr(np.array(img))
    if result is None or result.boxes is None:
        return []

    words = []
    for box, text, score in zip(result.boxes, result.txts, result.scores):
        if not text or not text.strip():
            continue
        if score < _OCR_CONF_THRESHOLD:
            continue
        xs = [float(p[0]) for p in box]
        ys = [float(p[1]) for p in box]
        words.append({
            'text': text,
            'x0':     min(xs) * pt_per_pixel,
            'top':    min(ys) * pt_per_pixel,
            'x1':     max(xs) * pt_per_pixel,
            'bottom': max(ys) * pt_per_pixel,
            'fontname': '',
        })
    return words


def _check_dependencies() -> None:
    errors = []
    try:
        import pypdfium2  # noqa: F401, PLC0415
    except ImportError:
        errors.append("  pypdfium2 が未インストール → pip install pypdfium2")
    try:
        import rapidocr  # noqa: F401, PLC0415
    except ImportError:
        errors.append("  rapidocr が未インストール → pip install rapidocr")
    try:
        from PIL import Image  # noqa: F401, PLC0415
    except ImportError:
        errors.append("  Pillow が未インストール → pip install Pillow")
    if errors:
        raise ImportError(
            "--scan モードに必要なパッケージが不足しています:\n" + "\n".join(errors)
        )
