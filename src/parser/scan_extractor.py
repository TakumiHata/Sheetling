"""スキャンPDF（画像PDF）のOCRによるデータ抽出。

extract_pdf_data() と同じスキーマで返すことで、
以降のパイプライン（grid.py / layout.py / excel.py）をそのまま流用できる。

必要パッケージ（--scan 使用時のみ）:
    pip install pypdfium2 pytesseract Pillow
    + Tesseract OCR 本体（Windows: UB Mannheim インストーラー）
      https://github.com/UB-Mannheim/tesseract/wiki
    + 日本語モデル: インストール時に "Japanese" / "Japanese (vertical)" を選択
"""

from src.utils.logger import get_logger

logger = get_logger(__name__)

# pytesseract の信頼スコア下限（0〜100）。-1 はTesseract内部の構造エントリで常に除外。
_OCR_CONF_THRESHOLD = 30


def extract_scan_pdf_data(pdf_path: str, dpi: int = 300, lang: str = 'jpn') -> dict:
    """スキャンPDFをOCRで解析し、extract_pdf_data() と同じ構造で返す。

    Args:
        pdf_path: 対象PDFパス
        dpi: ラスタライズ解像度（300推奨。低いとOCR精度が落ちる）
        lang: Tesseract言語コード。縦書きが混在する場合は 'jpn+jpn_vert'

    Returns:
        extract_pdf_data() と互換のページデータ辞書
    """
    _check_dependencies()

    import pypdfium2 as pdfium  # noqa: PLC0415

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
        words = _run_ocr(img, scale, lang)

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


def _run_ocr(img, scale: float, lang: str) -> list:
    """PIL Image を pytesseract でOCRし、pdfplumber互換の words リストに変換する。"""
    import pytesseract  # noqa: PLC0415

    pt_per_pixel = 1.0 / scale  # pixel → PDF pt

    data = pytesseract.image_to_data(img, lang=lang, output_type=pytesseract.Output.DICT)

    words = []
    for i, text in enumerate(data['text']):
        if not text or not text.strip():
            continue
        conf = float(data['conf'][i])
        if conf < _OCR_CONF_THRESHOLD:
            continue
        words.append({
            'text': text,
            'x0':     data['left'][i] * pt_per_pixel,
            'top':    data['top'][i]  * pt_per_pixel,
            'x1':    (data['left'][i] + data['width'][i])  * pt_per_pixel,
            'bottom': (data['top'][i]  + data['height'][i]) * pt_per_pixel,
            'fontname': '',
            'font_size': None,
        })
    return words


def _check_dependencies() -> None:
    errors = []
    try:
        import pypdfium2  # noqa: F401, PLC0415
    except ImportError:
        errors.append("  pypdfium2 が未インストール → pip install pypdfium2")
    try:
        import pytesseract  # noqa: F401, PLC0415
    except ImportError:
        errors.append("  pytesseract が未インストール → pip install pytesseract")
    try:
        from PIL import Image  # noqa: F401, PLC0415
    except ImportError:
        errors.append("  Pillow が未インストール → pip install Pillow")
    if errors:
        raise ImportError(
            "--scan モードに必要なパッケージが不足しています:\n" + "\n".join(errors) +
            "\nまた Tesseract OCR 本体と日本語モデル（jpn）のインストールも必要です。"
        )
