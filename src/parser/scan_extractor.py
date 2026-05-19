"""スキャンPDF（画像PDF）のOCRによるデータ抽出。

extract_pdf_data() と同じスキーマで返すことで、
以降のパイプライン（grid.py / layout.py / excel.py）をそのまま流用できる。

必要パッケージ（--scan 使用時のみ）:
    pip install pymupdf pytesseract
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

    import fitz  # noqa: PLC0415

    doc = fitz.open(pdf_path)
    pages = []
    zoom = dpi / 72  # PDF pt (72dpi 基準) → pixel への拡大率

    for page_num, page in enumerate(doc, 1):
        logger.info(f"[scan] ページ {page_num}/{len(doc)} OCR処理中...")
        width = page.rect.width
        height = page.rect.height

        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        words = _run_ocr(pix, zoom, lang)

        pages.append({
            'page_number': page_num,
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

    doc.close()
    total_words = sum(len(p['words']) for p in pages)
    logger.info(f"[scan] {len(pages)} ページ完了（抽出語数: {total_words}）")
    return {'pages': pages}


def _run_ocr(pix, zoom: float, lang: str) -> list:
    """PyMuPDF の Pixmap を pytesseract でOCRし、pdfplumber互換の words リストに変換する。"""
    import io
    import pytesseract  # noqa: PLC0415
    from PIL import Image  # noqa: PLC0415

    scale = 1.0 / zoom  # pixel → PDF pt

    img = Image.open(io.BytesIO(pix.tobytes("png")))
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
            'x0':    data['left'][i] * scale,
            'top':   data['top'][i]  * scale,
            'x1':   (data['left'][i] + data['width'][i])  * scale,
            'bottom': (data['top'][i] + data['height'][i]) * scale,
            'fontname': '',
            'font_size': None,
        })
    return words


def _check_dependencies() -> None:
    errors = []
    try:
        import fitz  # noqa: F401, PLC0415
    except ImportError:
        errors.append("  pymupdf が未インストール → pip install pymupdf")
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
