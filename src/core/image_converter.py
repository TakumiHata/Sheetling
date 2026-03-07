"""
PDFを画像（PNG）に変換するモジュール。
PyMuPDF (fitz) を使用して各ページをPNG画像に変換する。
PDF内蔵フォントを利用するため、システムに日本語フォントがなくても正しく描画される。
"""

from pathlib import Path
import fitz  # PyMuPDF
from src.utils.logger import get_logger

logger = get_logger(__name__)


class ImageConverter:
    """PDFの各ページを画像に変換して保存する"""

    def __init__(self, dpi: int = 200):
        self.dpi = dpi

    def convert(self, pdf_path: str, out_dir: Path) -> list[str]:
        """
        PDFを画像に変換し、out_dirにPNGとして保存する。

        Returns:
            list of image file paths
        """
        logger.info(f"Converting PDF to images: {pdf_path}")
        pdf_name = Path(pdf_path).stem

        # PDFドキュメントを開く
        doc = fitz.open(pdf_path)
        image_paths = []

        # PyMuPDFのデフォルト(72dpi)に対する倍率を計算して解像度を調整する
        zoom = self.dpi / 72.0
        matrix = fitz.Matrix(zoom, zoom)

        for i, page in enumerate(doc):
            # ページを指定した解像度の画像データに変換
            pix = page.get_pixmap(matrix=matrix)
            img_path = out_dir / f"{pdf_name}_page{i + 1}.png"
            pix.save(str(img_path))
            image_paths.append(str(img_path))
            logger.info(f"  Page {i + 1} → {img_path}")

        doc.close()
        logger.info(f"✅ {len(image_paths)} page(s) converted to images")
        return image_paths
