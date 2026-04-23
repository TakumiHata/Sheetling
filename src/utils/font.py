import re

from src.core.constants import BORDER_MEDIUM_MAX_LW, BORDER_THIN_MAX_LW


def linewidth_to_border_style(linewidth: float) -> str:
    if linewidth <= BORDER_THIN_MAX_LW:
        return 'thin'
    if linewidth <= BORDER_MEDIUM_MAX_LW:
        return 'medium'
    return 'thick'


def normalize_font_name(raw_name):
    if not raw_name:
        return None
    if isinstance(raw_name, bytes):
        return None
    if isinstance(raw_name, str) and raw_name.startswith("b'"):
        return None
    name = re.sub(r'^[A-Z]{6}\+', '', raw_name)
    if 'Mincho' in name or '明朝' in name:
        return 'MS 明朝'
    return 'MS ゴシック'
