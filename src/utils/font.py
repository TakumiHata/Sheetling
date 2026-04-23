import re


def linewidth_to_border_style(linewidth: float) -> str:
    if linewidth <= 1.0:
        return 'thin'
    if linewidth <= 2.0:
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
        return 'MS明朝'
    return 'MSゴシック'
