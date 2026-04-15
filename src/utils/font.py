import re


FONT_ALIASES: dict = {
    'MS Gothic':     'MS Gothic',
    'MSGothic':      'MS Gothic',
    'MS PGothic':    'MS PGothic',
    'MSPGothic':     'MS PGothic',
    'MS Mincho':     'MS Mincho',
    'MSMincho':      'MS Mincho',
    'MS PMincho':    'MS PMincho',
    'MSPMincho':     'MS PMincho',
    'MS UI Gothic':  'MS UI Gothic',
    'MSUIGothic':    'MS UI Gothic',
    'Meiryo':        'Meiryo',
    'Meiryo UI':     'Meiryo UI',
    'MeiryoUI':      'Meiryo UI',
    'Yu Gothic':     'Yu Gothic',
    'YuGothic':      'Yu Gothic',
    'Yu Mincho':     'Yu Mincho',
    'YuMincho':      'Yu Mincho',
}


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
    name = name.replace('-', ' ').strip()
    return FONT_ALIASES.get(name, name) or None
