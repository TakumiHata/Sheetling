from src.core.constants import HORIZONTAL_GAP_FACTOR


def has_japanese(text: str) -> bool:
    return any(
        '\u3040' <= c <= '\u30ff'
        or '\u4e00' <= c <= '\u9fff'
        or '\uff00' <= c <= '\uffef'
        for c in text
    )


def join_word_texts(texts: list) -> str:
    combined = ''.join(texts)
    if has_japanese(combined):
        return combined
    return ' '.join(t for t in texts if t.strip())


def split_by_horizontal_gap(words: list, gap_factor: float = HORIZONTAL_GAP_FACTOR) -> list:
    if len(words) <= 1:
        return [words]
    sw = sorted(words, key=lambda w: float(w.get('x0', 0)))
    groups: list = [[sw[0]]]
    for w in sw[1:]:
        prev = groups[-1][-1]
        prev_x1 = float(prev.get('x1', prev.get('x0', 0)))
        curr_x0 = float(w.get('x0', 0))
        gap = curr_x0 - prev_x1
        avg_fs = (float(prev.get('font_size', prev.get('size', 10)))
                  + float(w.get('font_size', w.get('size', 10)))) / 2
        threshold = avg_fs * gap_factor
        if gap > threshold:
            groups.append([w])
        else:
            groups[-1].append(w)
    return groups
