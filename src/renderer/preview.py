from pathlib import Path


def _compute_cell_metrics(img_w, img_h, max_c, max_r, content_bounds):
    if content_bounds:
        page_w = content_bounds.get('page_width', 1.0)
        page_h = content_bounds.get('page_height', 1.0)
        px_per_pt_x = img_w / page_w
        px_per_pt_y = img_h / page_h
        offset_x = content_bounds.get('min_x', 0.0) * px_per_pt_x
        offset_y = content_bounds.get('min_y', 0.0) * px_per_pt_y
        cell_w = content_bounds.get('grid_w', 1.0) * px_per_pt_x
        cell_h = content_bounds.get('grid_h', 1.0) * px_per_pt_y
    else:
        offset_x = 0.0
        offset_y = 0.0
        cell_w = img_w / max_c
        cell_h = img_h / max_r
    return offset_x, offset_y, cell_w, cell_h


def _draw_grid_lines(draw, img_w, img_h, max_c, max_r, cx, cy):
    for c in range(max_c + 1):
        x = cx(c)
        if 0 <= x < img_w:
            draw.line([(x, 0), (x, img_h)], fill='#E0E0E0', width=1)
    for r in range(max_r + 1):
        y = cy(r)
        if 0 <= y < img_h:
            draw.line([(0, y), (img_w, y)], fill='#E0E0E0', width=1)


def _draw_borders(draw, elements, cx, cy, cell_w, cell_h):
    border_width = max(2, int(min(cell_w, cell_h) / 7))
    for elem in elements:
        if elem.get('type') != 'border_rect':
            continue
        r1 = cy(elem['row'] - 1)
        r2 = cy(elem['end_row'] - 1)
        c1 = cx(elem['col'] - 1)
        c2 = cx(elem['end_col'] - 1)
        borders = elem.get('borders', {'top': True, 'bottom': True, 'left': True, 'right': True})
        if borders.get('top',    True): draw.line([(c1, r1), (c2, r1)], fill='black', width=border_width)
        if borders.get('bottom', True): draw.line([(c1, r2), (c2, r2)], fill='black', width=border_width)
        if borders.get('left',   True): draw.line([(c1, r1), (c1, r2)], fill='black', width=border_width)
        if borders.get('right',  True): draw.line([(c2, r1), (c2, r2)], fill='black', width=border_width)


def _draw_greyout(draw, elements, img_w, img_h, cx, cy):
    border_elems = [e for e in elements if e.get('type') == 'border_rect']
    if not border_elems:
        return
    content_max_col = max(e.get('end_col', e['col']) for e in border_elems)
    content_max_row = max(e.get('end_row', e['row']) for e in border_elems)
    grey_x = cx(content_max_col - 1)
    grey_y = cy(content_max_row - 1)
    grey_fill = (210, 210, 210)
    if grey_x < img_w:
        draw.rectangle([(grey_x, 0), (img_w, img_h)], fill=grey_fill)
    if grey_y < img_h:
        right_limit = min(grey_x, img_w)
        draw.rectangle([(0, grey_y), (right_limit, img_h)], fill=grey_fill)


def _draw_labels(draw, max_c, max_r, cx, cy, cell_w, cell_h, img_w, img_h):
    from PIL import ImageFont
    try:
        font = ImageFont.load_default(size=max(8, int(cell_h * 0.8)))
    except TypeError:
        font = ImageFont.load_default()
    label_color = (200, 0, 0)
    for c in range(1, max_c + 1, 5):
        lx = cx(c - 1) + cell_w / 2
        if 0 <= lx < img_w:
            draw.text((lx, 1), str(c), fill=label_color, font=font)
    for r in range(1, max_r + 1, 5):
        ly = cy(r - 1) + cell_h / 2
        if 0 <= ly < img_h:
            draw.text((1, ly), str(r), fill=label_color, font=font)


def generate_border_preview(page_layout: dict, grid_params: dict, output_path: str,
                            pdf_image_path: str | None = None,
                            content_bounds: dict | None = None) -> None:
    from PIL import Image, ImageDraw

    max_c = int(grid_params.get('max_cols', 54))
    max_r = int(grid_params.get('max_rows', 42))

    if pdf_image_path and Path(pdf_image_path).exists():
        with Image.open(pdf_image_path) as ref:
            img_w, img_h = ref.size
    else:
        img_w = int(20.0 * max_c) + 1
        img_h = int(14.0 * max_r) + 1

    img = Image.new('RGB', (img_w, img_h), 'white')
    draw = ImageDraw.Draw(img)

    offset_x, offset_y, cell_w, cell_h = _compute_cell_metrics(
        img_w, img_h, max_c, max_r, content_bounds)

    def cx(col: float) -> int: return int(offset_x + col * cell_w)
    def cy(row: float) -> int: return int(offset_y + row * cell_h)

    elements = page_layout.get('elements', [])
    _draw_grid_lines(draw, img_w, img_h, max_c, max_r, cx, cy)
    _draw_borders(draw, elements, cx, cy, cell_w, cell_h)
    _draw_greyout(draw, elements, img_w, img_h, cx, cy)
    _draw_labels(draw, max_c, max_r, cx, cy, cell_w, cell_h, img_w, img_h)

    img.save(output_path)


def _draw_run_overlay(draw, runs, cx, cy, cell_w, cell_h, color):
    border_width = max(2, int(min(cell_w, cell_h) / 7))
    for run in runs:
        if run['type'] == 'H':
            r = run['row']
            y = cy(r - 1)
            x1 = cx(run['col_start'] - 1)
            x2 = cx(run['col_end'] - 1)
            draw.line([(x1, y), (x2, y)], fill=color, width=border_width)
        else:
            c = run['col']
            x = cx(c - 1)
            y1 = cy(run['row_start'] - 1)
            y2 = cy(run['row_end'] - 1)
            draw.line([(x, y1), (x, y2)], fill=color, width=border_width)


def _label_anchor(run: dict, cx, cy, cell_w, cell_h) -> tuple:
    """ラン上のラベル配置位置を ID で分散させる。

    密集領域 (テーブル等) で複数ランが同じ col_range/row_range を共有する場合、
    中央配置だと全ラベルが重なる。ID を 5 段階で割って配置位置をズラす。
    """
    rid = run['id']
    if run['type'] == 'H':
        cs, ce = run['col_start'], run['col_end']
        span = ce - cs
        frac = 0.15 + 0.7 * ((rid % 5) / 4)
        anchor_col = cs + span * frac - 1
        x = cx(anchor_col)
        y = cy(run['row'] - 1) - int(cell_h * 0.55)
    else:
        rs, re_ = run['row_start'], run['row_end']
        span = re_ - rs
        frac = 0.15 + 0.7 * ((rid % 5) / 4)
        anchor_row = rs + span * frac - 1
        x = cx(run['col'] - 1) + 2
        y = cy(anchor_row) - int(cell_h * 0.4)
    return int(x), int(y)


def _draw_label_with_background(draw, x: int, y: int, text: str, font, color):
    bbox = draw.textbbox((x, y), text, font=font)
    pad = 1
    bg = (255, 255, 255, 220)
    border = (140, 0, 0, 220)
    draw.rectangle(
        [(bbox[0] - pad, bbox[1] - pad), (bbox[2] + pad, bbox[3] + pad)],
        fill=bg, outline=border, width=1,
    )
    draw.text((x, y), text, fill=color, font=font)


def _draw_run_id_labels(draw, runs, cx, cy, cell_w, cell_h, color):
    from PIL import ImageFont
    try:
        font = ImageFont.load_default(size=max(8, int(cell_h * 0.65)))
    except TypeError:
        font = ImageFont.load_default()
    for run in runs:
        x, y = _label_anchor(run, cx, cy, cell_w, cell_h)
        _draw_label_with_background(draw, x, y, str(run['id']), font, color)


def generate_diff_overlay(pdf_image_path: str, runs_with_ids: list, grid_params: dict,
                          output_path: str, content_bounds: dict | None = None) -> None:
    """PDF 原本画像に現プレビュー罫線を半透明赤で重ね、各罫線に ID を描画する。

    LLM はこの 1 枚を見て、赤線が PDF の黒線と一致するかだけを判定する。
    座標推定の負担を減らし、削除指示は ID 番号のみで完結する。
    """
    from PIL import Image, ImageDraw

    max_c = int(grid_params.get('max_cols', 54))
    max_r = int(grid_params.get('max_rows', 42))

    base = Image.open(pdf_image_path).convert('RGBA')
    img_w, img_h = base.size
    overlay = Image.new('RGBA', (img_w, img_h), (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)

    offset_x, offset_y, cell_w, cell_h = _compute_cell_metrics(
        img_w, img_h, max_c, max_r, content_bounds)

    def cx(col: float) -> int: return int(offset_x + col * cell_w)
    def cy(row: float) -> int: return int(offset_y + row * cell_h)

    line_color = (220, 20, 20, 180)
    label_color = (140, 0, 0, 255)
    _draw_run_overlay(draw, runs_with_ids, cx, cy, cell_w, cell_h, line_color)
    _draw_run_id_labels(draw, runs_with_ids, cx, cy, cell_w, cell_h, label_color)

    out = Image.alpha_composite(base, overlay).convert('RGB')
    out.save(output_path)
