import pdfplumber
from typing import Dict, Any, Optional


def _remove_containing_rects(rects: list) -> list:
    """
    他の矩形を完全に内包している（より大きな外枠）矩形を除去する。
    内包される側（小さい矩形＝実際のセル境界）を保持し、
    内包する側（大きい外枠・行枠）を除去することで罫線の重複描画を防ぐ。
    """
    tol = 1.0  # 座標誤差許容範囲（pt）
    to_remove = set()
    for i, a in enumerate(rects):
        if i in to_remove:
            continue
        for j, b in enumerate(rects):
            if i == j or j in to_remove:
                continue
            # a が b を完全に含み、かつ同一矩形でない
            a_contains_b = (
                a['x0'] - tol <= b['x0'] and
                a['x1'] + tol >= b['x1'] and
                a['top'] - tol <= b['top'] and
                a['bottom'] + tol >= b['bottom']
            )
            is_same = (
                abs(a['x0'] - b['x0']) < tol and
                abs(a['x1'] - b['x1']) < tol and
                abs(a['top'] - b['top']) < tol and
                abs(a['bottom'] - b['bottom']) < tol
            )
            if a_contains_b and not is_same:
                to_remove.add(i)
                break
    return [r for i, r in enumerate(rects) if i not in to_remove]


def _to_hex_color(color) -> Optional[str]:
    """pdfplumber のカラー値を Excel 用 RRGGBB 16進文字列に変換する。None の場合は None を返す。"""
    if color is None:
        return None
    if isinstance(color, (int, float)):
        # グレースケール (0.0=黒 〜 1.0=白)
        v = int(round(float(color) * 255))
        return f"{v:02X}{v:02X}{v:02X}"
    if isinstance(color, (list, tuple)):
        if len(color) == 3:
            # RGB
            r, g, b = [int(round(c * 255)) for c in color]
            return f"{r:02X}{g:02X}{b:02X}"
        if len(color) == 4:
            # CMYK
            c, m, y, k = color
            r = int(round((1 - c) * (1 - k) * 255))
            g = int(round((1 - m) * (1 - k) * 255))
            b = int(round((1 - y) * (1 - k) * 255))
            return f"{r:02X}{g:02X}{b:02X}"
    return None


def extract_pdf_data(pdf_path: str) -> Dict[str, Any]:
    """
    """
    extracted_pages = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            # テキストの抽出 (x0, top, x1, bottom, text + フォント情報・色情報)
            all_words = page.extract_words(
                extra_attrs=['fontname', 'size', 'non_stroking_color']
            )
            
            # 表データの抽出
            # snap_tolerance: 近接する平行エッジを同一線にまとめる距離(pt)。
            #   デフォルト3pt だと隣接セルの左辺・右辺が2本として認識され列が倍増するため大きめに設定。
            # edge_min_length: この長さ未満のエッジを無視する(pt)。
            #   短い装飾的な線分がテーブル境界として誤検出されるのを抑制する。
            _table_settings = {
                "snap_tolerance": 5,
                "snap_y_tolerance": 5,
                "join_tolerance": 5,
                "join_y_tolerance": 5,
                "edge_min_length": 5,
                "intersection_tolerance": 5,
                "intersection_x_tolerance": 5,
                "intersection_y_tolerance": 5,
            }
            tables = page.find_tables(table_settings=_table_settings)
            table_bboxes = [table.bbox for table in tables]
            # 各テーブルの列境界X座標リスト・行境界Y座標リスト（グリッド生成用）
            table_col_x_positions = []
            table_row_y_positions = []
            # 各テーブルの全セルbbox一覧（グリッド座標計算用・LLMには渡さない）
            table_cells = []
            for table in tables:
                try:
                    valid_cells = [c for c in table.cells if c is not None]
                    col_xs = sorted(set(
                        [float(c[0]) for c in valid_cells] +
                        [float(c[2]) for c in valid_cells]
                    ))
                    row_ys = sorted(set(
                        [float(c[1]) for c in valid_cells] + [float(table.bbox[3])]
                    ))
                    table_col_x_positions.append(col_xs)
                    table_row_y_positions.append(row_ys)
                    cells_2d = [
                        [
                            {'x0': float(c[0]), 'top': float(c[1]),
                             'x1': float(c[2]), 'bottom': float(c[3])}
                            if c is not None else None
                            for c in row.cells
                        ]
                        for row in table.rows
                    ]
                    table_cells.append(cells_2d)
                except Exception:
                    table_col_x_positions.append([])
                    table_row_y_positions.append([])
                    table_cells.append([])
            # font_color フィールドを hex 文字列に変換（None の場合は省略）
            words = []
            for w in all_words:
                word = dict(w)
                raw_color = word.pop('non_stroking_color', None)
                hex_color = _to_hex_color(raw_color)
                if hex_color is not None:
                    word['font_color'] = hex_color
                # font_size は小数点以下1桁に丸める
                raw_size = word.pop('size', None)
                if raw_size is not None:
                    word['font_size'] = round(float(raw_size), 1)
                words.append(word)

            # 縦文字の抽出（upright=False の文字を x 位置でグループ化）
            # extract_words() は水平テキスト前提のため、縦文字は page.chars から別途収集する
            non_upright = [
                c for c in page.chars
                if not c.get('upright', True) and c.get('text', '').strip()
            ]
            if non_upright:
                groups: list[list] = []
                for ch in sorted(non_upright, key=lambda c: float(c['x0'])):
                    cx = (float(ch['x0']) + float(ch['x1'])) / 2
                    sz = float(ch.get('size', 10))
                    placed = False
                    for g in groups:
                        gx = (float(g[0]['x0']) + float(g[0]['x1'])) / 2
                        if abs(cx - gx) < sz:  # font size を近接判定の閾値に使う
                            g.append(ch)
                            placed = True
                            break
                    if not placed:
                        groups.append([ch])
                for g in groups:
                    g_sorted = sorted(g, key=lambda c: float(c.get('top', c.get('y0', 0))))
                    content = ''.join(c.get('text', '') for c in g_sorted)
                    if not content.strip():
                        continue
                    first, last = g_sorted[0], g_sorted[-1]
                    v_entry: dict = {
                        'x0':          float(first['x0']),
                        'top':         float(first.get('top', first.get('y0', 0))),
                        'x1':          float(last['x1']),
                        'bottom':      float(last.get('bottom', last.get('y1', 0))),
                        'text':        content,
                        'is_vertical': True,
                    }
                    raw_color = first.get('non_stroking_color')
                    hex_c = _to_hex_color(raw_color)
                    if hex_c:
                        v_entry['font_color'] = hex_c
                    raw_sz = first.get('size')
                    if raw_sz:
                        v_entry['font_size'] = round(float(raw_sz), 1)
                    if first.get('fontname'):
                        v_entry['fontname'] = first['fontname']
                    words.append(v_entry)

            # 表の内部構造（2次元配列）の取得（find_tables と同じ設定を使う）
            table_data = page.extract_tables(table_settings=_table_settings)
            # 扱いやすくするため、改行文字等が含まれていたら除去
            cleaned_tables = []
            for table in table_data:
                cleaned_table = []
                for row in table:
                    cleaned_row = [cell.replace('\n', ' ') if isinstance(cell, str) else cell for cell in row]
                    cleaned_table.append(cleaned_row)
                cleaned_tables.append(cleaned_table)
            
            # ページサイズの取得
            width = page.width
            height = page.height
            page_area = float(width) * float(height)

            # 矩形枠の抽出（フォームフィールド・罫線ボックス等）
            # ページ全体を覆う矩形（ページ境界・背景）は除外する
            rects = []
            for r in page.rects:
                rect_area = (r['x1'] - r['x0']) * (r['bottom'] - r['top'])
                if rect_area < 0.85 * page_area:
                    rects.append({
                        'x0': float(r['x0']),
                        'top': float(r['top']),
                        'x1': float(r['x1']),
                        'bottom': float(r['bottom'])
                    })

            rects = _remove_containing_rects(rects)

            # 水平・垂直エッジの抽出（罫線辺ごとの精度向上のため）
            # page.lines（明示的な線分）と page.rects の4辺を収集する
            # 重複排除のため set を使用（座標を 0.5pt 単位で丸めて比較）
            _h_seen: set = set()
            _v_seen: set = set()
            h_edges: list = []  # 水平エッジ: {'x0', 'x1', 'y'}
            v_edges: list = []  # 垂直エッジ: {'x', 'y0', 'y1'}

            def _r05(v: float) -> float:
                """0.5pt 単位で丸める。0.1pt では近接重複を取りこぼすため粗めに統一。"""
                return round(v * 2) / 2

            def _add_h(x0: float, x1: float, y: float) -> None:
                # [修正] 0.1pt → 0.5pt 単位に変更し、座標誤差による重複エッジを排除する
                key = (_r05(min(x0, x1)), _r05(max(x0, x1)), _r05(y))
                if key not in _h_seen:
                    _h_seen.add(key)
                    h_edges.append({'x0': key[0], 'x1': key[1], 'y': key[2]})

            def _add_v(x: float, y0: float, y1: float) -> None:
                # [修正] 同上
                key = (_r05(x), _r05(min(y0, y1)), _r05(max(y0, y1)))
                if key not in _v_seen:
                    _v_seen.add(key)
                    v_edges.append({'x': key[0], 'y0': key[1], 'y1': key[2],
                                    'span': key[2] - key[1]})

            for line in page.lines:
                lx0 = float(line['x0'])
                lx1 = float(line['x1'])
                lt  = float(line.get('top',    line.get('y0', 0)))
                lb  = float(line.get('bottom', line.get('y1', lt)))
                if abs(lb - lt) < 2.0 and abs(lx1 - lx0) > 2.0:   # 水平線
                    _add_h(lx0, lx1, (lt + lb) / 2)
                elif abs(lx1 - lx0) < 2.0 and abs(lb - lt) > 2.0:  # 垂直線
                    _add_v((lx0 + lx1) / 2, lt, lb)

            for r in page.rects:
                rect_area = (r['x1'] - r['x0']) * (r['bottom'] - r['top'])
                if rect_area >= 0.85 * page_area:
                    continue  # ページ全体を覆う矩形は除外
                # ストロークなし（塗りつぶしのみ）の矩形はエッジ検出に使わない。
                # ただし極細矩形（線として描かれた罫線）は stroking_color に依らず含める。
                rw = float(r['x1']) - float(r['x0'])
                rh = float(r['bottom']) - float(r['top'])
                is_line_like = rh < 2.0 or rw < 2.0
                has_stroke = r.get('stroking_color') is not None
                if not has_stroke and not is_line_like:
                    continue
                rx0, rx1 = float(r['x0']), float(r['x1'])
                rt,  rb  = float(r['top']), float(r['bottom'])
                _add_h(rx0, rx1, rt)  # 上辺
                _add_h(rx0, rx1, rb)  # 下辺
                _add_v(rx0, rt, rb)   # 左辺
                _add_v(rx1, rt, rb)   # 右辺

            _merge_gap = 2.0  # pt: これ以下のギャップは連結する

            # [修正追加] 同一X座標の垂直エッジを連結してセグメントを統合する。
            # PDF の描画命令が1本の縦線を複数の短い線分に分割して出力する場合、
            # 微小ギャップ（gap_tol 以内）を橋渡しして1本にまとめ、二重線を防ぐ。
            _by_x: dict = {}
            for _e in v_edges:
                _by_x.setdefault(_e['x'], []).append((_e['y0'], _e['y1']))
            v_edges = []
            for _x, _segs in _by_x.items():
                _segs_sorted = sorted(_segs)
                _merged = [list(_segs_sorted[0])]
                for _y0, _y1 in _segs_sorted[1:]:
                    if _y0 <= _merged[-1][1] + _merge_gap:
                        _merged[-1][1] = max(_merged[-1][1], _y1)
                    else:
                        _merged.append([_y0, _y1])
                for _y0, _y1 in _merged:
                    v_edges.append({'x': _x, 'y0': _y0, 'y1': _y1, 'span': _y1 - _y0})

            # [修正追加] 同一Y座標の水平エッジを連結してセグメントを統合する。
            # 垂直エッジと同様に、PDF が水平線を短い線分に分割して出力する場合の
            # 30%オーバーラップ閾値通過失敗を防ぐ。
            _by_y: dict = {}
            for _e in h_edges:
                _by_y.setdefault(_e['y'], []).append((_e['x0'], _e['x1']))
            h_edges = []
            for _y, _segs in _by_y.items():
                _segs_sorted = sorted(_segs)
                _merged = [list(_segs_sorted[0])]
                for _x0, _x1 in _segs_sorted[1:]:
                    if _x0 <= _merged[-1][1] + _merge_gap:
                        _merged[-1][1] = max(_merged[-1][1], _x1)
                    else:
                        _merged.append([_x0, _x1])
                for _x0, _x1 in _merged:
                    h_edges.append({'x0': _x0, 'x1': _x1, 'y': _y})

            page_data = {
                "page_number": page_number,
                "width": float(width),
                "height": float(height),
                "words": words,
                "table_bboxes": table_bboxes,
                "table_col_x_positions": table_col_x_positions,
                "table_row_y_positions": table_row_y_positions,
                "table_cells": table_cells,
                "table_data": cleaned_tables,     # \n をスペース置換済み（後方互換）
                "table_data_raw": table_data,     # \n を保持（複数行検出用）
                "rects": rects,
                "h_edges": h_edges,
                "v_edges": v_edges,
            }
            extracted_pages.append(page_data)
            
    return {"pages": extracted_pages}
