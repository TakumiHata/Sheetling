"""
Doclingを使用してPDFから構造化データ（テキスト・表）を抽出するモジュール。
A4方眼Excel変換に必要なレイアウト情報（Bbox）と表構造（セル結合情報）を出力する。
"""

import json
import math
from pathlib import Path
from collections import OrderedDict

from docling.document_converter import DocumentConverter

from src.core.config import config
from src.utils.logger import get_logger

logger = get_logger(__name__)


class PdfExtractor:
    """Doclingを使用してPDFから構造化情報を抽出する"""

    def __init__(self):
        from docling.document_converter import DocumentConverter, PdfFormatOption
        from docling.datamodel.pipeline_options import PdfPipelineOptions, TesseractCliOcrOptions
        from docling.datamodel.base_models import InputFormat
        
        # OCRを有効化する設定 (Tesseractがインストールされている前提)
        pipeline_options = PdfPipelineOptions()
        pipeline_options.do_ocr = True
        pipeline_options.ocr_options = TesseractCliOcrOptions()
        
        self.grid_unit = config.grid.unit_pt
        self.converter = DocumentConverter(
            format_options={
                InputFormat.PDF: PdfFormatOption(
                    pipeline_options=pipeline_options
                )
            }
        )

    def extract(self, pdf_path: str, out_dir: Path) -> dict:
        """
        PDFを解析し、テキスト要素と表構造を抽出する。

        Returns:
            dict with keys: json_path, md_path
        """
        logger.info(f"Starting Docling PDF extraction for: {pdf_path}")
        pdf_name = Path(pdf_path).stem

        # Doclingによる変換実行
        result = self.converter.convert(pdf_path)
        doc = result.document

        all_pages = []
        # Doclingのドキュメントからページ情報を取得
        # ページごとの高さを保持（Executorで改ページ行計算に使用）
        page_heights = []
        for page_no, page in doc.pages.items():
            all_pages.append({
                "page_number": page_no,
                "width": float(page.size.width),
                "height": float(page.size.height),
                "cells": []
            })
            page_heights.append(float(page.size.height))

        # 各ページのグリッドを初期化
        rows_count = config.grid.target_rows
        cols_count = config.grid.target_cols
        grids = []
        for p in all_pages:
            grids.append([[{"text": "", "border": False, "merged": False} for _ in range(cols_count + 1)] for _ in range(rows_count + 1)])

        def to_top_down(l, t, r, b, ph):
            if t > b: # Bottom-up
                return l, ph - t, r, ph - b
            return l, t, r, b

        def stamp_to_grid(page_grid, l, t, r, b, pw, text, is_table_cell=False):
            if not text or not text.strip(): return
            text = text.strip()

            # 座標をグリッド単位に変換 (12pt固定)
            sr = max(1, min(rows_count, math.floor(t / self.grid_unit) + 1))
            sc = max(1, min(cols_count, math.floor(l / self.grid_unit) + 1))
            er = max(1, min(rows_count, math.floor(b / self.grid_unit) + 1))
            ec = max(1, min(cols_count, math.floor(r / self.grid_unit) + 1))
            
            # 値のセット（重複時は連結）
            if not page_grid[sr][sc]["text"]:
                page_grid[sr][sc]["text"] = text
            elif text not in page_grid[sr][sc]["text"]:
                page_grid[sr][sc]["text"] += "\n" + text
            
            # マージ判定（表セルの場合はより厳密に）
            threshold = 1.1 
            do_merge_r = (b - t) > (self.grid_unit * threshold)
            do_merge_c = (r - l) > (self.grid_unit * threshold)

            # マージフラグをセット
            for rr in range(sr, er + 1):
                for cc in range(sc, ec + 1):
                    if (rr > sr and do_merge_r) or (cc > sc and do_merge_c):
                        # 他のテキストがある場所をマージで消さない
                        if not page_grid[rr][cc]["text"]:
                            page_grid[rr][cc]["merged"] = True

        # ページごとのテーブルエリアを事前に収集
        table_areas = [[] for _ in range(len(all_pages))]
        for item, _level in doc.iterate_items():
            if hasattr(item, "data") and hasattr(item.data, "table_cells"):
                if not item.prov: continue
                prov = item.prov[0]
                p_idx = prov.page_no - 1
                if 0 <= p_idx < len(all_pages):
                    table_areas[p_idx].append(prov.bbox)

        def is_inside_table(page_idx, bbox):
            for t_bbox in table_areas[page_idx]:
                # 完全に含まれているか、大幅に重なっているか判定 (交差面積比率)
                # 判定をシンプルにするため、bboxの中心点がテーブル内にあるかチェック
                cx = (bbox.l + bbox.r) / 2
                cy = (bbox.t + bbox.b) / 2
                if t_bbox.l <= cx <= t_bbox.r and t_bbox.t <= cy <= t_bbox.b:
                    return True
            return False

        # コンテンツの抽出
        for item, _level in doc.iterate_items():
            if not item.prov: continue
            prov = item.prov[0]
            page_idx = prov.page_no - 1
            if page_idx < 0 or page_idx >= len(all_pages): continue
            
            page_grid = grids[page_idx]
            page_data = all_pages[page_idx]
            bbox = prov.bbox

            if hasattr(item, "text") and not hasattr(item, "data"):
                # 通常テキスト (テーブル内のテキストは重複を避けるためスキップ)
                if is_inside_table(page_idx, bbox):
                    continue
                l, t, r, b = to_top_down(bbox.l, bbox.t, bbox.r, bbox.b, page_data["height"])
                stamp_to_grid(page_grid, l, t, r, b, page_data["width"], item.text)
            elif hasattr(item, "data") and hasattr(item.data, "table_cells"):
                # 表データ：テーブルのBboxを基準にセルの行列から配置
                tl, tt, tr, tb = to_top_down(bbox.l, bbox.t, bbox.r, bbox.b, page_data["height"])
                t_sr = max(1, min(rows_count, math.floor(tt / self.grid_unit) + 1))
                t_sc = max(1, min(cols_count, math.floor(tl / self.grid_unit) + 1))
                
                # 146行目付近: テーブル内の各セルの座標を正規化するための事前走査
                # col_index ごとに、グリッド上の開始・終了位置の統計をとる
                col_starts = {} # col_index -> list of csc
                col_ends = {}   # col_index -> list of cec
                for cell in item.data.table_cells:
                    c_bbox = getattr(cell, "bbox", None)
                    if c_bbox:
                        cl, ct, cr, cb = to_top_down(c_bbox.l, c_bbox.t, c_bbox.r, c_bbox.b, page_data["height"])
                        csc = max(1, min(cols_count, math.floor(cl / self.grid_unit) + 1))
                        cec = max(1, min(cols_count, math.floor(cr / self.grid_unit) + 1))
                        
                        # デバッグ用: 属性名を確認
                        logger.debug(f"DEBUG: cell type={type(cell)}, attributes={dir(cell)}")
                        # Pydanticモデルの場合は model_dump も試す
                        idx = getattr(cell, "column_index", getattr(cell, "col_index", getattr(cell, "col_idx", 0)))
                        span = getattr(cell, "col_span", 1)
                        if idx not in col_starts: col_starts[idx] = []
                        col_starts[idx].append(csc)
                        
                        last_idx = idx + span - 1
                        if last_idx not in col_ends: col_ends[last_idx] = []
                        col_ends[last_idx].append(cec)
                
                # 各列の代表値を決定
                norm_starts = {}
                norm_ends = {}
                sorted_indices = sorted(col_starts.keys())
                last_end = t_sc - 1 # テーブルの開始位置の1つ前を初期値にする
                
                for idx in sorted_indices:
                    # この列の開始位置の統計（最小値）を採用しつつ、前の列の終了位置より1つ以上後ろであることを保証
                    csc = max(min(col_starts[idx]), last_end + 1)
                    norm_starts[idx] = csc
                    
                    # 終了位置。次の列がある場合はその直前、ない場合は統計の最大値
                    cec = max(col_ends[idx])
                    norm_ends[idx] = max(csc, cec) # 少なくとも開始位置以上
                    last_end = norm_ends[idx]
                
                # 列間が隙間なく並ぶように補正 (i番目の終わりは、i+1番目の始まりの1つ前)
                for i in range(len(sorted_indices) - 1):
                    curr_idx = sorted_indices[i]
                    next_idx = sorted_indices[i+1]
                    if norm_ends[curr_idx] < norm_starts[next_idx] - 1:
                        norm_ends[curr_idx] = norm_starts[next_idx] - 1
                
                # テーブル内の各セルの正確なマッピング
                for cell in item.data.table_cells:
                    # テキストが空でも枠線情報を保持するためスキップしない
                    text = cell.text.strip() if cell.text else ""
                    c_bbox = getattr(cell, "bbox", None)
                    
                    if c_bbox:
                        # Bboxから行方向の座標を特定
                        cl, ct, cr, cb = to_top_down(c_bbox.l, c_bbox.t, c_bbox.r, c_bbox.b, page_data["height"])
                        csr = max(1, min(rows_count, math.floor(ct / self.grid_unit) + 1))
                        cer = max(1, min(rows_count, math.floor(cb / self.grid_unit) + 1))
                        
                        # 列方向は正規化された境界を使用
                        cur_col_idx = getattr(cell, "column_index", getattr(cell, "col_index", getattr(cell, "col_idx", 0)))
                        csc = norm_starts.get(cur_col_idx, max(1, min(cols_count, math.floor(cl / self.grid_unit) + 1)))
                        last_col_idx = cur_col_idx + getattr(cell, "col_span", 1) - 1
                        cec = norm_ends.get(last_col_idx, max(1, min(cols_count, math.floor(cr / self.grid_unit) + 1)))
                        
                        # 行スパン：論理スパンとBboxの高さの小さい方を採用（極端なマージを防ぐ）
                        logical_rs = getattr(cell, "row_span", 1)
                        bbox_rs = cer - csr + 1
                        crs = max(logical_rs, bbox_rs)
                        # ただし、テキストがない空セルの場合は論理スパンに従う
                        if not text:
                            crs = logical_rs
                        
                        ccs = max(1, cec - csc + 1)
                    else:
                        # Bboxがない場合は行列インデックスから配置
                        cur_row_idx = getattr(cell, "row_index", 0)
                        cur_col_idx = getattr(cell, "column_index", getattr(cell, "col_index", getattr(cell, "col_idx", 0)))
                        csr = t_sr + cur_row_idx
                        csc = norm_starts.get(cur_col_idx, t_sc + cur_col_idx)
                        last_col_idx = cur_col_idx + getattr(cell, "col_span", 1) - 1
                        cec = norm_ends.get(last_col_idx, csc + getattr(cell, "col_span", 1) - 1)
                        crs = getattr(cell, "row_span", 1)
                        ccs = cec - csc + 1

                    # 範囲を境界線グリッドとしてマークし、結合をセット
                    if csr <= rows_count and csc <= cols_count:
                        # 起点セルにテキストをセット
                        if text:
                            if not page_grid[csr][csc]["text"]:
                                page_grid[csr][csc]["text"] = text
                            elif text not in page_grid[csr][csc]["text"]:
                                page_grid[csr][csc]["text"] += "\n" + text
                        
                        # スパン範囲の処理
                        for irr in range(csr, min(rows_count + 1, csr + crs)):
                            for icc in range(csc, min(cols_count + 1, csc + ccs)):
                                page_grid[irr][icc]["border"] = True
                                if irr > csr or icc > csc:
                                    page_grid[irr][icc]["merged"] = True

        # グリッドの集約（Excel要素の書き出し）
        for page_idx, page_data in enumerate(all_pages):
            grid = grids[page_idx]
            final_cells = []
            visited = [[False for _ in range(cols_count + 1)] for _ in range(rows_count + 1)]

            for r in range(1, rows_count + 1):
                for c in range(1, cols_count + 1):
                    # 起点条件：テキストがあるか、またはボーダーがあり未訪問
                    if visited[r][c] or not (grid[r][c]["text"] or grid[r][c]["border"]): continue
                    
                    rs, cs = 1, 1
                    # 横方向のスパン展開
                    for next_c in range(c + 1, cols_count + 1):
                        if visited[r][next_c] or grid[r][next_c]["text"]: break
                        # マージフラグ（PDF解析時に結合されていると判断されたもの）のみ飲み込む
                        if grid[r][next_c]["merged"]:
                            cs += 1
                        else: break
                    
                    # 縦方向のスパン展開
                    for next_r in range(r + 1, rows_count + 1):
                        all_match = True
                        for check_c in range(c, c + cs):
                            if visited[next_r][check_c] or grid[next_r][check_c]["text"]:
                                all_match = False
                                break
                            if not (grid[next_r][check_c]["merged"]):
                                all_match = False
                                break
                        if all_match: rs += 1
                        else: break
                    
                    # 訪問済みマーク
                    for vr in range(r, r + rs):
                        for vc in range(c, c + cs):
                            visited[vr][vc] = True
                    
                    # 要素として登録
                    # テキストがない場合でも、大きな空セル（マージ領域）として出力
                    final_cells.append({
                        "text": grid[r][c]["text"],
                        "r": r, "c": c, "rs": rs, "cs": cs,
                        "border": grid[r][c]["border"]
                    })
            page_data["cells"] = final_cells

        extracted = {
            "source": pdf_name,
            "pages": all_pages,
            "page_heights": page_heights,
            "grid_config": {
                "rows": rows_count,
                "cols": cols_count
            }
        }

        # JSON出力
        json_path = out_dir / f"{pdf_name}.json"
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(extracted, f, indent=2, ensure_ascii=False)
        logger.info(f"Extraction JSON saved to: {json_path}")

        # Markdown出力 (Docling自体のエクスポート機能を利用)
        md_content = doc.export_to_markdown()
        md_path = out_dir / f"{pdf_name}.md"
        with open(md_path, "w", encoding="utf-8") as f:
            f.write(md_content)
        logger.info(f"Extraction MD saved to: {md_path}")

        return {
            "json_path": str(json_path),
            "md_path": str(md_path),
        }
