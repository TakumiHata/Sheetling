"""
Doclingを使用してPDFから構造化データ（テキスト・表）を抽出するモジュール。
A4方眼Excel変換に必要なレイアウト情報（Bbox）と表構造（セル結合情報）を出力する。
"""

import json
import math
from pathlib import Path
import pandas as pd

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
        cumulative_heights = []
        current_cumulative = 0.0
        
        for page_no, page in doc.pages.items():
            height = float(page.size.height)
            all_pages.append({
                "page_number": page_no,
                "width": float(page.size.width),
                "height": height,
                "cells": []
            })
            page_heights.append(height)
            cumulative_heights.append(current_cumulative)
            current_cumulative += height

        # グリッドの最大サイズ (定数)
        rows_count = config.grid.target_rows
        cols_count = config.grid.target_cols

        # 全抽出要素を1つのフラットなリストに収集 (Pandasで一気に処理するため)
        # item: {"text": str, "l": float, "t": float, "r": float, "b": float, "page": int, "border": bool, "type": str, "rs": int, "cs": int}
        elements = []

        def to_top_down(l, t, r, b, ph):
            if t > b: return l, ph - t, r, ph - b
            return l, t, r, b

        # テーブル領域の収集 (テキストと重複判定用)
        table_areas = [[] for _ in range(len(all_pages))]
        for item, _level in doc.iterate_items():
            if hasattr(item, "data") and hasattr(item.data, "table_cells") and item.prov:
                prov = item.prov[0]
                p_idx = prov.page_no - 1
                if 0 <= p_idx < len(all_pages):
                    table_areas[p_idx].append(prov.bbox)

        def is_inside_table(page_idx, bbox):
            cx, cy = (bbox.l + bbox.r) / 2, (bbox.t + bbox.b) / 2
            for t_bbox in table_areas[page_idx]:
                if t_bbox.l <= cx <= t_bbox.r and t_bbox.t <= cy <= t_bbox.b:
                    return True
            return False

        table_count = 0
        # コンテンツの抽出フェーズ (単純なダンプ)
        for item, _level in doc.iterate_items():
            if not item.prov: continue
            prov = item.prov[0]
            page_idx = prov.page_no - 1
            if page_idx < 0 or page_idx >= len(all_pages): continue
            
            page_data = all_pages[page_idx]
            bbox = prov.bbox

            if hasattr(item, "text") and not hasattr(item, "data"):
                # 通常テキスト
                if is_inside_table(page_idx, bbox) or not item.text.strip(): continue
                l, t, r, b = to_top_down(bbox.l, bbox.t, bbox.r, bbox.b, page_data["height"])
                
                # 前ページまでの高さを累積する
                t += cumulative_heights[page_idx]
                b += cumulative_heights[page_idx]
                
                elements.append({
                    "text": item.text.strip(),
                    "l": l, "t": t, "r": r, "b": b,
                    "page_idx": page_idx,
                    "border": False, "type": "text", "rs": 1, "cs": 1,
                    "table_id": -1, "row_idx": -1, "col_idx": -1
                })
            
            elif hasattr(item, "data") and hasattr(item.data, "table_cells"):
                # 表データ
                table_id = table_count
                table_count += 1
                for cell in item.data.table_cells:
                    c_bbox = getattr(cell, "bbox", None)
                    if not c_bbox: continue # Bboxがない空セルは方眼計算ができないためスキップ
                    
                    text = cell.text.strip() if cell.text else ""
                    cl, ct, cr, cb = to_top_down(c_bbox.l, c_bbox.t, c_bbox.r, c_bbox.b, page_data["height"])
                    
                    # 前ページまでの高さを累積する
                    ct += cumulative_heights[page_idx]
                    cb += cumulative_heights[page_idx]
                    
                    row_idx = getattr(cell, "start_row_offset_idx", -1)
                    col_idx = getattr(cell, "start_col_offset_idx", -1)
                    end_row_idx = getattr(cell, "end_row_offset_idx", row_idx)
                    end_col_idx = getattr(cell, "end_col_offset_idx", col_idx)
                    
                    elements.append({
                        "text": text,
                        "l": cl, "t": ct, "r": cr, "b": cb,
                        "page_idx": page_idx,
                        "border": True, "type": "table_cell",
                        "rs": 1, "cs": 1,
                        "table_id": table_id, 
                        "row_idx": row_idx, "col_idx": col_idx,
                        "end_row_idx": end_row_idx, "end_col_idx": end_col_idx
                    })

        # --- Pandas を用いたベクトル演算による方眼マッピング ---
        if elements:
            df = pd.DataFrame(elements)
            
            # クリップ関数のベクトル化（列用: 1 <= val <= max_val、行用: 下限 1 のみ）
            def clip_col(series, max_val):
                return series.clip(lower=1, upper=max_val)

            def clip_row(series):
                return series.clip(lower=1)
            
            # 各要素の座標を方眼(grid_unit)で割ってインデックス化
            # Y座標(行)は複数ページにまたがるため最大値の上限を設けない
            df["sr"] = clip_row((df["t"] / self.grid_unit).apply(math.floor) + 1)
            df["sc"] = clip_col((df["l"] / self.grid_unit).apply(math.floor) + 1, cols_count)
            df["er"] = clip_row((df["b"] / self.grid_unit).apply(math.floor) + 1)
            df["ec"] = clip_col((df["r"] / self.grid_unit).apply(math.floor) + 1, cols_count)
            
            # --- テーブルセルの境界を論理インデックス単位で揃える（隙間をなくして結合） ---
            mask_table = df["table_id"] != -1
            if mask_table.any():
                # テーブルごとに処理
                for (page_idx, tid), tdf in df[mask_table].groupby(["page_idx", "table_id"]):
                    # 1. 各論理列の境界X座標(左端・右端)を計算
                    # min_colの左端はsc, 各colの右端(end_col_idxの右)のecを集める
                    col_starts = tdf.groupby("col_idx")["sc"].min()
                    col_ends = tdf.groupby("end_col_idx")["ec"].max()
                    
                    # 連続する列の境界（前列のendと次列のstart）の平均を新しい境界とする
                    col_boundaries = {}
                    all_cols = sorted(set(col_starts.index) | set(col_ends.index))
                    if all_cols:
                        min_c = int(min(all_cols))
                        max_c = int(max(all_cols))
                        col_boundaries[min_c] = col_starts[min_c] # 左端
                        for c in range(min_c, max_c):
                            if c in col_ends and (c+1) in col_starts:
                                # 境界は平均をとって丸める
                                boundary = round((col_ends[c] + col_starts[c+1]) / 2)
                                col_boundaries[c+1] = boundary
                            else:
                                # 歯抜けの場合は前のendを使うか次のstartを使う
                                if c in col_ends: col_boundaries[c+1] = col_ends[c]
                                elif (c+1) in col_starts: col_boundaries[c+1] = col_starts[c+1]
                        col_boundaries[max_c + 1] = col_ends[max_c] # 右端

                    # 2. 各論理行の境界Y座標(上端・下端)を計算
                    row_starts = tdf.groupby("row_idx")["sr"].min()
                    row_ends = tdf.groupby("end_row_idx")["er"].max()
                    
                    row_boundaries = {}
                    all_rows = sorted(set(row_starts.index) | set(row_ends.index))
                    if all_rows:
                        min_r = int(min(all_rows))
                        max_r = int(max(all_rows))
                        row_boundaries[min_r] = row_starts[min_r] # 上端
                        for r in range(min_r, max_r):
                            if r in row_ends and (r+1) in row_starts:
                                boundary = round((row_ends[r] + row_starts[r+1]) / 2)
                                row_boundaries[r+1] = boundary
                            else:
                                if r in row_ends: row_boundaries[r+1] = row_ends[r]
                                elif (r+1) in row_starts: row_boundaries[r+1] = row_starts[r+1]
                        row_boundaries[max_r + 1] = row_ends[max_r] # 下端

                    # 3. 計算された境界を元のデータフレームに適用
                    for idx, row in tdf.iterrows():
                        c_start = row["col_idx"]
                        c_end = row["end_col_idx"]
                        r_start = row["row_idx"]
                        r_end = row["end_row_idx"]
                        
                        if c_start in col_boundaries and (c_end + 1) in col_boundaries:
                            df.at[idx, "sc"] = col_boundaries[c_start]
                            # 境界を共有するため、右端は次の列の左端 - 1 にはしない（隣接させるには方眼上で sr と次の sr の関係で決まる）
                            # 方眼セル単位での専有領域は sc から ec なので、結合させるなら ec = boundary - 1 とすると隙間ができる。
                            # 隣接するセルの sc が boundary のとき、手前のセルの ec を boundary - 1 にすることで物理的に隣接する。
                            df.at[idx, "ec"] = col_boundaries[c_end + 1] - 1
                        
                        if r_start in row_boundaries and (r_end + 1) in row_boundaries:
                            df.at[idx, "sr"] = row_boundaries[r_start]
                            df.at[idx, "er"] = row_boundaries[r_end + 1] - 1

            # --- 空行の圧縮処理 ---
            # ページ間などの過剰な空行の連続を最大1行に圧縮する
            used_rows = set()
            for r_start, r_end in zip(df["sr"], df["er"]):
                used_rows.update(range(int(r_start), int(r_end) + 1))
                
            if used_rows:
                max_used = max(used_rows)
                row_map = {}
                current_new_r = 1
                empty_count = 0
                for r in range(1, max_used + 1):
                    if r in used_rows:
                        row_map[r] = current_new_r
                        current_new_r += 1
                        empty_count = 0
                    else:
                        if empty_count < 1:
                            row_map[r] = current_new_r
                            current_new_r += 1
                            empty_count += 1
                        else:
                            row_map[r] = current_new_r - 1
                            
                df["sr"] = df["sr"].map(row_map)
                df["er"] = df["er"].map(row_map)

            # スパン決定: 補正後の仮想座標から物理的なスパンを再計算
            df["rs"] = (df["er"] - df["sr"] + 1).clip(lower=1)
            df["cs"] = (df["ec"] - df["sc"] + 1).clip(lower=1)
            
            # --- ページごとの改ページ行(物理行)を記録 ---
            page_breaks_list = []
            for p_idx in range(len(all_pages) - 1):
                p_df = df[df["page_idx"] == p_idx]
                if not p_df.empty:
                    page_breaks_list.append(int(p_df["er"].max()))
                else:
                    page_breaks_list.append(0)
            
            # 重複の統合 (同じページ、同じsr, scに落ちた要素のテキストを結合)
            # Borderフラグは一つでもTrueならTrueにする
            grouped = df.groupby(["page_idx", "sr", "sc"]).agg({
                "text": lambda x: "\n".join(set(filter(None, x))),
                "rs": "max",
                "cs": "max",
                "border": "any"
            }).reset_index()
            
            # 出力用へ詰め直し
            for _, row in grouped.iterrows():
                # textが空でborderもないものはスキップ（ゴミ要素）
                if not row["text"] and not row["border"]: continue
                
                final_cells = all_pages[row["page_idx"]]["cells"]
                final_cells.append({
                    "text": row["text"],
                    "r": int(row["sr"]),
                    "c": int(row["sc"]),
                    "rs": int(row["rs"]),
                    "cs": int(row["cs"]),
                    "border": bool(row["border"])
                })

        extracted = {
            "source": pdf_name,
            "pages": all_pages,
            "page_heights": page_heights,
            "page_breaks": page_breaks_list if 'page_breaks_list' in locals() else [],
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
