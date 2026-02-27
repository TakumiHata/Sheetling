import json
from pathlib import Path
from src.core.placement_generator import PlacementGenerator
from src.core.code_generator import CodeGenerator
from src.utils.logger import get_logger

logger = get_logger(__name__)


class PromptBuilder:
    """
    固定プロンプトテンプレートにMD/JSONを埋め込み、
    LLMに渡す完全なプロンプトテキストを生成する。
    """

    def __init__(self, output_dir: str, template_path: str = "templates/excel_gen_prompt.md"):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.template_path = Path(template_path)

    def build(self, md_path: str, json_path: str, pdf_name: str) -> str:
        """
        MD/JSONファイルを読み込み、テンプレートに埋め込んだプロンプトを生成・保存する。

        Args:
            md_path: MarkItDownで生成されたMDファイルのパス
            json_path: HybridAnalyzerで生成された統合JSONファイルのパス
            pdf_name: 元PDFの名前（拡張子なし）

        Returns:
            生成されたプロンプトファイルのパス
        """
        logger.info(f"Building prompt for: {pdf_name}")

        # テンプレート読み込み
        template = self._load_template()

        # MD/JSONの読み込み
        md_content = Path(md_path).read_text(encoding="utf-8")

        with open(json_path, "r", encoding="utf-8") as f:
            json_data = json.load(f)

        # JSON圧縮: 冗長な情報を除外
        compressed_json = self._compress_json(json_data)

        # グリッド単位の取得
        grid_unit_pt = 10.0
        if json_data.get("pages") and json_data["pages"][0].get("page"):
            grid_unit_pt = json_data["pages"][0]["page"].get("grid_unit_pt", 10.0)

        # グリッド列数・行数の取得
        grid_cols = 120
        grid_rows = 170
        if json_data.get("pages") and json_data["pages"][0].get("page"):
            page_info = json_data["pages"][0]["page"]
            grid_cols = page_info.get("grid_cols", 120)
            grid_rows = page_info.get("grid_rows", 170)

        # セルサイズの計算（5ptグリッドに最適化）
        scale_factor = 1.0
        row_height = round(grid_unit_pt * scale_factor, 1)

        # A4幅に比例した列幅を計算
        a4_printable_px = 720
        col_px = (a4_printable_px / grid_cols) * scale_factor
        col_width = round(max((col_px - 5) / 7, 0.5), 2)

        # ページ数
        page_count = len(json_data.get("pages", []))
        output_filename = f"{pdf_name}.xlsx"

        # 配置命令リストの生成
        placement_gen = PlacementGenerator()
        placement_result = placement_gen.generate(compressed_json)

        if placement_result.warnings:
            for w in placement_result.warnings:
                logger.warning(f"配置命令: {w}")

        # 完全なPythonコードの生成
        code_gen = CodeGenerator()
        generated_code = code_gen.generate(
            placement_result=placement_result,
            grid_cols=grid_cols,
            grid_rows=grid_rows,
            col_width=col_width,
            row_height=row_height,
            page_count=page_count,
            output_filename=output_filename,
            pdf_name=pdf_name,
        )

        # テーブル構造サマリーの生成
        table_summary = self._compute_table_structure(compressed_json)

        # テンプレートへの埋め込み
        prompt = template.replace("{{MARKDOWN_CONTENT}}", md_content)
        prompt = prompt.replace("{{GENERATED_CODE}}", generated_code)
        prompt = prompt.replace("{{TABLE_STRUCTURE_SUMMARY}}", table_summary)
        prompt = prompt.replace("{{PDF_NAME}}", pdf_name)
        prompt = prompt.replace("{{OUTPUT_FILENAME}}", output_filename)
        prompt = prompt.replace("{{PAGE_COUNT}}", str(page_count))

        # ファイル出力
        output_path = self.output_dir / f"{pdf_name}_prompt.txt"
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(prompt)

        # 生成コードも別ファイルとして保存（そのまま実行可能）
        code_path = self.output_dir / f"{pdf_name}_gen.py"
        with open(code_path, "w", encoding="utf-8") as f:
            f.write(generated_code)

        logger.info(f"Prompt saved to: {output_path}")
        logger.info(f"Generated code saved to: {code_path}")
        return str(output_path)

    def _load_template(self) -> str:
        """プロンプトテンプレートを読み込む。"""
        if not self.template_path.exists():
            logger.warning(f"Template not found at {self.template_path}, using built-in template")
            return self._builtin_template()

        return self.template_path.read_text(encoding="utf-8")

    def _compute_table_structure(self, json_data: dict) -> str:
        """
        JSONのline要素からテーブルの列・行境界を自動検出し、
        LLMが直接使える構造サマリーを生成する。
        """
        summaries = []

        for page_data in json_data.get("pages", []):
            elements = page_data.get("elements", [])

            # line要素を抽出
            lines = [e for e in elements if e.get("type") == "line"]
            if not lines:
                continue

            # 縦線と横線を分類（grid_bboxの半開区間に基づく）
            v_cols = set()  # 縦線の列位置
            h_rows = set()  # 横線の行位置
            all_rows = set()
            all_cols = set()

            for line in lines:
                bbox = line.get("grid_bbox", {})
                rs = bbox.get("row_start", 0)
                re = bbox.get("row_end", 0)
                cs = bbox.get("col_start", 0)
                ce = bbox.get("col_end", 0)

                all_rows.update([rs, re])
                all_cols.update([cs, ce])

                # 縦線: 列幅が1（半開区間で ce - cs == 1）
                if ce - cs <= 1 and re - rs > 1:
                    v_cols.add(cs)
                # 横線: 行幅が1（半開区間で re - rs == 1）
                elif re - rs <= 1 and ce - cs > 1:
                    h_rows.add(rs)

            if not v_cols or not h_rows:
                continue

            v_cols_sorted = sorted(v_cols)
            h_rows_sorted = sorted(h_rows)

            # テーブル領域の特定
            table_row_min = h_rows_sorted[0]
            table_row_max = h_rows_sorted[-1]
            table_col_min = v_cols_sorted[0]
            table_col_max = v_cols_sorted[-1]

            summary = f"### ページ {page_data.get('page', {}).get('page_number', '?')}\n\n"
            summary += f"テーブル領域: 行 {table_row_min}〜{table_row_max}, 列 {table_col_min}〜{table_col_max}\n\n"

            # 列構造の計算（縦線の間がデータ列）
            summary += "**列構造（place_cellのc1〜c2に使用）:**\n"
            for i in range(len(v_cols_sorted) - 1):
                left_border = v_cols_sorted[i]
                right_border = v_cols_sorted[i + 1]
                # 左の縦線の次の列 〜 右の縦線の前の列
                data_col_start = left_border + 1
                data_col_end = right_border - 1
                if data_col_start <= data_col_end:
                    col_label = chr(ord('A') + i)
                    summary += f"- 列{col_label}: col {data_col_start}〜{data_col_end}（縦線 {left_border} と {right_border} の間）\n"

            summary += "\n"

            # 行構造の計算（横線の間がデータ行）
            summary += "**行構造（place_cellのr1〜r2に使用）:**\n"
            for i in range(len(h_rows_sorted) - 1):
                top_border = h_rows_sorted[i]
                bottom_border = h_rows_sorted[i + 1]
                data_row_start = top_border + 1
                data_row_end = bottom_border - 1
                if data_row_start <= data_row_end:
                    summary += f"- 行{i+1}: row {data_row_start}〜{data_row_end}（横線 {top_border} と {bottom_border} の間）\n"

            summary += "\n"

            # 罫線描画のためのガイド
            summary += "**罫線の描画方法:**\n"
            summary += f"- 縦線位置: {v_cols_sorted}\n"
            summary += f"- 横線位置: {h_rows_sorted}\n"
            summary += "- 罫線はplace_cellを使わず、直接 `ws.cell(row=r, column=c).border = ...` で設定すること\n"
            summary += f"- 縦方向: 行 {table_row_min}〜{table_row_max} で各縦線位置にborder_leftを設定\n"
            summary += f"- 横方向: 列 {table_col_min}〜{table_col_max} で各横線位置にborder_topを設定\n"

            summaries.append(summary)

        if not summaries:
            return "テーブル構造が検出されませんでした。JSONのgrid_bboxを参考に配置してください。"

        return "\n".join(summaries)

    def _compress_json(self, json_data: dict) -> dict:
        """
        トークン節約のため、JSONから冗長な情報を除外する。
        - テキストが空のtext要素を除外
        - 空白のみのテキストを除外
        - style情報が全てnull/Falseの要素のstyleを省略
        """
        compressed = {
            "pdf_name": json_data.get("pdf_name"),
            "pages": []
        }

        for page_data in json_data.get("pages", []):
            compressed_page = {
                "page": page_data.get("page"),
                "elements": []
            }

            for elem in page_data.get("elements", []):
                # テキスト要素で空の場合はスキップ
                if elem["type"] == "text":
                    text = elem.get("text", "")
                    if not text or not text.strip():
                        continue

                # スタイル情報の圧縮
                style = elem.get("style", {})
                compressed_style = {}
                if style.get("stroke_width") and style["stroke_width"] > 0:
                    compressed_style["stroke_width"] = style["stroke_width"]

                border = style.get("border", {})
                if any(border.values()):
                    compressed_style["border"] = border

                compressed_elem = {
                    "type": elem["type"],
                    "grid_bbox": elem.get("grid_bbox"),
                }

                if compressed_style:
                    compressed_elem["style"] = compressed_style

                if elem.get("text"):
                    compressed_elem["text"] = elem["text"]

                if elem.get("font_size"):
                    compressed_elem["font_size"] = elem["font_size"]

                compressed_page["elements"].append(compressed_elem)

            compressed["pages"].append(compressed_page)

        return compressed

    def _builtin_template(self) -> str:
        """テンプレートファイルが見つからない場合のフォールバック。"""
        return """# Excel生成プロンプト

## Markdown（テキスト・論理構造）
```markdown
{{MARKDOWN_CONTENT}}
```

## 生成済みPythonコード（事前計算済み）
```python
{{GENERATED_CODE}}
```

## 指示
上記の生成済みコードをレビューし、Markdownのテキストが漏れていないか確認してください。
漏れがあれば追加し、最終版のPythonコードのみを出力してください。
"""
