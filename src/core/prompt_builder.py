import json
from pathlib import Path
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
        json_content = json.dumps(compressed_json, indent=2, ensure_ascii=False)

        # グリッド単位の取得
        grid_unit_pt = 10.0
        if json_data.get("pages") and json_data["pages"][0].get("page"):
            grid_unit_pt = json_data["pages"][0]["page"].get("grid_unit_pt", 10.0)

        # テンプレートへの埋め込み
        page_count = len(json_data.get("pages", []))

        # グリッド列数・行数の取得
        grid_cols = 60
        grid_rows = 85
        if json_data.get("pages") and json_data["pages"][0].get("page"):
            page_info = json_data["pages"][0]["page"]
            grid_cols = page_info.get("grid_cols", 60)
            grid_rows = page_info.get("grid_rows", 85)

        # セルサイズの計算（grid_unitの1.3倍で余裕を持たせる）
        # fitToPageで自動スケーリングされるため、少し大きめでOK
        scale_factor = 1.3
        row_height = round(grid_unit_pt * scale_factor, 1)  # 10pt → 13pt

        # A4幅に比例した列幅を計算
        a4_printable_px = 720
        col_px = (a4_printable_px / grid_cols) * scale_factor
        col_width = round(max((col_px - 5) / 7, 0.5), 1)

        prompt = template.replace("{{MARKDOWN_CONTENT}}", md_content)
        prompt = prompt.replace("{{JSON_CONTENT}}", json_content)
        prompt = prompt.replace("{{GRID_UNIT_PT}}", str(grid_unit_pt))
        prompt = prompt.replace("{{ROW_HEIGHT}}", str(row_height))
        prompt = prompt.replace("{{COL_WIDTH}}", str(col_width))
        prompt = prompt.replace("{{PAGE_COUNT}}", str(page_count))
        prompt = prompt.replace("{{PDF_NAME}}", pdf_name)
        prompt = prompt.replace("{{OUTPUT_FILENAME}}", f"{pdf_name}.xlsx")

        # ファイル出力
        output_path = self.output_dir / f"{pdf_name}_prompt.txt"
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(prompt)

        logger.info(f"Prompt saved to: {output_path}")
        return str(output_path)

    def _load_template(self) -> str:
        """プロンプトテンプレートを読み込む。"""
        if not self.template_path.exists():
            logger.warning(f"Template not found at {self.template_path}, using built-in template")
            return self._builtin_template()

        return self.template_path.read_text(encoding="utf-8")

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
                if style.get("fill_color"):
                    compressed_style["fill_color"] = style["fill_color"]
                if style.get("stroke_color"):
                    compressed_style["stroke_color"] = style["stroke_color"]
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

## JSON（座標・色彩・物理レイアウト）
```json
{{JSON_CONTENT}}
```

## 指示
上記のMarkdownとJSONを基に、openpyxlを使用して方眼Excel（{{GRID_UNIT_PT}}pt単位）を
生成するPythonコードを出力してください。ファイル名は `{{OUTPUT_FILENAME}}` です。
"""
