import subprocess
from pathlib import Path
from jinja2 import Environment, FileSystemLoader

from src.core.placement import PlacementResult
from src.utils.logger import get_logger

logger = get_logger(__name__)

TEMPLATE_DIR = Path(__file__).parent / "templates"


class CodeGenerator:
    """
    配置命令リストから完全な Python スクリプトを生成する。
    テンプレートエンジン (Jinja2) を用いて基本的なスケルトンに値を流し込み、
    Ruff による自動整形を行ってLLMフレンドリーなコードを出力する。
    """

    def __init__(self):
        # Jinja2 環境の初期化
        self.env = Environment(loader=FileSystemLoader(TEMPLATE_DIR))

    def generate(
        self,
        placement_result: PlacementResult,
        grid_cols: int,
        grid_rows: int,
        col_width: float,
        row_height: float,
        page_count: int,
        output_filename: str,
        pdf_name: str,
        scale_factor: float = 1.0,
        page_breaks: list = None,
    ) -> str:
        """
        配置命令リストから完全な Python スクリプトを生成する。
        """
        text_cmds = []
        for cmd in placement_result.commands:
            if cmd.category in ("text_outside", "text_table"):
                text_cmds.append({
                    "r1": cmd.r1, "c1": cmd.c1, "r2": cmd.r2, "c2": cmd.c2,
                    "escaped_value": cmd.value.replace('\\', '\\\\').replace('"', '\\"'),
                    "scaled_font_size": round(cmd.font_size * scale_factor, 1),
                    "bold_str": ", bold=True" if cmd.font_bold else "",
                    "align": cmd.alignment or "left"
                })

        max_r = 1
        max_c = 1
        for cmd in text_cmds:
            max_r = max(max_r, cmd["r2"])
            max_c = max(max_c, cmd["c2"])
        
        for le in placement_result.line_elements:
            if le.orientation == "horizontal":
                max_r = max(max_r, le.row_start)
                max_c = max(max_c, le.col_end)
            else:
                max_r = max(max_r, le.row_end)
                max_c = max(max_c, le.col_start)

        context = {
            "pdf_name": pdf_name,
            "grid_cols": grid_cols,
            "grid_rows": grid_rows,
            "print_max_col": max_c,
            "print_max_row": max_r,
            "col_width": col_width,
            "row_height": row_height,
            "page_count": page_count,
            "page_breaks": page_breaks or [],
            "output_filename": output_filename.replace('\\', '\\\\'),
            "text_cmds": text_cmds,
            "line_elements": placement_result.line_elements
        }

        template = self.env.get_template("excel_macro_template.py.j2")
        raw_code = template.render(context)
        
        code = self._format_code_with_ruff(raw_code)

        try:
            compile(code, f"{pdf_name}_gen.py", "exec")
            logger.info(f"コード生成・整形完了: text配置={len(text_cmds)}件, 罫線={len(placement_result.line_elements)}件")
        except SyntaxError as e:
            logger.error(f"生成コードにSyntaxError: {e}")

        return code

    def _format_code_with_ruff(self, raw_code: str) -> str:
        try:
            result = subprocess.run(
                ["ruff", "format", "-"],
                input=raw_code,
                text=True,
                capture_output=True,
                check=True
            )
            return result.stdout
        except subprocess.CalledProcessError as e:
            logger.warning(f"Ruff によるフォーマットに失敗しました。未整形のコードを返します。エラー: {e.stderr}")
            return raw_code
        except FileNotFoundError:
            logger.warning("Ruff コマンドが見つかりません。未整形のコードを返します。")
            return raw_code


class PromptBuilder:
    """
    固定プロンプトテンプレートに与えられたコンテンツを埋め込み、
    LLMに渡す完全なプロンプトテキストを生成する。
    """

    def __init__(self, output_dir: str, template_path: str = "templates/excel_gen_prompt.md"):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.template_path = Path(template_path)

    def build(
        self,
        md_content: str,
        generated_code: str,
        table_summary: str,
        pdf_name: str,
        output_filename: str,
        page_count: int,
    ) -> str:
        logger.info(f"Building prompt for: {pdf_name}")
        template = self._load_template()

        prompt = template.replace("{{MARKDOWN_CONTENT}}", md_content)
        prompt = prompt.replace("{{GENERATED_CODE}}", generated_code)
        prompt = prompt.replace("{{TABLE_STRUCTURE_SUMMARY}}", table_summary)
        prompt = prompt.replace("{{PDF_NAME}}", pdf_name)
        prompt = prompt.replace("{{OUTPUT_FILENAME}}", output_filename)
        prompt = prompt.replace("{{PAGE_COUNT}}", str(page_count))

        output_path = self.output_dir / f"{pdf_name}_prompt.txt"
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(prompt)

        code_path = self.output_dir / f"{pdf_name}_gen.py"
        with open(code_path, "w", encoding="utf-8") as f:
            f.write(generated_code)

        logger.info(f"Prompt saved to: {output_path}")
        logger.info(f"Generated code saved to: {code_path}")
        return str(output_path)

    def _load_template(self) -> str:
        if not self.template_path.exists():
            logger.warning(f"Template not found at {self.template_path}, using built-in template")
            return self._builtin_template()
        return self.template_path.read_text(encoding="utf-8")

    def _builtin_template(self) -> str:
        return """# Excel生成プロンプト

## Markdown（テキスト・論理構造）
```markdown
{{MARKDOWN_CONTENT}}
```

## 生成済みPythonコード（事前計算済み）
```python
{{GENERATED_CODE}}
```

## テーブル構造
{{TABLE_STRUCTURE_SUMMARY}}

## 指示
上記の生成済みコードをレビューし、Markdownのテキストが漏れていないか確認してください。
漏れがあれば追加し、最終版のPythonコードのみを出力してください。
"""
