import json
from pathlib import Path
from src.utils.logger import get_logger

logger = get_logger(__name__)


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
        """
        与えられたデータをテンプレートに埋め込んだプロンプトを生成・保存する。

        Args:
            md_content: MarkItDownで生成されたMarkdownのテキスト内容
            generated_code: 事前生成されたPythonコード
            table_summary: テーブル構造のサマリーテキスト
            pdf_name: 元PDFの名前（拡張子なし）
            output_filename: 出力するExcelファイル名
            page_count: ページ数

        Returns:
            生成されたプロンプトファイルのパス
        """
        logger.info(f"Building prompt for: {pdf_name}")

        # テンプレート読み込み
        template = self._load_template()

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

        # 生成コードも別ファイルとして一緒に保存（デバッグ及びそのまま実行可能にするため）
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

## テーブル構造
{{TABLE_STRUCTURE_SUMMARY}}

## 指示
上記の生成済みコードをレビューし、Markdownのテキストが漏れていないか確認してください。
漏れがあれば追加し、最終版のPythonコードのみを出力してください。
"""
