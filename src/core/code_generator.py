import subprocess
from pathlib import Path
from jinja2 import Environment, FileSystemLoader

from src.core.placement_generator import PlacementResult
from src.utils.logger import get_logger

logger = get_logger(__name__)

# テンプレートディレクトリのパスを設定
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

        # text 要素の専用データ整形
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

        # 実際の使用範囲（最大行・列）を計算して印刷範囲を最適化する
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

        # Jinja2 コンテキスト変数の用意
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

        # テンプレート読み込み
        template = self.env.get_template("excel_macro_template.py.j2")
        raw_code = template.render(context)
        
        # ruff でコードを整形
        code = self._format_code_with_ruff(raw_code)

        # 生成コードの文法チェック
        try:
            compile(code, f"{pdf_name}_gen.py", "exec")
            logger.info(f"コード生成・整形完了: text配置={len(text_cmds)}件, 罫線={len(placement_result.line_elements)}件")
        except SyntaxError as e:
            logger.error(f"生成コードにSyntaxError: {e}")

        return code

    def _format_code_with_ruff(self, raw_code: str) -> str:
        """
        Ruff フォーマッタを使用して文字列の Python コードを整形する。
        """
        try:
            # ruff format コマンドを標準入力経由で実行 (- によって stdin を読み込む)
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
