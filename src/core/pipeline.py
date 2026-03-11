"""
Sheetling パイプライン。
新しい多段パイプライン方式:
1. 解析 (pdfplumber)
2. 構造化・座標計算 (LLMプロンプト生成)
3. 描画 (openpyxl)
"""

import json
from pathlib import Path

from src.parser.pdf_extractor import extract_pdf_data
from src.templates.prompts import CHUNKING_PROMPT, STRUCTURE_ALIGNMENT_PROMPT, GRID_MAPPING_PROMPT, COMMAND_GENERATION_PROMPT, PAGE_FIT_PROMPT, EXCEL_CODE_GEN_PROMPT, CODE_ERROR_FIXING_PROMPT
from src.core.config import config
from src.utils.logger import get_logger

logger = get_logger(__name__)


class SheetlingPipeline:
    """
    1. PDF を解析してプロンプトを出力する (Phase 1)。
    2. ユーザーがLLMから得たJSONを実行し、Excel方眼紙を生成する (Phase 3)。
    """

    def __init__(self, output_base_dir: str):
        self.output_base_dir = Path(output_base_dir)

    def generate_prompts(self, pdf_path: str, in_base_dir: str = "data/in") -> dict:
        """
        Phase 1: PDFを解析し、LLMに渡すためのプロンプトを data/out/ に出力する。
        """
        logger.info(f"--- [Phase 1] PDF解析 & プロンプト生成: {Path(pdf_path).name} ---")
        path_obj = Path(pdf_path)
        pdf_name = path_obj.stem

        # 出力先のディレクトリを作成
        try:
            rel_path = path_obj.parent.relative_to(Path(in_base_dir))
            out_dir = self.output_base_dir / rel_path
        except ValueError:
            out_dir = self.output_base_dir / pdf_name
            
        out_dir.mkdir(parents=True, exist_ok=True)

        # PDFから情報を抽出 (markdown_content と pages)
        extracted_data = extract_pdf_data(pdf_path)
        markdown_str = extracted_data["markdown_content"]
        
        extracted_json_path = out_dir / f"{pdf_name}_extracted.json"
        with open(extracted_json_path, "w", encoding="utf-8") as f:
            json.dump(extracted_data, f, indent=2, ensure_ascii=False)

        # 抽出データを文字列化してプロンプトに埋め込む
        input_data_str = json.dumps(extracted_data, indent=2, ensure_ascii=False)
        
        # STEP 2 用の入力テンプレート文字列
        step2_input_template = f"==== Markdown テキスト ====\n{markdown_str}\n\n==== STEP 1 (チャンク抽出) の出力結果 ====\n[ここにSTEP 1のJSON出力結果を貼り付けてください]"

        prompt_1 = CHUNKING_PROMPT.format(input_data=input_data_str)
        prompt_2 = STRUCTURE_ALIGNMENT_PROMPT.format(input_data=step2_input_template)
        prompt_3 = GRID_MAPPING_PROMPT.format(input_data="[ここにSTEP 2の出力（JSON部分のみ）を貼り付けてください]")
        prompt_4 = COMMAND_GENERATION_PROMPT.format(input_data="[ここにSTEP 3の出力（JSON部分のみ）を貼り付けてください]")
        prompt_5 = PAGE_FIT_PROMPT.format(input_data="[ここにSTEP 4の出力（JSON部分のみ）を貼り付けてください]")
        prompt_6 = EXCEL_CODE_GEN_PROMPT.format(input_data="[ここにSTEP 5の出力（JSON部分のみ）を貼り付けてください]")

        # プロンプト保存用のディレクトリを作成
        prompts_dir = out_dir / "prompts"
        prompts_dir.mkdir(parents=True, exist_ok=True)

        # プロンプトを別々のファイルとして出力
        prompt_1_path = prompts_dir / f"{pdf_name}_prompt_step1.txt"
        prompt_2_path = prompts_dir / f"{pdf_name}_prompt_step2.txt"
        prompt_3_path = prompts_dir / f"{pdf_name}_prompt_step3.txt"
        prompt_4_path = prompts_dir / f"{pdf_name}_prompt_step4.txt"
        prompt_5_path = prompts_dir / f"{pdf_name}_prompt_step5.txt"
        prompt_6_path = prompts_dir / f"{pdf_name}_prompt_step6.txt"
        
        with open(prompt_1_path, "w", encoding="utf-8") as f:
            f.write(prompt_1)
        with open(prompt_2_path, "w", encoding="utf-8") as f:
            f.write(prompt_2)
        with open(prompt_3_path, "w", encoding="utf-8") as f:
            f.write(prompt_3)
        with open(prompt_4_path, "w", encoding="utf-8") as f:
            f.write(prompt_4)
        with open(prompt_5_path, "w", encoding="utf-8") as f:
            f.write(prompt_5)
        with open(prompt_6_path, "w", encoding="utf-8") as f:
            f.write(prompt_6)

        # 生成コード保存用の空ファイルを作成 (STEP 6)
        generated_code_path = out_dir / f"{pdf_name}_gen.py"
        if not generated_code_path.exists():
            with open(generated_code_path, "w", encoding="utf-8") as f:
                f.write("# Please paste final AI Python code (from STEP 6) here.\n")

        logger.info(f"✅ Phase 1 完了: {pdf_name}")
        logger.info(f"  抽出データ: {extracted_json_path}")
        logger.info(f"  プロンプトSTEP1〜6を出力しました")
        logger.info(f"  ※ STEP1から順にLLMに入力し、最終的な出力結果を {generated_code_path} に保存してください。")

        return {
            "json_path": str(extracted_json_path),
            "prompt_step1_path": str(prompt_1_path),
            "prompt_step6_path": str(prompt_6_path),
            "generated_code_base_path": str(generated_code_path)
        }

    def render_excel(self, pdf_name: str, specific_out_dir: str = None) -> str:
        """
        Phase 3: AI出力の生成コードを読み込み、Excel方眼紙を描画する。
        """
        logger.info(f"--- [Phase 3] Excel生成: {pdf_name} ---")
        if specific_out_dir:
            out_dir = Path(specific_out_dir)
        else:
            out_dir = self.output_base_dir / pdf_name
        
        output_xlsx_path = out_dir / f"{pdf_name}.xlsx"
        generated_code_path = out_dir / f"{pdf_name}_gen.py"

        # 生成コード (STEP 6) が存在すれば実行
        if generated_code_path.exists():
            with open(generated_code_path, "r", encoding="utf-8") as f:
                content = f.read().strip()
            
            # プレースホルダーのみの場合や極端に短い場合はスキップ
            # "# Please paste" から始まっている場合でも、改行以降に有効なコードがあれば実行するように変更
            code_lines = [line for line in content.splitlines() if not line.strip().startswith("#")]
            actual_code = "\n".join(code_lines).strip()
            is_placeholder = len(actual_code) < 50
            
            if content and not is_placeholder:
                logger.info(f"✨ 生成されたコードを実行します: {generated_code_path.name}")
                import subprocess
                import os
                import sys
                
                try:
                    env = os.environ.copy()
                    env["PYTHONPATH"] = os.getcwd()
                    
                    result = subprocess.run(
                        [sys.executable, generated_code_path.name],
                        cwd=str(out_dir),
                        env=env,
                        capture_output=True,
                        text=True
                    )
                    
                    if result.returncode == 0:
                        temp_xlsx = out_dir / "output.xlsx"
                        if temp_xlsx.exists():
                            temp_xlsx.replace(output_xlsx_path)
                            logger.info(f"✅ Phase 3 完了 (コード生成経由): {output_xlsx_path}")
                            return str(output_xlsx_path)
                        else:
                            error_msg = "生成コードは正常終了しましたが、output.xlsx が生成されませんでした。"
                            logger.error(f"❌ {error_msg}")
                            self._generate_error_prompt(out_dir, pdf_name, error_msg, content)
                    else:
                        error_msg = f"生成コードの実行に失敗しました:\n{result.stderr}"
                        logger.error(f"❌ {error_msg}")
                        self._generate_error_prompt(out_dir, pdf_name, error_msg, content)
                except Exception as e:
                    error_msg = f"生成コード実行中に例外が発生しました: {e}"
                    logger.error(f"❌ {error_msg}")
                    self._generate_error_prompt(out_dir, pdf_name, error_msg, content)
            else:
                logger.warning(f"⚠️ 生成コードファイル {generated_code_path.name} が空、または未編集です。")
        else:
            logger.error(f"❌ 生成コードファイル {generated_code_path.name} が見つかりません。STEP 6 の結果を保存してください。")

        raise RuntimeError(f"Excelの生成に失敗しました ({pdf_name})")

    def _generate_error_prompt(self, out_dir: Path, pdf_name: str, error_msg: str, current_code: str):
        prompt_text = CODE_ERROR_FIXING_PROMPT.format(error_msg=error_msg, code=current_code)
        prompts_dir = out_dir / "prompts"
        prompts_dir.mkdir(parents=True, exist_ok=True)
        error_prompt_path = prompts_dir / f"{pdf_name}_prompt_error_fix.txt"
        with open(error_prompt_path, "w", encoding="utf-8") as f:
            f.write(prompt_text)
        logger.info(f"💡 エラー修正用プロンプトを出力しました: {error_prompt_path}")


