"""
Sheetling 用の LLM プロンプト定義集。

  GEN_CODE_TEMPLATE        — _gen.py をスクリプトで直接生成するための string.Template
  VISUAL_REVIEW_PROMPT     — ビジョンLLMによる視覚的検証用プロンプト
  CODE_ERROR_FIXING_PROMPT — 生成コードのエラー修正プロンプト
"""
from string import Template

GRID_SIZES = {
    "small": {
        "col_width_mm": "4.0",
        "row_height_mm": "4.0",
        "max_cols": 62,
        "max_rows": 76,
        "excel_col_width": 1.45,
        "excel_row_height": 11.34,
        "margin_left": 0.43,
        "margin_right": 0.43,
        "margin_top": 0.41,
        "margin_bottom": 0.41,
        "default_font_size": 7
    },
    "medium": {
        "col_width_mm": "6.0",
        "row_height_mm": "6.0",
        "max_cols": 36,
        "max_rows": 50,
        "excel_col_width": 2.53,
        "excel_row_height": 17.01,
        "margin_left": 0.47,
        "margin_right": 0.47,
        "margin_top": 0.41,
        "margin_bottom": 0.41,
        "default_font_size": 9
    },
    "large": {
        "col_width_mm": "8.0",
        "row_height_mm": "8.0",
        "max_cols": 26,
        "max_rows": 38,
        "excel_col_width": 3.61,
        "excel_row_height": 22.68,
        "margin_left": 0.51,
        "margin_right": 0.51,
        "margin_top": 0.49,
        "margin_bottom": 0.49,
        "default_font_size": 11
    },
    "pattern_1": {
        "col_width_mm": "2.76",
        "row_height_mm": "6.44",
        "max_cols": 90,
        "max_rows": 47,
        "excel_col_width": 1.0,
        "excel_row_height": 18.25,
        "margin_left": 0.43,
        "margin_right": 0.43,
        "margin_top": 0.41,
        "margin_bottom": 0.41,
        "default_font_size": 7
    },
    "pattern_2": {
        "col_width_mm": "5.52",
        "row_height_mm": "6.44",
        "max_cols": 45,
        "max_rows": 47,
        "excel_col_width": 2.0,
        "excel_row_height": 18.25,
        "margin_left": 0.43,
        "margin_right": 0.43,
        "margin_top": 0.41,
        "margin_bottom": 0.41,
        "default_font_size": 7
    }
}


# ---------------------------------------------------------------------------
# 自動生成テンプレート（LLM不要）
# ---------------------------------------------------------------------------

# _gen.py をスクリプトで直接生成するための string.Template。
# $変数名 が grid_params と pdf_name で置換される。
# Python コード内の { } はそのまま使えるため .format() より安全。
GEN_CODE_TEMPLATE = Template("""\
import json
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.pagebreak import Break

MAX_ROWS = $max_rows
MAX_COLS = $max_cols
COL_OFFSET = 1
ROW_PADDING = 2
EXCEL_COL_WIDTH = $excel_col_width
EXCEL_ROW_HEIGHT = $excel_row_height
DEFAULT_FONT_SIZE = $default_font_size

_dir = Path(__file__).parent
data = json.loads((_dir / "${pdf_name}_layout.json").read_text(encoding="utf-8"))

wb = Workbook()
ws = wb.active
thin = Side(style='thin')
total_pages = len(data)

# 全ページ分の方眼範囲（ページ全体）に方眼サイズを適用
_full_cols = MAX_COLS + COL_OFFSET
_full_rows = MAX_ROWS * total_pages + ROW_PADDING

for c in range(1, _full_cols + 1):
    ws.column_dimensions[get_column_letter(c)].width = EXCEL_COL_WIDTH
for r in range(1, _full_rows + 1):
    ws.row_dimensions[r].height = EXCEL_ROW_HEIGHT


def apply_border(ws, s_row, e_row, s_col, e_col, borders):
    has_top    = borders.get("top",    True)
    has_bottom = borders.get("bottom", True)
    has_left   = borders.get("left",   True)
    has_right  = borders.get("right",  True)
    for r in range(s_row, e_row + 1):
        for c in range(s_col, e_col + 1):
            target = ws.cell(row=r, column=c)
            try:
                target.alignment = Alignment(horizontal='left', vertical='center', wrap_text=False)
            except AttributeError:
                pass
            top    = thin if (r == s_row and has_top)    else None
            bottom = thin if (r == e_row and has_bottom) else None
            left   = thin if (c == s_col and has_left)   else None
            right  = thin if (c == e_col and has_right)  else None
            try:
                target.border = Border(top=top, bottom=bottom, left=left, right=right)
            except AttributeError:
                pass


max_used_row = ROW_PADDING
max_used_col = COL_OFFSET

for page in data:
    page_number = page["page_number"]
    row_offset = (page_number - 1) * MAX_ROWS + ROW_PADDING

    if page_number > 1:
        ws.row_breaks.append(Break(id=(page_number - 1) * MAX_ROWS + ROW_PADDING))

    for item in page["elements"]:
        if item["type"] == "text":
            r = item["row"] + row_offset
            c = item["col"] + COL_OFFSET
            try:
                cell = ws.cell(row=r, column=c)
                cell.value = item["content"]
                if item.get("is_vertical"):
                    cell.alignment = Alignment(text_rotation=255, vertical='top', wrap_text=False)
                else:
                    cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=False)
                font_kwargs = {"size": item.get("font_size") or DEFAULT_FONT_SIZE}
                if item.get("font_color"):
                    font_kwargs["color"] = item["font_color"]
                cell.font = Font(**font_kwargs)
                max_used_row = max(max_used_row, r)
                max_used_col = max(max_used_col, c)
            except AttributeError:
                pass

        elif item["type"] == "border_rect":
            s_row = item["row"] + row_offset
            e_row = item["end_row"] + row_offset
            s_col = item["col"] + COL_OFFSET
            e_col = item["end_col"] + COL_OFFSET
            apply_border(ws, s_row, e_row, s_col, e_col,
                         item.get("borders", {"top": True, "bottom": True, "left": True, "right": True}))
            max_used_row = max(max_used_row, e_row)
            max_used_col = max(max_used_col, e_col)

ws.page_setup.paperSize = $paper_size
ws.page_setup.orientation = '$orientation'
ws.page_margins.left = $margin_left
ws.page_margins.right = $margin_right
ws.page_margins.top = $margin_top
ws.page_margins.bottom = $margin_bottom
ws.print_area = f"A1:{get_column_letter(max_used_col)}{max_used_row}"

wb.save("output.xlsx")
""")


# ビジョンLLM（画像入力対応AIチャット）向けの視覚的検証プロンプト。
# JSONは渡さず、PDFページ画像 + 自動生成した罫線プレビュー画像の2枚比較で罫線差分を検出する。
# ページごとに生成し、対応する2枚のPNG画像と一緒にLLMに投入して使用する。
VISUAL_REVIEW_PROMPT = """\
以下の2枚の画像を比較し、**罫線（枠線）の過不足のみ**を指摘してください。
- 画像1: PDFの原本（ページ {page_number}）
- 画像2: スクリプトが自動生成した罫線プレビュー（ページ {page_number}）

## グリッドパラメータ
- 方眼サイズ: {col_width_mm}mm × {row_height_mm}mm（1マスの大きさ）
- グリッドサイズ: {max_rows} 行 × {max_cols} 列
- 座標は罫線プレビュー画像の左上セルを (row=1, col=1) とし、右・下方向に増加します

## 出力形式

PDFと比較して罫線に差異がある箇所のみ、以下のJSON形式で出力してください。
修正不要な場合は `{{"corrections": []}}` のみ出力してください。

- **add_border**: 画像2（Excelプレビュー）に存在しないが、画像1（PDF）には存在する罫線 → 追加
- **remove_border**: 画像2（Excelプレビュー）に存在するが、画像1（PDF）には存在しない罫線 → 削除

```json
{{
  "corrections": [
    {{"action": "add_border",    "page": {page_number}, "row": <開始行>, "end_row": <終了行>, "col": <開始列>, "end_col": <終了列>, "borders": {{"top": true, "bottom": true, "left": true, "right": true}}}},
    {{"action": "remove_border", "page": {page_number}, "row": <開始行>, "end_row": <終了行>, "col": <開始列>, "end_col": <終了列>}}
  ]
}}
```

【重要】出力はJSONのみ。説明文・コードブロック記号は不要です。
【重要】テキストの差異は無視してください。罫線のみが対象です。
【補足】add_border の borders は辺ごとに指定できます。不要な辺は false にしてください（例: 下線のみ → {{"bottom": true, "top": false, "left": false, "right": false}}）。"""


# ---------------------------------------------------------------------------
CODE_ERROR_FIXING_PROMPT = """あなたはPythonプログラミングのエキスパートです。
以下のExcel生成用Pythonスクリプトを実行したところ、エラーが発生しました（または期待する結果が得られませんでした）。
エラーメッセージと現在のコードを分析し、問題を修正した新しいPythonスクリプトを生成してください。

【最重要：Python文法の遵守】
修正後のコード内では、必ず Python の予約語である `True`/`False` （先頭大文字）を使用してください。小文字の `true`/`false` を含めないでください。

【エラーメッセージ・実行結果】
{error_msg}

【現在のコード】
```python
{code}
```
【既知のエラーパターンと修正方法】
- `AttributeError: 'dict' object has no attribute 'to_tree'`
  → `ws.page_margins = {{...}}` のように dict を代入している。以下の属性代入形式に修正すること:
  ```python
  ws.page_margins.left = 0.47
  ws.page_margins.right = 0.47
  ws.page_margins.top = 0.41
  ws.page_margins.bottom = 0.41
  ws.page_margins.header = 0.0
  ws.page_margins.footer = 0.0
  ```
- `AttributeError: 'MergedCell' object attribute 'value' is read-only`
  → セルへの値代入や `merge_cells` 等の処理を `try...except AttributeError: pass` で囲むこと。

【最重要】出力は、修正済みの実行可能な Python コードのみをマークダウンのコードブロック (```python ... ```) で出力してください。前後の挨拶、解説、謝罪、思考プロセスなどは一切不要です。
【最重要】出力するコード内には、`[cite: ...]` のような参照タグやアノテーション（例: ` [cite: 271]`）を絶対に含めないでください。SyntaxErrorの原因となります。純粋で実行可能なPythonコードのみを出力してください。また、コードブロックの外にいかなるテキストも記述しないでください。"""
