"""
Sheetling 用の LLM プロンプト定義集。
2ステップ・パイプライン方式:
  Step 1: TABLE_ANCHOR_PROMPT  — PDF解析データ → Excelレイアウト仕様JSON（列アンカー確定）
  Step 2: CODE_GEN_PROMPT      — レイアウト仕様JSON → Python(openpyxl)コード
"""

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
        "margin_bottom": 0.41
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
        "margin_bottom": 0.41
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
        "margin_bottom": 0.49
    }
}

TABLE_ANCHOR_PROMPT = """あなたはPDF解析データからExcelレイアウト仕様を生成する設計者です。

## 座標について（最重要）

入力データの各要素には **事前計算済みのExcel座標** が付与されています。
独自に座標を計算せず、必ず以下のフィールドをそのまま参照してください。

| フィールド | 意味 |
|---|---|
| `words[i]._row` | テキストのExcel行番号 |
| `words[i]._col` | テキストのExcel列番号（開始） |
| `rects[i]._row` / `_end_row` | 矩形の開始・終了行番号 |
| `rects[i]._col` / `_end_col` | 矩形の開始・終了列番号 |
| `_grid.table_cell_ranges[i]` | テーブルiの全セルのExcel範囲リスト（枠線描画用） |

## 処理手順

### Step 1: border_rect 要素の生成（テーブルセル境界）

各テーブル（インデックスi）について：
- `_grid.table_cell_ranges[i]` の各エントリを `border_rect` 要素として1つずつ出力する
- 各エントリの `row`, `end_row`, `col`, `end_col` をそのまま使用する

### Step 2: border_rect 要素の生成（rects）

各 rect の `_row`, `_col`, `_end_row`, `_end_col` をそのまま使用する。
`table_bboxes` のいずれかと座標が完全一致する rect はスキップする。

### Step 3: テキスト要素の生成

`words` を `_row` の値でグループ化し、同一 `_row` のwordを1つの `text` 要素に結合する。
- `row` = グループのwordの `_row` 値
- `col` = グループ内で最小の `_col` 値
- `content` = グループ内のwordを左から順に半角スペースで結合したテキスト
- `end_col` = `col` + `content` の文字数（概算）

## 出力フォーマット

[
  {{
    "page_number": 1,
    "elements": [
      {{
        "type": "text",
        "content": "請求書",
        "row": 2,
        "col": 20,
        "end_col": 28
      }},
      {{
        "type": "border_rect",
        "row": 8,
        "end_row": 9,
        "col": 2,
        "end_col": 17
      }},
      {{
        "type": "border_rect",
        "row": 8,
        "end_row": 9,
        "col": 18,
        "end_col": 27
      }}
    ]
  }}
]

入力データ:
{input_data}

【最重要】出力は `[` で始まり `]` で終わる純粋なJSON配列文字列のみとしてください。Markdownのコードブロック(```json等)、前後の説明文、思考プロセス、検証コメントは一切含めないでください。JSON以外の文字を1文字でも出力するとSTEP 2で処理できなくなります。"""


CODE_GEN_PROMPT = """あなたはPythonプログラミングのエキスパートです。
以下のExcelレイアウト仕様（JSON）を元に、openpyxlを使ったExcel生成スクリプトを作成してください。

## グリッド設定（必須）
全列・全行に以下の値を**ハードコード**で適用してください（mmからの再計算・変数化は禁止）：
- 列幅: `ws.column_dimensions[get_column_letter(c)].width = {excel_col_width}` （これはExcel固有の単位値。{col_width_mm}mmに相当。変換禁止）
- 行高: `ws.row_dimensions[r].height = {excel_row_height}` （これはExcel固有の単位値。{row_height_mm}mmに相当。変換禁止）
- 適用範囲: 列1〜{max_cols}、行1〜(総ページ数 × {max_rows})

## 要素の種類

入力JSONに含まれる要素は `text` と `border_rect` の2種類のみです。

## text 処理パターン（厳守・改変禁止）

以下のコードを**一字一句そのまま**使用してください。`ws.max_row` / `ws.max_column` による境界チェック、セルの存在確認、独自のガード処理は**絶対に追加しないこと**。

```python
if item["type"] == "text":
    r = item["row"] + row_offset
    try:
        cell = ws.cell(row=r, column=item["col"])
        cell.value = item["content"]
        cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=False)
    except AttributeError:
        pass
```

## border_rect 処理パターン（必須）

枠線のみを描画します（値なし）。隣接する `border_rect` の外枠が接することで内部の区切り線が自然に形成されます。

```python
from openpyxl.styles import Border, Side, Alignment
thin = Side(style='thin')

def apply_outer_border(ws, s_row, e_row, s_col, e_col):
    for r in range(s_row, e_row + 1):
        for c in range(s_col, e_col + 1):
            target = ws.cell(row=r, column=c)
            try:
                target.alignment = Alignment(horizontal='left', vertical='center', wrap_text=False)
            except AttributeError:
                pass
            top    = thin if r == s_row else None
            bottom = thin if r == e_row else None
            left   = thin if c == s_col else None
            right  = thin if c == e_col else None
            try:
                target.border = Border(top=top, bottom=bottom, left=left, right=right)
            except AttributeError:
                pass

elif item["type"] == "border_rect":
    apply_outer_border(
        ws,
        item["row"] + row_offset, item["end_row"] + row_offset,
        item["col"], item["end_col"]
    )
```

## セルスタイル（必須）
- すべてのセル: `Alignment(horizontal='left', vertical='center', wrap_text=False)`
- セル結合（ws.merge_cells）は使用禁止

## 印刷設定（必須）
- 用紙サイズ: `ws.page_setup.paperSize = 9`（A4、定数は使わず直接代入）
- 向き: `ws.page_setup.orientation = 'portrait'`（定数は使わず直接代入）
- スケーリング: fitToWidth / fitToPage は設定しない（等倍100%が大前提）
- 余白（数値は変更禁止）:
  ```python
  ws.page_margins.left = {margin_left}
  ws.page_margins.right = {margin_right}
  ws.page_margins.top = {margin_top}
  ws.page_margins.bottom = {margin_bottom}
  ```
- 印刷範囲: 入力JSONに `print_range` があればそれを使用。なければ、入力データの各要素から実際に使用されている最大行・最大列を求め、`A1:{{最大列文字}}{{最大行番号}}` を設定すること。`max_rows` や `total_pages × max_rows` を印刷範囲の計算に**使用しないこと**（コンテンツが収まらず2ページ目に溢れる原因になる）。

## 複数ページ対応
- 全ページを1シート（`wb.active`）に縦に並べる
- 2ページ目以降のrow_offset = `(page_number - 1) * {max_rows}`
- ページ境界に改ページ設定:
  ```python
  from openpyxl.worksheet.pagebreak import Break
  ws.row_breaks.append(Break(id=(page_number - 1) * {max_rows}))
  ```
- グリッド・印刷範囲の行数は必ず `総ページ数 × {max_rows}` とすること。データの最終使用行に定数を加算するバッファ処理は**禁止**。

## 技術的制約
- `True`/`False`（Python形式）を使用。JSONの`true`/`false`は禁止
- `[cite: ...]` のようなアノテーションタグは絶対に含めない（SyntaxErrorの原因）
- 出力ファイル名: `output.xlsx`

入力データ（Step 1のレイアウト仕様）:
【重要】以下の入力データに `[` で始まる JSON 配列が含まれています。JSON配列の前後に説明文や検証コメントが含まれている場合は完全に無視し、`[` から `]` までのJSON部分のみを使用してください。
{input_data}

出力はPythonコードのみをマークダウンのコードブロック（```python ... ```）で出力してください。前後の挨拶・説明文・思考プロセスは一切不要です。"""


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
【最重要】出力は、修正済みの実行可能な Python コードのみをマークダウンのコードブロック (```python ... ```) で出力してください。前後の挨拶、解説、謝罪、思考プロセスなどは一切不要です。
【最重要】出力するコード内には、`[cite: ...]` のような参照タグやアノテーション（例: ` [cite: 271]`）を絶対に含めないでください。SyntaxErrorの原因となります。純粋で実行可能なPythonコードのみを出力してください。また、コードブロックの外にいかなるテキストも記述しないでください。
【重要】`AttributeError: 'MergedCell' object attribute 'value' is read-only` を回避するため、セルへの値の代入や `merge_cells` 等の処理は、必ず `try...except AttributeError:` で囲み、重複座標エラーを握り潰す（passする）ように修正してください。"""
