# LLM Excel生成プロンプトテンプレート

あなたはPDF解析結果から「方眼Excel」を再構築するエキスパートPythonエンジニアです。
提供された**配置命令リスト**に従い、openpyxlを使用したPythonコードを生成してください。

---

## 入力データ

### Markdown（論理構造・テキスト ＝ 内容の正解）

```markdown
{{MARKDOWN_CONTENT}}
```

### JSON（座標・色彩・物理レイアウト ＝ 配置の参考）

```json
{{JSON_CONTENT}}
```

---

## 配置命令リスト（最重要 — そのままコードに展開すること）

以下の命令リストは、JSONの座標データから**事前計算済み**です。
座標変換（半開区間→閉区間）、テーブル構造へのスナップ、重複チェックはすべて完了しています。

**⚠ 命令リストの座標を変更してはいけません。そのまま使ってください。**

{{PLACEMENT_COMMANDS}}

---

## テーブル構造（参考情報）

{{TABLE_STRUCTURE_SUMMARY}}

---

## 生成ルール

### 方眼Excelの基本設定
- **グリッド単位**: {{GRID_UNIT_PT}}pt（全列幅・全行高をこの値に統一した正方形セル）
- **列幅**: {{COL_WIDTH}}（openpyxl単位）
- **行高**: {{ROW_HEIGHT}}pt
- **用紙**: A4縦、PDFが{{PAGE_COUNT}}ページなので印刷も{{PAGE_COUNT}}ページに収める

### コード生成の最重要ルール

1. **配置命令リストをそのまま展開すること** — 座標や引数を自分で計算・変更しないこと
2. **line要素の罫線描画**には `place_cell` を使わず、配置命令リストの `draw_table_borders` 呼び出しをそのまま使うこと
3. **セル座標の重複は厳禁** — 配置命令リストは事前に重複チェック済みなので、追加の `place_cell` 呼び出しを独自に入れないこと

### データの優先順位
1. **配置命令リスト** → 座標・スタイルの正。そのまま使う
2. **テキスト内容** → Markdownが正。配置命令のvalueと異なる場合はMarkdownを優先
3. **テーブル構造** → 罫線位置の参考

### コード生成の要件
1. openpyxlのみを使用すること（pandas不可）
2. ファイル名は `{{OUTPUT_FILENAME}}` で保存
3. テキストは1文字も漏らさないこと（特に金額・日付・数値）
4. セル結合は配置命令の座標範囲に基づいて自動適用される
5. 背景色がある領域は `PatternFill` で再現
6. 罫線は `Border` / `Side` で再現
7. 全てのコードを1つのPythonスクリプトにまとめること
8. `#!/usr/bin/env python3` で始めること
9. **下記の `place_cell` 関数をそのまま使い、配置命令をこの関数経由で展開すること**
10. **下記の `draw_table_borders` 関数をそのまま使うこと**
11. `ws.merged_cells.ranges[i]` （set型のためインデックス不可）、`cell.is_merged`（存在しない属性）は使用禁止

### 出力形式
Pythonコードのみを出力してください。説明は不要です。

以下のコード骨格に従ってください：

```python
#!/usr/bin/env python3
"""方眼Excel生成スクリプト - {{PDF_NAME}}"""
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Border, Side, Font
from openpyxl.utils import get_column_letter

def place_cell(ws, r1, c1, r2, c2, value="", font=None, alignment=None, fill=None, border=None):
    """セルに値・スタイルを設定し結合する。重複する既存結合は自動解除される。
    この関数は変更せずそのまま使うこと。"""
    # 重複する既存の結合範囲を自動解除
    overlapping = [mr.coord for mr in ws.merged_cells.ranges
                   if mr.min_row <= r2 and mr.max_row >= r1
                   and mr.min_col <= c2 and mr.max_col >= c1]
    for coord in overlapping:
        ws.unmerge_cells(coord)
    # 左上セルに値・スタイルを設定
    cell = ws.cell(row=r1, column=c1, value=value)
    if font:
        cell.font = font
    if alignment:
        cell.alignment = alignment
    if fill:
        cell.fill = fill
    # 罫線は全セルに適用（結合後はMergedCellとなり設定不可のため）
    if border:
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                ws.cell(row=r, column=c).border = border
    # セル結合
    if r2 > r1 or c2 > c1:
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)

def draw_table_borders(ws, v_cols, h_rows, row_min, row_max, col_min, col_max):
    """テーブルの罫線を描画する。place_cellとの競合を避けるため直接cell.borderを設定する。"""
    side_thin = Side(border_style="thin", color="000000")
    # 縦線
    for c in v_cols:
        for r in range(row_min, row_max + 1):
            cell = ws.cell(row=r, column=c)
            existing = cell.border
            cell.border = Border(
                left=side_thin,
                right=existing.right,
                top=existing.top,
                bottom=existing.bottom
            )
    # 横線
    for r in h_rows:
        for c in range(col_min, col_max + 1):
            cell = ws.cell(row=r, column=c)
            existing = cell.border
            cell.border = Border(
                left=existing.left,
                right=existing.right,
                top=side_thin,
                bottom=existing.bottom
            )

def main():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # --- 1. グリッド設定（方眼紙） ---
    for col_idx in range(1, {{GRID_COLS}} + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = {{COL_WIDTH}}
    for row_idx in range(1, {{GRID_ROWS}} + 1):
        ws.row_dimensions[row_idx].height = {{ROW_HEIGHT}}

    # --- 2. rect要素（背景色） ---
    # 配置命令リストの「rect要素の配置命令」をここにそのまま展開する

    # --- 3. text要素（テキスト・フォント・配置） ---
    # 配置命令リストの「text要素の配置命令」をここにそのまま展開する
    # ⚠ 座標は事前計算済み。変更しないこと

    # --- 4. テーブル罫線（place_cellを使わないこと） ---
    # 配置命令リストの「テーブル罫線の描画命令」をここにそのまま展開する

    # --- 5. 印刷設定（必須、必ずこの通り設定すること） ---
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.paperSize = 9  # A4
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = {{PAGE_COUNT}}  # PDFのページ数に合わせる

    wb.save("{{OUTPUT_FILENAME}}")

if __name__ == "__main__":
    main()
```
