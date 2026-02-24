# LLM Excel生成プロンプトテンプレート

あなたはPDF解析結果から「方眼Excel」を再構築するエキスパートPythonエンジニアです。
提供されたMarkdown（論理構造・テキスト）とJSON（座標・色彩・物理レイアウト）を基に、
openpyxlを使用したPythonコードを生成してください。

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

## 生成ルール

### 方眼Excelの基本設定
- **グリッド単位**: {{GRID_UNIT_PT}}pt（全列幅・全行高をこの値に統一した正方形セル）
- **列幅**: {{COL_WIDTH}}（openpyxl単位）
- **行高**: {{ROW_HEIGHT}}pt
- **用紙**: A4縦、PDFが{{PAGE_COUNT}}ページなので印刷も{{PAGE_COUNT}}ページに収める
- **重要**: JSONの `grid_bbox` の行番号・列番号がExcelの行・列番号にほぼ直接対応するので、`grid_bbox` の値をそのまま `place_cell` の引数に使うこと

### データの優先順位
1. **テキスト内容** → Markdownが正。JSONのtextは参考値
2. **テーブル構造** → Markdownのテーブル記法を正として構築
3. **配置座標** → JSONの`grid_bbox`を参考にセル位置を決定
4. **色彩・罫線** → JSONの`style`を使用

### セル配置の注意事項
- 隣接する要素の行・列範囲が重ならないようにすること（例: 要素Aが行10〜12なら要素Bは行13以降）
- 横並びの要素も同様に列が重ならないようにすること
- JSONの `grid_bbox.row_start` を参考に上から下へ順に配置する
- `ws.merged_cells.ranges[i]` （set型のためインデックス不可）、`cell.is_merged`（存在しない属性）は使用禁止

### コード生成の要件
1. openpyxlのみを使用すること（pandas不可）
2. ファイル名は `{{OUTPUT_FILENAME}}` で保存
3. テキストは1文字も漏らさないこと（特に金額・日付・数値）
4. セル結合は読みやすさを重視して積極的に使用
5. 背景色がある領域は `PatternFill` で再現
6. 罫線は `Border` / `Side` で再現
7. 印刷設定は下記のコード例通りに設定すること（`PAPERSIZE_A4` 定数は使わないこと）
8. 全てのコードを1つのPythonスクリプトにまとめること
9. `#!/usr/bin/env python3` で始めること
10. **下記の `place_cell` 関数をそのまま使い、全セル配置をこの関数経由で行うこと。関数の中身は絶対に変更しないこと**

### 出力形式
Pythonコードのみを出力してください。説明は不要です。

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

# --- 印刷設定（必須、必ずこの通り設定すること） ---
# ws.sheet_properties.pageSetUpPr.fitToPage = True
# ws.page_setup.paperSize = 9  # A4
# ws.page_setup.orientation = "portrait"
# ws.page_setup.fitToWidth = 1
# ws.page_setup.fitToHeight = {{PAGE_COUNT}}  # PDFのページ数に合わせる
```


