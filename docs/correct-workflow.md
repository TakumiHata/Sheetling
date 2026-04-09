# correct ワークフロー

correct コマンドの詳細フロー、プロンプトの設計意図、LLM への渡し方、corrections JSON の仕様を説明します。

---

## 全体フロー

```
auto 実行（前提）
    │
    ├─→ PDF ページ画像 (.png)
    ├─→ 罫線プレビュー画像 (.png)
    ├─→ 検証プロンプト (.txt)
    └─→ 空テンプレート (_visual_corrections.json)
         │
         ▼  ユーザーがビジョンLLMに投入
         │
    修正指示 JSON をファイルに保存
         │
         ▼  correct コマンド実行
         │
    apply_corrections()
         │  修正指示を layout JSON に適用
         ▼
    rerender_after_corrections()
         │  修正済み layout JSON から Excel を再生成
         ▼
    修正版 Excel (.xlsx)
```

---

## 検証プロンプトの設計

`VISUAL_REVIEW_PROMPT`（`src/templates/prompts.py`）は以下の情報をビジョン LLM に伝えます：

### 入力画像の説明

| 画像 | 内容 |
|------|------|
| 画像1 | PDF ページの原本画像 |
| 画像2 | 自動生成した罫線プレビュー画像 |

プレビュー画像の要素：
- **薄いグレー線** (`#E0E0E0`): グリッド背景線 → 無視対象
- **太い黒線**: 実際に Excel に描画される罫線
- **赤いラベル** (RGB: 200, 0, 0): 5行/5列間隔で表示される座標数値。JSON の `row` / `col` に直接対応
- **グレーアウト領域**: コンテンツ範囲外

### 判定基準

| 状況 | アクション |
|------|----------|
| PDF にあるがプレビューにない罫線 | `add_border` |
| プレビューにあるが PDF にない罫線 | `remove_border` |
| 範囲がずれている罫線 | `remove_border` + `add_border` |

無視すべき差異：テキスト・フォント差異、グリッド背景線、薄い罫線・飾り線・影

### 座標範囲の制約

プロンプトにはテンプレート変数でコンテンツ有効範囲が埋め込まれます：

```
row: 1 〜 {content_max_row}
col: 1 〜 {content_max_col}
```

LLM がこの範囲外の座標を指定することを防止します。

---

## corrections JSON の仕様

### add_border

```json
{
  "action": "add_border",
  "page": 1,
  "row": 3,
  "end_row": 8,
  "col": 2,
  "end_col": 15,
  "borders": {"top": true, "bottom": true, "left": true, "right": true}
}
```

| フィールド | 必須 | 説明 |
|----------|------|------|
| `action` | はい | `"add_border"` |
| `page` | はい | ページ番号 |
| `row` | はい | 開始行 |
| `end_row` | はい | 終了行 |
| `col` | はい | 開始列 |
| `end_col` | はい | 終了列 |
| `borders` | いいえ | 辺の指定。省略時は全辺 `true`。各辺を独立して指定可能 |

`borders` の例：下線のみ追加する場合

```json
{"top": false, "bottom": true, "left": false, "right": false}
```

### remove_border

```json
{
  "action": "remove_border",
  "page": 1,
  "row": 3,
  "end_row": 8,
  "col": 2,
  "end_col": 15
}
```

| フィールド | 必須 | 説明 |
|----------|------|------|
| `action` | はい | `"remove_border"` |
| `page` | はい | ページ番号 |
| `row` | はい | 開始行 |
| `end_row` | はい | 終了行 |
| `col` | はい | 開始列 |
| `end_col` | はい | 終了列 |

### フィールド名の互換性

コードは `row_end` / `col_end` も受け付けますが、プロンプトでは `end_row` / `end_col` を正式フィールド名として指定しています。

### 差異なしの場合

```json
{"corrections": []}
```

---

## 安全機構

### 1. add_border のコンテンツ範囲クランプ

`apply_corrections()` は `add_border` の `end_row` / `end_col` を既存コンテンツの最大範囲内にクランプします。

```python
_end_row = min(_end_row, content_bounds["max_row"])
_end_col = min(_end_col, content_bounds["max_col"])
```

`content_bounds` は既存の `border_rect` 要素の `end_row` / `end_col` の最大値から計算されます。

**目的**: LLM が範囲外の座標を指定しても、Excel 上ではみ出すことを防止。

### 2. remove_border の包含判定

`remove_border` は指定範囲に **完全に包含される** `border_rect` のみを削除します。

```
判定: e.row >= 指定row AND e.end_row <= 指定end_row
      AND e.col >= 指定col AND e.end_col <= 指定end_col
```

**重要**: overlap（重複）判定ではなく containment（包含）判定を使用します。

```
例:
  外枠: (1, 1) → (10, 10)
  削除指定: (3, 3) → (8, 8)
  結果: 外枠は保持される（完全包含されていないため）
```

overlap 判定だと外枠など大きなボーダーが巻き添えで削除されてしまうため、この設計が採用されています。

---

## Excel 再生成

`rerender_after_corrections()` は以下を実行します：

1. 修正済み `{pdf_name}_{grid_size}_layout.json` を読み込み
2. `{pdf_name}_{grid_size}_grid_params.json` を読み込み
3. `_render_layout_to_xlsx()` で Excel を再生成

レイアウト JSON 全体を再レンダリングするため、修正箇所だけでなく既存のテキスト・罫線もすべて含めた完全な Excel が出力されます。

---

## 罫線プレビュー画像の生成

`_generate_border_preview()` は `auto` 実行時に以下を描画します：

| 要素 | 描画内容 |
|------|---------|
| グリッド背景線 | 薄いグレー (`#E0E0E0`) 1px |
| 罫線 | 太い黒線（幅は `min(cell_w, cell_h) / 7` 以上 2px） |
| 座標ラベル | 赤 (200, 0, 0) で 5行/5列間隔 |
| コンテンツ範囲外 | グレー (210, 210, 210) で塗りつぶし |

コンテンツ境界情報（`content_bounds`）を使用して、PDF 画像上の正しい位置に罫線を描画します。これにより、PDF 原本画像とプレビュー画像の罫線位置が一致し、LLM が正確に差分を検出できます。
