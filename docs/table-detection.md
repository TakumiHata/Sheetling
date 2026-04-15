# テーブル検出とテキスト配置

pdfplumber のテーブル検出パラメータ、ページ外枠フィルタ、テーブル内テキスト配置の word 優先フォールバック戦略について説明します。

---

## pdfplumber テーブル検出パラメータ

`src/parser/pdf_extractor.py` で以下のパラメータを使用しています：

```python
_table_settings = {
    "snap_tolerance": 5,
    "snap_y_tolerance": 5,
    "join_tolerance": 5,
    "join_y_tolerance": 5,
    "edge_min_length": 5,
    "intersection_tolerance": 5,
    "intersection_x_tolerance": 5,
    "intersection_y_tolerance": 5,
}
```

| パラメータ | 値 | 役割 |
|----------|---|------|
| `snap_tolerance` | 5pt | 近接する平行エッジを同一線にまとめる距離。デフォルト3ptだと隣接セルの左辺・右辺が2本として認識され列が倍増するため大きめに設定 |
| `join_tolerance` | 5pt | 分断されたエッジを結合する距離 |
| `edge_min_length` | 5pt | この長さ未満のエッジを無視。短い装飾的線分がテーブル境界として誤検出されるのを抑制 |
| `intersection_tolerance` | 5pt | エッジ交点の許容誤差 |

---

## ページ外枠フィルタ

PDF には「ページ全体を囲む罫線」が含まれることがあり、これをテーブルとして誤検出しないための除外ロジックです。

### テーブルの外枠除外

```
除外条件: テーブル面積 ≥ ページ面積の80%  AND  セル数 ≤ 4
```

- 面積が大きく、セルが少ない = 本物のテーブルではなくページ枠
- 条件の AND により、大きくても多数のセルを持つテーブルは保持される

### 矩形の外枠除外

```
除外条件: 矩形面積 ≥ ページ面積の80%
```

- ページ全体を覆う背景矩形やページ境界を除外

---

## テーブル罫線（border_rect）の生成

`compute_grid_coords()`（`src/core/grid.py`） 内で、テーブルの `table_cells` 2D配列からセルごとに border_rect を生成します。

### 結合セルの処理

```
┌───────────────┬─────┐
│  結合セル      │  B  │
│  (bbox あり)   │     │
├───────────────┼─────┤
│  None (延長)   │  D  │
└───────────────┴─────┘
```

- **bbox がある セル**: 罫線を生成（4辺すべて `true`）
- **`None` のセル**: 結合セルの延長部分 → 罫線を生成しない

これにより結合セル内部の不要な縦線・横線が自動的に除去されます。

### 重複排除

`_is_near_duplicate()`（`src/core/layout.py`） でグリッド座標が ±1 以内の既登録要素を重複とみなします。グリッドの量子化誤差（±1行/列のずれ）で実質同一の矩形が二重登録されるのを防ぎます。

---

## テーブル外の矩形処理

テーブルに属さない `rects` は形状に応じて分類されます：

| 形状 | 判定基準 | 生成される border_rect |
|------|---------|---------------------|
| 水平罫線 | 高さがごく小さい（線の太さ程度） | `borders: {top: true}` のみ |
| 垂直罫線 | 幅がごく小さい | `borders: {left: true}` のみ |
| 通常の矩形 | 上記以外 | `borders: {top, bottom, left, right: true}` |

---

## テーブル内テキスト配置：word 優先フォールバック戦略

`_table_text_elements_from_2d()`（`src/core/layout.py`） は 2段階の戦略でテーブルセル内のテキストを配置します。

### 段階1: word 座標ベース配置（優先）

各セルの bbox に該当する word を実座標で検索し、word の座標を使って精密に配置します。

```
処理フロー:
1. セル bbox に含まれる word を検索（許容誤差: ±2.0pt）
2. word を行番号（_row）でグループ化
3. 同一 _row 内で視覚行を分割（ギャップ 3.0pt）
4. 水平ギャップで分割 → _split_by_horizontal_gap()
5. テキスト結合 → _join_word_texts()（日本語判定）
6. フォント情報（color, size, name）を保持して text 要素を生成
```

**利点**: word の実座標を使うため、セル内でのテキスト位置が正確に再現されます。

### 段階2: 2D 配列フォールバック

word が見つからない場合、`extract_tables()` の 2D 配列からテキストを取得し、セル範囲内に行を分散配置します。

**フォールバック条件**: セル bbox 内に該当する word が1つも見つからない場合のみ発動。

---

## 水平ギャップ分割 `split_by_horizontal_gap()`（`src/utils/text.py`）

同一行内のワードを水平方向の間隔で分割します。

```
分割基準:
  ギャップ = 次ワードの x0 − 前ワードの x1
  閾値 = (前ワードのfont_size + 次ワードのfont_size) / 2 × gap_factor

  ギャップ > 閾値 → 別グループに分割
```

- `gap_factor` のデフォルト値は `2.0`
- フォントサイズに比例した動的な閾値を使用するため、大きい文字は広いギャップでも連続、小さい文字は狭いギャップでも分割されます

### 使用箇所

- テーブル内テキスト配置（`_table_text_elements_from_2d()` 内）
- テーブル外テキスト配置（`_collect_text_elements()` 内）

いずれも `src/core/layout.py` に定義されています。

---

## テキスト結合ルール

### `join_word_texts()`（`src/utils/text.py`）

```
日本語を含む → スペースなしで結合: "請求" + "書" → "請求書"
英数字のみ  → スペース区切り:      "Total" + "Amount" → "Total Amount"
```

### `has_japanese()`（`src/utils/text.py`）

以下の Unicode 範囲で日本語を判定：
- `U+3040–U+30FF`: ひらがな・カタカナ
- `U+4E00–U+9FFF`: CJK 統合漢字
- `U+FF00–U+FFEF`: 全角英数・記号
