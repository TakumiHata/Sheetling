# correct ワークフロー

correct コマンドの詳細フロー、プロンプトの設計意図、LLM への渡し方、corrections JSON の仕様を説明します。

---

## 全体フロー

```
auto 実行（前提）
    │
    ├─→ PDF ページ画像 (.png)
    ├─→ 差分オーバーレイ画像 (_diff_pageN.png)  ← PDF原本 + 赤線オーバーレイ
    ├─→ ID 付き罫線リスト (_edges_pageN.json)   ← inclusive 座標
    ├─→ フェーズ1プロンプト (_phase1_prompt_pageN.txt)
    ├─→ フェーズ1修正 JSON (_phase1_corrections_pageN.json)  ← ユーザーが記入
    ├─→ フェーズ2プロンプト (_phase2_prompt_pageN.txt)
    └─→ フェーズ2修正 JSON (_phase2_corrections_pageN.json)  ← ユーザーが記入
         │
         ▼  correct コマンド実行
         │  phase1 + phase2 の corrections をマージして適用
    apply_corrections()
         │
    rerender_after_corrections()
         │
    修正版 Excel (.xlsx)
```

---

## 2フェーズ設計の意図

| フェーズ | タスク | 難易度 | 出力形式 |
|--------|--------|--------|---------|
| フェーズ1 | 余分な赤線の ID を列挙する | 低（Yes/No 判定） | `remove_edges` + ID リスト |
| フェーズ2 | 不足している黒線を座標で追加する | 高（座標推定） | `add_edge` + inclusive 座標 |

削除は ID 番号を指定するだけで完結するため LLM の誤りが少ない。
追加は座標推定を伴うため別フェーズに分離し、LLM の認知負荷を下げる。

---

## 入力画像の説明

`auto` 実行時に生成される `_diff_pageN.png`:

| 要素 | 内容 |
|------|------|
| 黒線 | PDF 原本の罫線 |
| 半透明赤線 | 現在のプレビュー罫線（各線に ID ラベル付き） |

LLM は「赤線が黒線と一致するか」をこの 1 枚で判断する。

---

## フェーズ1: 削除判定

### プロンプトの役割

`VISUAL_PHASE1_PROMPT`（`src/templates/prompts.py`）は以下を渡す:
- 差分オーバーレイ画像
- ID 付き罫線リスト JSON（edges JSON）

LLM は各赤線を確認し、PDF に対応する黒線がない場合だけ ID を返す。

### edges JSON の座標系（inclusive）

edges JSON の `col_end` / `row_end` は **inclusive**（最終セルの番号そのもの）:

```json
{"edges": [
  {"id": 1, "type": "H", "row": 5, "col_start": 3, "col_end": 12},
  {"id": 2, "type": "V", "col": 3, "row_start": 1, "row_end": 8}
]}
```

col 3〜12 の水平線なら `col_end=12`（13 ではない）。

### 出力形式

```json
{"corrections": [
  {"action": "remove_edges", "page": 1, "ids": [3, 17, 42]}
]}
```

差異なしの場合: `{"corrections": []}`

---

## フェーズ2: 追加判定

### プロンプトの役割

`VISUAL_PHASE2_PROMPT`（`src/templates/prompts.py`）は以下を渡す:
- フェーズ1と同じ差分オーバーレイ画像
- グリッド情報（行数・列数・セル寸法）

LLM は PDF に黒線があるが赤線が対応していない箇所を特定し、inclusive 座標で返す。

### 座標の規約

| 記法 | 意味 |
|------|------|
| `H, row=N` | N 行目の上辺にある水平線 |
| `V, col=N` | N 列目の左辺にある垂直線 |
| `col_start / col_end` | 両端のセル番号（**両端 inclusive**） |
| `row_start / row_end` | 両端のセル番号（**両端 inclusive**） |

例: col 3〜12 の H 線 → `col_start=3, col_end=12`（13 ではない）

### 出力形式

```json
{"corrections": [
  {"action": "add_edge", "page": 1, "type": "H", "row": 5, "col_start": 3, "col_end": 12},
  {"action": "add_edge", "page": 1, "type": "V", "col": 3, "row_start": 1, "row_end": 8}
]}
```

---

## corrections JSON の仕様

### add_edge

| フィールド | 必須 | 説明 |
|----------|------|------|
| `action` | はい | `"add_edge"` |
| `page` | はい | ページ番号 |
| `type` | はい | `"H"`（水平）または `"V"`（垂直） |
| `row` | H のみ | 水平線の行番号 |
| `col_start` | H のみ | 左端セル番号（inclusive） |
| `col_end` | H のみ | 右端セル番号（inclusive） |
| `col` | V のみ | 垂直線の列番号 |
| `row_start` | V のみ | 上端セル番号（inclusive） |
| `row_end` | V のみ | 下端セル番号（inclusive） |

コード側で `col_end` / `row_end` に +1 して内部の exclusive 形式に変換する。

### remove_edges

```json
{"action": "remove_edges", "page": 1, "ids": [3, 17, 42]}
```

| フィールド | 必須 | 説明 |
|----------|------|------|
| `action` | はい | `"remove_edges"` |
| `page` | はい | ページ番号 |
| `ids` | はい | 削除するランの ID 配列（単一でも `[7]` の形式） |

---

## correct コマンドの動作

`_apply_corrections_for_pair`（`src/main.py`）は以下の優先順でファイルを検索:

1. `page_*/*_phase1_corrections_pageN.json` + `page_*/*_phase2_corrections_pageN.json`
2. フラット配置: `*_phase1_corrections_pageN.json` + `*_phase2_corrections_pageN.json`
3. 旧フォーマット: `page_*/*_visual_corrections_pageN.json`（後方互換）

全ファイルの `corrections` リストをマージして一括適用する。

---

## Excel 再生成

`rerender_after_corrections()` は以下を実行:

1. 修正済み `{pdf_name}_{grid_size}_layout.json` を読み込み
2. `{pdf_name}_{grid_size}_grid_params.json` を読み込み
3. `render_layout_to_xlsx()`（`src/renderer/excel.py`）で Excel を再生成

---

## 関連ドキュメント

- [アーキテクチャ](architecture.md) — パイプライン全体のデータフロー・モジュール構成
- [グリッドシステム](grid-system.md) — グリッド座標・コンテンツ境界の計算ロジック
- [チューニングガイド](tuning-guide.md) — GRID_SIZES 調整・トラブルシューティング
