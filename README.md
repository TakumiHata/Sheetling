# Sheetling

PDFを解析し、自動生成したPythonコードを実行することで、**任意のPDFレイアウトを維持したままA4/A3方眼Excelに変換**するツールです。
テーブルの列ズレのない高精度な方眼紙レイアウトへの変換を、LLMを最小限に抑えた完全自動パイプラインで実現します。

---

## 仕組み

```
PDF → [auto] 解析 → グリッド座標計算 → レイアウトJSON生成 → _gen.py 生成 → Excel出力
                                                                    ↓
                                              （任意）ビジョンLLMによる視覚的検証
                                                                    ↓
                                             [correct] 修正適用 → Excel再生成
```

| コマンド | 実行者 | 内容 |
|---------|--------|------|
| `auto` | スクリプト | PDF解析 → グリッド座標計算 → レイアウトJSON自動生成 → `_gen.py` 生成 → Excel出力。視覚的検証プロンプトも出力 |
| `correct` | 人間 + ビジョンLLM | PDFページ画像 + 検証プロンプトをAIチャットに投入。修正指示JSONを適用し Excel を再生成 |

---

## セットアップ

### Docker環境を使用する場合
```bash
docker compose up -d --build
```

### ローカルのPython環境を使用する場合
```bash
pip install -r requirements.txt
```

---

## 実行手順

### auto：PDF → Excel 自動生成

`data/in/` に対象のPDFを置いて実行します：

```bash
# 全PDF一括処理
python -m src.main auto [--grid-size <size>]

# 特定PDFのみ
python -m src.main auto --pdf data/in/sample.pdf [--grid-size <size>]
```

| 処理 | 内容 |
|------|------|
| PDF解析 | `pdfplumber` でテキスト・罫線・矩形を抽出し、Excel グリッド座標を計算 |
| レイアウトJSON生成 | `table_border_rects` / `rects` / `words` から直接生成（LLM不要） |
| 座標検証・補正 | 整合性チェック・重複除去・クランプをスクリプトで確実に適用 |
| テキスト補完 | 抽出データと照合して欠落テキストを補完 |
| `_gen.py` 生成 | テンプレートから Excel生成スクリプトを生成 |
| 視覚的検証プロンプト生成 | ページごとに `_visual_review_page{N}.txt` を出力 |
| Excel出力 | `_gen.py` を実行して `_Python版.xlsx` を生成 |

実行後、`data/out/<pdf_name>/` に以下が生成されます：

```text
data/out/<pdf_name>/
├── <pdf_name>_extracted.json     # PDFから抽出した生データ（グリッド座標付き）
├── <pdf_name>_grid_params.json   # グリッドパラメータ（罫線後処理用）
├── <pdf_name>_layout.json        # レイアウトJSON（_gen.py が参照）
├── <pdf_name>_gen.py             # 自動生成された Excel生成スクリプト
├── <pdf_name>_Python版.xlsx      # 生成された Excel ファイル
└── prompts/
    ├── page_1/
    │   ├── <pdf_name>_page1.png                         # PDFページ画像
    │   ├── <pdf_name>_visual_review_page1.txt           # 視覚的検証プロンプト
    │   └── <pdf_name>_visual_corrections_page1.json     # LLM修正指示（ユーザーが編集）
    └── page_2/
        └── ...
```

---

### 視覚的検証（任意）：ビジョンLLMで再現度を高める

`auto` 実行後、再現度をさらに高めたい場合にオプションで実施します。

**手順：**

1. PDFを開き、対象ページをスクリーンショットまたは画像として保存する
2. 社内AIチャット（画像入力対応）に以下を貼り付ける：
   - PDFページの画像
   - `prompts/<pdf_name>_visual_review_page{N}.txt` の内容
3. LLMが出力した修正指示JSON（`{"corrections": [...]}` 形式）を以下に保存する：
   ```
   data/out/<pdf_name>/prompts/<pdf_name>_visual_corrections.json
   ```
4. `correct` コマンドで修正を適用する（次節）

**LLMが検出できる問題：**
- 欠落しているテキスト
- 位置がズレているテキスト
- 欠落している罫線・枠
- 不要な罫線

> [!NOTE]
> スクリプトが計算した精密な座標にLLMの視覚的な判断を組み合わせることで、座標精度と再現性の両方を確保できます。

---

### correct：視覚的検証の修正を適用

```bash
# 特定PDFのみ
python -m src.main correct --pdf sample

# data/out/ 以下の全 *_visual_corrections.json を一括処理
python -m src.main correct
```

`_visual_corrections_page{N}.json` の修正指示を `_layout.json` に適用し、`_gen.py` を再生成して Excel を出力します。

**`_visual_corrections.json` の形式：**

```json
{
  "corrections": [
    {"action": "add_text",     "page": 1, "row": 5, "col": 3, "content": "追加テキスト"},
    {"action": "fix_text",     "page": 1, "row": 3, "col": 5, "new_row": 4, "new_col": 6},
    {"action": "add_border",   "page": 1, "row": 3, "end_row": 8, "col": 2, "end_col": 15,
                               "borders": {"top": true, "bottom": true, "left": true, "right": true}},
    {"action": "remove_border","page": 1, "row": 3, "end_row": 5, "col": 2, "end_col": 8}
  ]
}
```

---

## CLIリファレンス

```
python -m src.main <command> [options]
```

| command | 説明 |
|---------|------|
| `auto` | PDF → Excel 自動生成（解析 → レイアウトJSON生成 → `_gen.py` 生成 → Excel出力） |
| `correct` | ビジョンLLMの修正指示（`_visual_corrections.json`）を適用して Excel を再生成 |

| オプション | 対象command | 説明 |
|-----------|------------|------|
| `--pdf <path>` | `auto`, `correct` | 処理対象PDFのパスまたはPDF名（省略時は全対象を処理） |
| `--grid-size <size>` | `auto` | 方眼サイズ: `small`（デフォルト）/ `medium` / `large` |

---

## 方眼サイズ（`--grid-size`）の詳細仕様

絶対等倍でのA4出力を保証するため、各サイズには数学的に計算された最大グリッド数が設定されています。A3など他の用紙サイズには動的に比例計算して対応します。

| グリッドサイズ | 方眼の大きさ (mm) | 最大列数 (A4縦) | 最大行数 (A4縦) | Excel設定値 (幅/高さ) | デフォルトフォント |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **`small`** | **約 4.0 mm** | **62 列** | **76 行** | 幅: 1.45 / 高さ: 11.34 | 7pt |
| **`medium`** | **約 6.0 mm** | **36 列** | **50 行** | 幅: 2.53 / 高さ: 17.01 | 9pt |
| **`large`** | **約 8.0 mm** | **26 列** | **38 行** | 幅: 3.61 / 高さ: 22.68 | 11pt |

> [!NOTE]
> これらの数値は、A4用紙の物理的な印字可能領域（余白を除いた約180mm x 270mm）を基準に、各サイズの正方形グリッドを敷き詰めた際の理論限界値です。A3等の他の用紙サイズはページ寸法に比例して自動計算されます。

---

## プロジェクト構成

```text
Sheetling/
├── src/
│   ├── main.py              # CLI エントリポイント（auto / correct）
│   ├── core/
│   │   └── pipeline.py      # パイプライン全体の制御・座標計算・自動生成・ボーダー後処理
│   ├── parser/
│   │   └── pdf_extractor.py # PDFデータ抽出 (pdfplumber)
│   ├── templates/
│   │   └── prompts.py       # コード生成テンプレート・視覚的検証プロンプト・グリッドサイズ設定
│   └── utils/
│       └── logger.py        # ログ出力管理
├── data/
│   ├── in/                  # 入力PDFディレクトリ
│   └── out/                 # 出力ディレクトリ（解析結果・Excel）
├── Dockerfile
├── docker-compose.yml
└── requirements.txt
```

---

## 使用パッケージ

| パッケージ | 用途 |
|-----------|------|
| `pdfplumber` | PDF内のテキスト、表、罫線の詳細な座標情報を抽出 |
| `openpyxl` | Excelファイルの生成、セルのスタイリング、印刷範囲設定 |

---

## エラー発生時

`auto` / `correct` コマンドで `_gen.py` の実行に失敗した場合、`prompts/<pdf_name>_prompt_error_fix.txt` にエラー修正用プロンプトが出力されます。その内容をAIチャットに投入してコードを修正し、`_gen.py` を上書き保存してから再度 `correct` を実行してください。

---

## 開発メモ

- `src/` と `data/` はコンテナにボリュームマウントされており、ホスト側の変更が即時反映されます。
- 罫線の後処理は `_extracted.json` と `_grid_params.json` が存在する場合のみ実行されます（`auto` 実行時に自動生成）。

### アーキテクチャ設計の考え方

#### スクリプトが担う処理（`auto` コマンド）
- **座標マッピング**：`_row` / `_col` は解析時に確定計算済みのためLLM不要
- **border_rect 生成**：`table_border_rects` / `rects` のフィールドをそのまま変換
- **text 生成**：`words` を `(_row, _col)` でグループ化するだけ
- **座標検証・補正**：整合性チェック・重複除去・クランプはすべて決定論的処理
- **コード生成**：`_gen.py` は固定テンプレート + パラメータ差し込みで生成

#### ビジョンLLMが担う処理（`correct` コマンド、任意）
- **視覚的な差分検出**：スクリプトが計算した精密な座標 × LLMの目視確認で再現度を向上
- **意味的な判断が必要なケース**：複雑なレイアウトでスクリプトが見落とした欠落要素の検出

### TODO / 改善案

#### 複数ページ・データ量大への対応

`auto` コマンドはページ全体を一括処理するため、ビジョン検証プロンプトはページごとに分割して出力されます。PDFのページ数が多い場合は、代表ページのみを検証することで作業量を削減できます。

**将来対応案：** `auto` コマンドに `--pages 1-3` のようなオプションを追加して、処理対象ページを絞り込む。
