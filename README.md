# Sheetling

PDFを解析し、スクリプトが自動生成したPythonコードを実行することで、**任意のPDFレイアウトを維持したままA4/A3方眼Excelに変換**するツールです。
テーブルの列ズレのない高精度な方眼紙レイアウトへの変換を、LLMを最小限に抑えた完全自動パイプラインで実現します。

---

## 仕組み（2フェーズ構成）

```
PDF → [Phase 1] 解析・座標計算
         ↓
      [auto] レイアウトJSON自動生成 + _gen.py 自動生成
         ↓
      [任意] ビジョンLLMによる視覚的検証・修正
         ↓
      [Phase 3] Excel描画 + 確定的ボーダー後処理
```

| フェーズ | コマンド | 実行者 | 内容 |
|----------|---------|--------|------|
| Phase 1 | `extract` | スクリプト | `pdfplumber` でPDFを解析。Excel グリッド座標を事前計算し、抽出データを `data/out/<pdf_name>/` に出力 |
| 自動生成 | `auto` | スクリプト | 抽出データから直接レイアウトJSONと `_gen.py` を生成。視覚的検証プロンプトも出力 |
| 視覚的検証 | `correct` | 人間 + ビジョンLLM | PDFページ画像 + 検証プロンプトをAIチャットに投入。修正指示JSONを `_gen.py` に反映 |
| Phase 3 | `generate` | スクリプト | `_gen.py` を実行してExcelを出力。罫線を確定的に後処理で上書き適用 |

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

### Phase 1：PDF解析 & 座標計算

`data/in/` に対象のPDFを置いて実行します：

```bash
# 全PDF一括処理
python -m src.main extract [--grid-size <size>]

# 特定PDFのみ
python -m src.main extract --pdf data/in/sample.pdf [--grid-size <size>]
```

**オプション：**

| オプション | デフォルト | 説明 |
|-----------|----------|------|
| `--grid-size` | `small` | 方眼サイズ（`small` / `medium` / `large`） |
| `--pdf` | なし（全PDF） | 処理対象PDFのパス |

実行後、`data/out/<pdf_name>/` に以下が生成されます：

```text
data/out/<pdf_name>/
├── <pdf_name>_extracted.json      # PDFから抽出した生データ（グリッド座標付き）
├── <pdf_name>_grid_params.json    # グリッドパラメータ（Phase 3 罫線後処理用）
├── <pdf_name>_gen.py              # (空) auto コマンドで自動生成される
└── prompts/
    ├── <pdf_name>_prompt_step1.txt    # 旧来フロー用: STEP 1 プロンプト（参照用）
    ├── <pdf_name>_prompt_step1_5.txt  # 旧来フロー用: STEP 1.5 プロンプト（参照用）
    └── <pdf_name>_prompt_step2.txt    # 旧来フロー用: STEP 2 プロンプト（参照用）
```

---

### auto：レイアウト自動生成（step1 + step1.5 + fill + step2 を一括自動化）

```bash
# 全PDF一括処理
python -m src.main auto [--grid-size <size>]

# 特定PDFのみ
python -m src.main auto --pdf data/in/sample.pdf [--grid-size <size>]
```

`extract` 済みの抽出データをもとに、以下をすべて自動で行います：

| 処理 | 内容 |
|------|------|
| レイアウトJSON生成 | `_extracted.json` の `table_border_rects` / `rects` / `words` から直接生成（旧 STEP 1 + 1.5 相当） |
| 座標検証・補正 | 座標整合性チェック・重複除去・クランプをスクリプトで確実に適用 |
| テキスト補完 | `_extracted.json` の `words` と照合して欠落テキストを補完（旧 `fill` 相当） |
| `_gen.py` 生成 | テンプレートから Excel生成スクリプトを生成。JSONはランタイムで読み込む設計 |
| 視覚的検証プロンプト生成 | ページごとに `_visual_review_page{N}.txt` を出力 |

実行後、`data/out/<pdf_name>/` に以下が追加されます：

```text
data/out/<pdf_name>/
├── <pdf_name>_gen.py                        # 自動生成された Excel生成スクリプト
└── prompts/
    ├── <pdf_name>_step1_5_input.json        # 自動生成されたレイアウトJSON
    ├── <pdf_name>_step1_5_output.json       # テキスト補完済みレイアウトJSON
    └── <pdf_name>_visual_review_page1.txt   # 視覚的検証プロンプト（ページごと）
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

`_visual_corrections.json` の修正指示を `_step1_5_output.json` に適用し、`_gen.py` を再生成します。

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

### Phase 3：Excel生成

```bash
# data/out/ 以下の全 *_gen.py を一括処理
python -m src.main generate
```

`data/out/<pdf_name>/<pdf_name>_Python版.xlsx` が生成されます。

**処理内容：**
- `_gen.py` を実行して Excel を生成（`_step1_5_output.json` をランタイムで読み込む）
- `_extracted.json` と `_grid_params.json` を参照し、罫線を確定的に後処理で上書き適用
- エラー発生時は `prompts/<pdf_name>_prompt_error_fix.txt` にエラー修正用プロンプトを出力

---

## CLIリファレンス

```
python -m src.main <phase> [options]
```

| phase | 説明 |
|-------|------|
| `extract` | Phase 1: PDF解析 & グリッド座標計算 |
| `auto` | step1 + step1.5 + fill + step2 を完全自動化。レイアウトJSON・`_gen.py`・視覚的検証プロンプトを生成 |
| `correct` | ビジョンLLMの修正指示（`_visual_corrections.json`）を適用して `_gen.py` を再生成 |
| `fill` | 旧来フロー用: STEP 1.5 出力のテキスト補完 & STEP 2 プロンプト更新 |
| `generate` | Phase 3: `_gen.py` を実行してExcel出力 |

| オプション | 対象phase | 説明 |
|-----------|----------|------|
| `--pdf <path>` | `extract`, `auto`, `fill`, `correct` | 処理対象PDFのパスまたはPDF名（省略時は全対象を処理） |
| `--grid-size <size>` | `extract`, `auto` | 方眼サイズ: `small`（デフォルト）/ `medium` / `large` |

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
│   ├── main.py              # CLI エントリポイント（extract / auto / correct / fill / generate）
│   ├── core/
│   │   └── pipeline.py      # フェーズ全体のフロー制御・座標計算・自動生成・ボーダー後処理
│   ├── parser/
│   │   └── pdf_extractor.py # Phase 1: PDFデータ抽出 (pdfplumber)
│   ├── templates/
│   │   └── prompts.py       # LLMプロンプト定義・コード生成テンプレート・グリッドサイズ設定
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

## よく発生するエラー（Phase 3）

`auto` コマンドで生成された `_gen.py` はテンプレートベースのため、旧来のLLM生成コードで発生していたエラー（`NameError: name 'thin'` 等）は発生しません。

万一エラーが発生した場合、`prompts/<pdf_name>_prompt_error_fix.txt` が出力されます。その内容をAIチャットに投入してコードを修正し、`_gen.py` を上書き保存してから再度 `generate` を実行してください。

---

## 開発メモ

- `src/` と `data/` はコンテナにボリュームマウントされており、ホスト側の変更が即時反映されます。
- `auto` コマンドは `extract` が完了している（`_extracted.json` が存在する）ことを前提とします。
- Phase 3 の罫線後処理は `_grid_params.json` が存在する場合のみ実行されます（`extract` を実行していれば自動生成されます）。
- 旧来フローの `fill` コマンドは引き続き利用可能です（LLMとの手動ステップを経由したい場合）。

### アーキテクチャ設計の考え方

#### スクリプトが担う処理（`auto` コマンド）
- **座標マッピング**：`_row` / `_col` は Phase 1 で確定計算済みのためLLM不要
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

**将来対応案：** `extract` コマンドに `--pages 1-3` のようなオプションを追加して、処理対象ページを絞り込む。
