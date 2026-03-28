# Sheetling

PDFを解析し、**任意のPDFレイアウトを維持したままA4/A3方眼Excelに変換**するツールです。
テーブルの列ズレのない高精度な方眼紙レイアウトへの変換を、LLMを最小限に抑えた完全自動パイプラインで実現します。
また、PDFがスキャン画像か通常テキストかを判定する `check` コマンドも備えています。

---

## 仕組み

```
PDF → [auto] 解析 → グリッド座標計算 → レイアウトJSON生成 → Excel直接描画
                                                              ↓
                                        （任意）ビジョンLLMによる視覚的検証
                                                              ↓
                                         [correct] 修正適用 → Excel再生成
```

| コマンド | 実行者 | 内容 |
|---------|--------|------|
| `auto` | スクリプト | PDF解析 → グリッド座標計算 → レイアウトJSON自動生成 → Excel直接描画。`1pt`・`2pt` の2サイズを同時出力。視覚的検証素材も自動生成 |
| `correct` | 人間 + ビジョンLLM | PDFページ画像 + 検証プロンプトをAIチャットに投入。修正指示JSONを適用し Excel を再生成 |
| `check` | スクリプト | `data/in/` 内の全PDFをスキャンし、テキスト有無で「通常PDF」「スキャンPDF（画像）」「エラー」に分類。結果CSVを `data/doc/` に出力 |

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
docker compose exec app python -m src.main auto

# 特定PDFのみ
docker compose exec app python -m src.main auto --pdf data/in/test1/001328648.pdf
```

| 処理 | 内容 |
|------|------|
| PDF解析 | `pdfplumber` でテキスト・罫線・矩形を抽出し、Excel グリッド座標を計算 |
| レイアウトJSON生成 | `table_border_rects` / `rects` / `words` から直接生成（LLM不要） |
| 座標検証・補正 | 整合性チェック・重複除去・クランプをスクリプトで確実に適用 |
| テキスト補完 | 抽出データと照合して欠落テキストを補完 |
| Excel出力 | `1pt`・`2pt` の2サイズをそれぞれ直接描画して出力 |
| 視覚的検証素材生成 | PDFページ画像・罫線プレビュー画像・検証プロンプトをページごとに出力 |

実行後、`data/out/<pdf_name>/` に以下が生成されます：

```text
data/out/<relative_path>/
├── <pdf_name>_extracted.json              # PDFから抽出した生データ（グリッド座標付き）
├── <pdf_name>_1pt_layout.json             # レイアウトJSON（1pt サイズ）
├── <pdf_name>_2pt_layout.json             # レイアウトJSON（2pt サイズ）
├── <pdf_name>_1pt_grid_params.json        # グリッドパラメータ（1pt サイズ）
├── <pdf_name>_2pt_grid_params.json        # グリッドパラメータ（2pt サイズ）
├── <pdf_name>_Python版_1pt.xlsx           # 生成された Excel ファイル（1pt）
├── <pdf_name>_Python版_2pt.xlsx           # 生成された Excel ファイル（2pt）
└── prompts/
    ├── 1pt/
    │   └── page_1/
    │       ├── <pdf_name>_page1.png                         # PDFページ画像
    │       ├── <pdf_name>_excel_page1.png                   # 罫線プレビュー画像（コンテンツ範囲外はグレー表示）
    │       ├── <pdf_name>_visual_review_page1.txt           # 視覚的検証プロンプト
    │       └── <pdf_name>_visual_corrections_page1.json     # LLM修正指示（ユーザーが編集）
    └── 2pt/
        └── page_1/
            └── ...
```

> [!NOTE]
> 出力先ディレクトリは `data/in/` からの相対パスで決まります。例: `data/in/test1/sample.pdf` → `data/out/test1/`

---

### 視覚的検証（任意）：ビジョンLLMで再現度を高める

`auto` 実行後、再現度をさらに高めたい場合にオプションで実施します。
PDFページ画像・罫線プレビュー画像はどちらも `auto` 実行時に自動生成されます。

**手順（サイズ・ページごとに繰り返す）：**

1. 社内AIチャット（画像入力対応）に以下を投入する：
   - `prompts/{grid_size}/page_{N}/<pdf_name>_page{N}.png`（PDFページ画像）
   - `prompts/{grid_size}/page_{N}/<pdf_name>_excel_page{N}.png`（罫線プレビュー画像）
   - `prompts/{grid_size}/page_{N}/<pdf_name>_visual_review_page{N}.txt` の内容（プロンプト）
2. LLMが出力した修正指示JSONを以下のファイルに上書き保存する：
   ```
   data/out/<relative_path>/prompts/<grid_size>/page_{N}/<pdf_name>_visual_corrections_page{N}.json
   ```
   ※ このファイルは `auto` 実行時に空テンプレートとして自動生成済み
3. 全ページ分完了したら `correct` コマンドで修正を適用する（次節）

**LLMが検出できる問題：**
- 欠落している罫線・枠
- 不要な罫線

> [!NOTE]
> スクリプトが計算した精密な座標にLLMの視覚的な判断を組み合わせることで、罫線の再現精度を向上できます。
> プレビュー画像ではコンテンツ範囲外がグレー表示され、プロンプトにも座標の範囲制約が含まれるため、AIの座標誤認を防止しています。

---

### check：PDF種別判定（スキャン画像 or テキスト）

`data/in/` 内の全PDFを再帰的にスキャンし、テキストが抽出できるかどうかで分類します。

```bash
docker compose exec app python -m src.main check
```

各PDFを以下の3種類に判定し、結果を `data/doc/pdflist_check.csv`（UTF-8 BOM付き）に出力します：

| 判定 | 説明 |
|------|------|
| 通常PDF（テキストあり） | 少なくとも1ページにテキストが含まれる |
| スキャンPDF（画像） | 全ページにテキストがない（画像のみ） |
| エラー | ファイルが破損している等で読み取り不可 |

**出力CSVの例：**

```csv
ファイルパス,ページ数,判定
test0/tirechange.pdf,1,通常PDF（テキストあり）
test1/scanned.pdf,5,スキャンPDF（画像）
```

> [!NOTE]
> `auto` コマンドで Excel に変換できるのは「通常PDF（テキストあり）」のみです。スキャンPDFは事前にOCR処理が必要です。

---

### correct：視覚的検証の修正を適用

```bash
# 特定PDFを指定して修正適用（推奨）
docker compose exec app python -m src.main correct --pdf data/in/test1/001328648.pdf

# data/out/ 以下の全 *_visual_corrections*.json を一括処理
docker compose exec app python -m src.main correct
```

`_visual_corrections_page{N}.json` の修正指示を `_layout.json` に適用し、Excel を再生成します。

**安全機構：**
- `add_border` の座標はコンテンツ範囲内に自動クランプされます（範囲外のはみ出しを防止）
- `remove_border` は指定範囲に完全に包含されるボーダーのみ削除します（外枠の巻き添え削除を防止）

**`_visual_corrections.json` の形式：**

```json
{
  "corrections": [
    {"action": "add_border",   "page": 1, "row": 3, "end_row": 8, "col": 2, "end_col": 15,
                               "borders": {"top": true, "bottom": true, "left": true, "right": true}},
    {"action": "remove_border","page": 1, "row": 3, "end_row": 5, "col": 2, "end_col": 8}
  ]
}
```

---

## CLIリファレンス

```
docker compose exec app python -m src.main <command> [options]
```

| command | 説明 |
|---------|------|
| `auto` | PDF → Excel 自動生成（解析 → レイアウトJSON生成 → Excel直接描画）。`1pt`・`2pt` の2サイズを同時出力 |
| `correct` | ビジョンLLMの修正指示（`_visual_corrections.json`）を適用して Excel を再生成 |
| `check` | `data/in/` 内の全PDFをスキャン画像 or テキストPDFに分類し、結果CSVを `data/doc/` に出力 |

| オプション | 対象command | 説明 |
|-----------|------------|------|
| `--pdf <path>` | `auto`, `correct` | 処理対象PDFのパスまたはPDF名（省略時は全対象を処理） |

---

## 対応用紙サイズ

PDF の寸法から用紙サイズ（A4/A3）と向き（縦/横）を自動検出します。

| 用紙 | 寸法 (pt) | 判定条件 |
|------|----------|---------|
| A4 | 595 × 842 | 長辺 ≤ 1000pt |
| A3 | 842 × 1190 | 長辺 > 1000pt |

---

## グリッドサイズ仕様

`auto` コマンドは `1pt`・`2pt` の2サイズを常に同時出力します。用途に応じて出力ファイルを選択してください。

### A4

| グリッド | 列幅 (mm) | 行高 (mm) | 縦 (cols × rows) | 横 (cols × rows) | フォント |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **1pt** | 3.48 | 6.44 | 62 × 42 | 96 × 30 | 7pt MS Gothic |
| **2pt** | 6.18 | 6.44 | 37 × 42 | 58 × 30 | 6pt MS Gothic |

### A3

| グリッド | 列幅 (mm) | 行高 (mm) | 縦 (cols × rows) | 横 (cols × rows) | フォント |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **1pt** | 3.48 | 6.44 | 92 × 61 | 128 × 44 | 7pt MS Gothic |
| **2pt** | 6.18 | 6.44 | 57 × 61 | 79 × 44 | 6pt MS Gothic |

> [!NOTE]
> セルサイズは A4/A3 で同一です。用紙が大きい分だけ列数・行数が増えます。
> `1pt` は高密度（細かいグリッド）、`2pt` は中密度（やや広いグリッド）です。Excel の列幅はデスクトップ Excel (MDW=8) での表示値に合わせて調整されています。

---

## プロジェクト構成

```text
Sheetling/
├── src/
│   ├── main.py              # CLI エントリポイント（auto / correct / check）
│   ├── core/
│   │   └── pipeline.py      # パイプライン全体の制御・座標計算・自動生成・Excel直接描画
│   ├── parser/
│   │   └── pdf_extractor.py # PDFデータ抽出 (pdfplumber)
│   ├── templates/
│   │   └── prompts.py       # 視覚的検証プロンプト・グリッドサイズ設定（A4/A3）
│   └── utils/
│       └── logger.py        # ログ出力管理
├── data/
│   ├── in/                  # 入力PDFディレクトリ
│   ├── out/                 # 出力ディレクトリ（解析結果・Excel）
│   └── doc/                 # check コマンドの出力ディレクトリ（PDF判定CSV）
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
| `Pillow` | 罫線プレビュー画像の生成（コンテンツ範囲外のグレーアウト含む） |

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
- **Excel直接描画**：`_render_layout_to_xlsx` で openpyxl に直接書き込み

#### ビジョンLLMが担う処理（`correct` コマンド、任意）
- **視覚的な差分検出**：スクリプトが計算した精密な座標 × LLMの目視確認で再現度を向上
- **意味的な判断が必要なケース**：複雑なレイアウトでスクリプトが見落とした欠落要素の検出
