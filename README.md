# Sheetling

PDFを解析し、LLMが生成したPythonコードを実行することで、**任意のPDFレイアウトを維持したままA4/A3方眼Excelに変換**するツールです。
4ステップ・パイプライン方式により、テーブルの列ズレのない高精度な方眼紙レイアウトへの変換を実現します。

---

## 仕組み（3フェーズ・4ステップ構成）

```
PDF → [Phase 1] 解析・プロンプト生成
         ↓
      [Phase 2] LLMによる推論とコード生成（手動・社内AIチャット等）
        STEP 1: レイアウトJSON生成
        STEP 1.5: JSON検証・補正
        ↓
      [fill] テキスト補完 & STEP 2プロンプト更新（スクリプト自動）
        ↓
        STEP 2: Pythonコード生成
         ↓
      [Phase 3] Excel描画 + 確定的ボーダー後処理
```

| フェーズ | コマンド | 実行者 | 内容 |
|----------|---------|--------|------|
| Phase 1 | `extract` | スクリプト | `pdfplumber` でPDFを解析。Excel グリッド座標を事前計算し、LLM用プロンプト・抽出データを `data/out/<pdf_name>/` に出力 |
| Phase 2 STEP 1〜1.5 | — | 人間 + LLM | STEP 1・1.5 のプロンプトを順にLLMに投入し、レイアウトJSONを生成・補正 |
| fill | `fill` | スクリプト | STEP 1.5 出力JSONに欠落テキストを自動補完し、STEP 2 プロンプトを更新 |
| Phase 2 STEP 2 | — | 人間 + LLM | STEP 2 プロンプト（補完済みJSON入り）をLLMに投入し、Excel生成Pythonコードを取得 |
| Phase 3 | `generate` | スクリプト | 生成コード（`_gen.py`）を実行してExcelを出力。罫線を確定的に後処理で上書き適用 |

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

### Phase 1：PDF解析 & プロンプト生成

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
├── <pdf_name>_gen.py              # (空) Phase 2で生成したコードをここに貼り付ける
└── prompts/
    ├── <pdf_name>_prompt_step1.txt    # STEP 1: レイアウトJSON生成プロンプト
    ├── <pdf_name>_prompt_step1_5.txt  # STEP 1.5: JSON検証・補正プロンプト
    └── <pdf_name>_prompt_step2.txt    # STEP 2: Pythonコード生成プロンプト
```

---

### Phase 2：LLMによる描画スクリプト生成（手動）

社内AIチャット等のLLMを使い、以下の順に実行します。

#### STEP 1 — レイアウトJSON生成

1. `prompts/<pdf_name>_prompt_step1.txt` の内容をLLMに貼り付ける
2. LLMが出力したJSON配列（`[` から `]` まで）をコピーする

#### STEP 1.5 — JSON検証・補正

1. `prompts/<pdf_name>_prompt_step1_5.txt` を開く
2. `[ここにSTEP 1の出力...]` をSTEP 1で得たJSONに置き換えてLLMに貼り付ける
3. LLMが出力した補正済みJSONをコピーし、**以下のパスに保存する**：

```
data/out/<pdf_name>/prompts/<pdf_name>_step1_5_input.json
```

---

### fill：テキスト補完 & STEP 2プロンプト更新

```bash
# 特定PDFのみ
python -m src.main fill --pdf sample

# data/out/ 以下の全 *_step1_5_input.json を一括処理
python -m src.main fill
```

**処理内容：**
- `_step1_5_input.json` を読み込み、`_extracted.json` と照合してLLMが見落としたテキストを自動補完
- 補完済みJSONを `prompts/<pdf_name>_step1_5_output.json` として保存
- `prompts/<pdf_name>_prompt_step2.txt` のプレースホルダーを補完済みJSONで自動置換

**オプション：**

| オプション | デフォルト | 説明 |
|-----------|----------|------|
| `--pdf` | なし（全対象） | PDF名（拡張子なし）または PDFファイルパス |

---

### Phase 2（続き）：STEP 2

#### STEP 2 — Pythonコード生成

1. `prompts/<pdf_name>_prompt_step2.txt`（fillにより補完済みJSONが埋め込まれている）をLLMに貼り付ける
2. 出力されたPythonコード全体を `data/out/<pdf_name>/<pdf_name>_gen.py` に貼り付けて保存する

---

### Phase 3：Excel生成

```bash
# data/out/ 以下の全 *_gen.py を一括処理
python -m src.main generate

# （--pdf オプションは generate では不要・全自動）
```

`data/out/<pdf_name>/<pdf_name>.xlsx` が生成されます。

**処理内容：**
- `_gen.py` を実行してExcelを生成
- `_extracted.json` と `_grid_params.json` を参照し、罫線を確定的に後処理で上書き適用（LLM生成コードのボーダー誤りを自動修正）
- エラー発生時は `prompts/<pdf_name>_prompt_error_fix.txt` にエラー修正用プロンプトを出力

---

## CLIリファレンス

```
python -m src.main <phase> [options]
```

| phase | 説明 |
|-------|------|
| `extract` | Phase 1: PDF解析 & プロンプト生成 |
| `fill` | Phase 1.5後処理: テキスト補完 & STEP 2プロンプト更新 |
| `generate` | Phase 3: 生成コードを実行してExcel出力 |

| オプション | 対象phase | 説明 |
|-----------|----------|------|
| `--pdf <path>` | `extract`, `fill` | 処理対象PDFのパスまたはPDF名（省略時は全対象を処理） |
| `--grid-size <size>` | `extract` | 方眼サイズ: `small`（デフォルト）/ `medium` / `large` |

---

## パイプラインのステップ詳解

### STEP 1（レイアウトJSON生成）
`pdfplumber` が抽出した `words`（テキスト・フォント色・サイズ）と `table_border_rects`（テーブルセル境界）、`rects`（矩形枠）に、事前計算済みのExcel行列番号（`_row`/`_col`等）が付与されています。LLMはこの座標をそのまま使用して `text` 要素と `border_rect` 要素のみからなるレイアウト仕様JSONを出力します。

### STEP 1.5（JSON検証・補正）
STEP 1の出力を検証・補正します。座標の整合性チェック、重複除去、欠落テキストの補完、座標のクランプを行います。

### fill（テキスト補完・自動）
STEP 1.5のLLM出力に対し、`_extracted.json` の `words` と照合して欠落テキストを確実に補完します。LLMへの依存なしに動作するため、LLMが見落としたテキストを毎回漏れなく補完できます。

### STEP 2（Pythonコード生成）
補正済みレイアウト仕様JSONをもとに、`openpyxl` ベースの実行可能な描画スクリプトを生成します。

### Phase 3（確定的ボーダー後処理）
`_gen.py` 実行後、`_extracted.json` から得た罫線データを `openpyxl` で直接適用し、LLM生成コードのボーダー誤りを上書き修正します。

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
│   ├── main.py              # CLI エントリポイント（extract / fill / generate）
│   ├── core/
│   │   └── pipeline.py      # フェーズ全体のフロー制御・グリッド座標計算・ボーダー後処理
│   ├── parser/
│   │   └── pdf_extractor.py # Phase 1: PDFデータ抽出 (pdfplumber)
│   ├── templates/
│   │   └── prompts.py       # LLMプロンプト定義・グリッドサイズ設定
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

## 開発メモ

- `src/` と `data/` はコンテナにボリュームマウントされており、ホスト側の変更が即時反映されます。
- `fill` コマンドは `_step1_5_input.json` が存在しない場合は警告を出して終了します。必ずSTEP 1.5の出力を保存してから実行してください。
- Phase 3 の罫線後処理は `_grid_params.json` が存在する場合のみ実行されます（Phase 1 を実行していれば自動生成されます）。
