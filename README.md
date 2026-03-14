# Sheetling

PDFを解析し、LLMが生成したPythonコードを実行することで、**任意のPDFレイアウトを維持したままA4方眼Excelに変換**するツールです。
2ステップ・パイプライン方式（列アンカー確定 → コード生成）により、テーブルの列ズレのない高精度な方眼紙レイアウトへの変換を実現します。

---

## 仕組み（2フェーズ・2ステップ構成）

```
PDF → [Phase 1] 解析・プロンプト生成 → [Phase 2] LLMによる2段階の推論とコード生成（手動） → [Phase 3] Excel描画
```

| フェーズ | 実行者 | 内容 |
|----------|--------|------|
| Phase 1 | スクリプト | `pdfplumber`でPDFを解析し、LLM用の2段階プロンプトと抽出データを生成。`data/out/<pdf_name>/prompts/` にファイルを出力 |
| Phase 2 | 人間 + LLM | STEP 1・2のプロンプトを順にLLMに投入し、最終的にExcel生成用Pythonスクリプトを出力 |
| Phase 3 | スクリプト | 保存したスクリプト（`_gen.py`）を環境上で実行し、方眼Excelを出力 |

---

## セットアップ

実行環境に合わせて、Dockerを使用するか、ローカルのPython環境を使用するか選択できます。

### Docker環境を使用する場合
```bash
docker compose up -d --build
```

### ローカルのPython環境を使用する場合（Dockerなし）
Python環境および `pip` がインストールされていることを前提とします。
以下のコマンドで必要なパッケージをインストールしてください。
```bash
pip install -r requirements.txt
```

---

## 実行手順

### Phase 1：PDF解析 & プロンプト生成

`data/in/` に対象のPDFを置いて、以下のコマンドを実行します：

```bash
# Docker環境の場合
docker compose exec app python -m src.main extract [--grid-size <size>]

# ローカルのPython環境の場合
python -m src.main extract [--grid-size <size>]
```
- `--grid-size <size>`：生成する方眼のサイズを指定できます。
  - `small`（デフォルト）
  - `medium`
  - `large`
- `--pdf <path>`：特定のPDFのみ実行する場合に付与します（例: `--pdf data/in/sample.pdf`）。

`data/out/<pdf_name>/` に以下の構成でファイルが生成されます：

```text
data/out/<pdf_name>/
├── <pdf_name>_extracted.json    # PDFから抽出した生データ
├── <pdf_name>_gen.py           # (空) Phase 2で生成したコードをここに貼り付ける
└── prompts/                    # 各ステップ用プロンプトファイル
    ├── <pdf_name>_prompt_step1.txt   # 列アンカー確定プロンプト（抽出データ込み）
    └── <pdf_name>_prompt_step2.txt   # Pythonコード生成プロンプト
```

### Phase 2：LLMによる描画スクリプト生成（手動）

1. `*_prompt_step1.txt` の内容をLLM（Gemini 2.5 Flash等）に入力します。
2. LLMが出力したJSON配列をコピーし、`*_prompt_step2.txt` 内の `[ここにSTEP 1の出力...]` の部分を書き換えて、再度LLMに入力します。
3. STEP 2で出力された**Pythonコード**全体をコピーし、`data/out/<pdf_name>/<pdf_name>_gen.py` に貼り付けて保存します。

### Phase 3：Excel生成

以下のコマンドを実行してExcelを描画します：

```bash
# Docker環境の場合
docker compose exec app python -m src.main generate

# ローカルのPython環境の場合
python -m src.main generate
```

`data/out/<pdf_name>/<pdf_name>.xlsx` が生成されます。

生成されるExcelファイルは、指定された**A4方眼紙レイアウト（完全に縦横比が1:1のピクセル設定）**上に適切に値の入力が行われます。縮尺が変更されたりスケールダウンが発生したりすることなく、100%等倍でA4用紙に印刷できます。

#### 方眼サイズ（`--grid-size`）の詳細仕様

絶対等倍でのA4出力を保証するため、各サイズには数学的に計算された最大グリッド数が設定されています。

| グリッドサイズ | 方眼の大きさ (mm) | 最大列数 (A4縦) | 最大行数 (A4縦) | Excel設定値 (幅/高さ) |
| :--- | :--- | :--- | :--- | :--- |
| **`small`** | **約 4.0 mm** | **62 列** | **76 行** | 幅: 1.45 / 高さ: 11.34 |
| **`medium`** | **約 6.0 mm** | **36 列** | **50 行** | 幅: 2.53 / 高さ: 17.01 |
| **`large`** | **約 8.0 mm** | **26 列** | **38 行** | 幅: 3.61 / 高さ: 22.68 |

> [!NOTE]
> これらの数値は、A4用紙の物理的な印字可能領域（余白を除いた約180mm x 270mm）を基準に、各サイズの正方形グリッドを敷き詰めた際の理論限界値です。

---

## パイプラインの各ステップ（STEP 1〜2）詳解

Phase 2でLLMに実行させる2つのプロンプトステップには、それぞれ明確な役割（責務）が定義されています。

1. **Step 1（レイアウト仕様JSON生成）**
   pdfplumberが抽出した `words`（テキスト・フォント色・サイズ）と `rects`（矩形枠・背景色）に、事前計算済みのExcel行列番号（`_row`/`_col`等）が付与されています。LLMはこの座標をそのまま使用して `text` 要素と `border_rect` 要素のみからなるレイアウト仕様JSONを出力します。

2. **Step 2（コード生成 - Code Generation）**
   Step 1のレイアウト仕様JSONをもとに、Python（`openpyxl`ベース）の実行可能な描画スクリプト（`_gen.py`）を生成します。`text` はセルへの値書き込み・フォント設定、`border_rect` は外枠罫線と背景色の塗りつぶしとして描画されます。

---

## プロジェクト構成

```text
Sheetling/
├── src/
│   ├── main.py             # 実行エントリポイント
│   ├── parser/
│   │   └── pdf_extractor.py # Phase 1: PDFデータ抽出 (pdfplumber)
│   ├── templates/
│   │   └── prompts.py       # Phase 1: 各ステップのLLMプロンプト定義
│   ├── core/
│   │   └── pipeline.py      # フェーズ全体のフロー制御
│   └── utils/
│       └── logger.py        # ログ出力管理
├── data/
│   ├── in/                  # 入力PDFディレクトリ
│   └── out/                 # 出力ディレクトリ（解析結果・Excel）
├── Dockerfile               # 環境構築用
├── docker-compose.yml       # サービス実行定義
└── requirements.txt         # 依存ライブラリ
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
