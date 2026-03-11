# Sheetling

PDFを解析し、LLMが生成したJSONコマンドを実行することで、**任意のPDFレイアウトを維持したままA4方眼Excelに変換**するツールです。  
多段パイプライン方式（解析 → 構造化 → 座標計算 → 描画）により、高精度な方眼紙レイアウト（約4mm角〜約8mm角などの正方形グリッド）への変換を実現します。

---

## 仕組み（3フェーズ・6ステップ構成）

```
PDF → [Phase 1] 解析・プロンプト生成 → [Phase 2] LLMによる6段階の推論とコード生成（手動） → [Phase 3] Excel描画
```

| フェーズ | 実行者 | 内容 |
|----------|--------|------|
| Phase 1 | スクリプト | `pdfplumber`でPDFを解析し、LLM用の6段階のプロンプトと座標データを生成。`data/out/<pdf_name>/prompts/` にファイルを出力 |
| Phase 2 | 人間 + LLM | STEP 1〜6のプロンプトを順にLLMに投入し、最終的にExcel生成用Pythonスクリプトを出力 |
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
    ├── <pdf_name>_prompt_step1.txt
    ├── <pdf_name>_prompt_step2.txt
    ├── <pdf_name>_prompt_step3.txt
    ├── <pdf_name>_prompt_step4.txt
    ├── <pdf_name>_prompt_step5.txt
    └── <pdf_name>_prompt_step6.txt
```

### Phase 2：LLMによる描画スクリプト生成（手動）

1. `*_prompt_step1.txt` の内容をLLM（Gemini 1.5 Pro等）に入力します。
2. LLMが出力したJSON配列をコピーし、**次のステップのプロンプトファイル**（`*_prompt_step2.txt`）内の「`[ここにSTEP Xの出力...を貼り付けてください]`」の部分を書き換えて、再度LLMに入力します。
3. この手順を STEP 1 から STEP 6 まで順番に繰り返します。
4. 最終ステップ（STEP 6）で出力された**Pythonコード**全体をコピーし、`data/out/<pdf_name>/<pdf_name>_gen.py` に貼り付けて保存します。

### Phase 3：Excel生成

以下のコマンドを実行してExcelを描画します：

```bash
# Docker環境の場合
docker compose exec app python -m src.main generate

# ローカルのPython環境の場合
python -m src.main generate
```

`data/out/<pdf_name>/<pdf_name>.xlsx` が生成されます。  

生成されるExcelファイルは、指定された**A4方眼紙レイアウト（完全に縦横比が1:1のピクセル設定）**上に適切にセル結合と値の入力が行われます。縮尺が変更されたりスケールダウンが発生したりすることなく、100%等倍でA4用紙に印刷できます。

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

## パイプラインの各ステップ（STEP 1〜6）詳解

Phase 2でLLMに実行させる6つのプロンプトステップには、それぞれ明確な役割（責務）が定義されています。

1. **Step 1（抽出 - Chunking）**
   `pdfplumber` によってPDFから抽出された物理的なテキスト要素や罫線要素（座標：Bounding Box）を読み込み、意味のある「チャンク」に分割します。
2. **Step 2（構造化アラインメント - Structural Alignment）**
   分断されたテキストの結合や、表（テーブル）データの行・列としての論理的な整理を行い、データの意味的な構造を再構築します。
3. **Step 3（グリッドマッピング - Grid Mapping）**
   選択した方眼サイズに従い、A4縦の有効印字領域内に各要素が収まるよう、元の物理ポイント(pt)座標から「Excelの行列インデックス」へのマッピング（割り当て計算）を行います。
4. **Step 4（コマンド生成 - Command Generation）**
   計算された座標データを基にして、Excel上で実行すべき具体的なアクション（例：`merge_and_set`（結合して値入力）、罫線の有無など）のコマンド形式に変換します。
5. **Step 5（ページフィット検証 - Page Fit Validation）**
   生成された全座標が、指定された方眼サイズの物理上限（最大列数・最大行数）を超えていないかの検証を行います。超過している場合は比率を保ったまま縮小再計算し、用紙に完璧に収まるように最終調整し、印刷範囲（`print_range`）を確定させます。
6. **Step 6（コード生成 - Code Generation）**
   最終確定したレイアウトコマンドをもとに、Python（`openpyxl`ベース）の実行可能な描画スクリプト（`_gen.py`）を生成します。（方眼のピクセルサイズ厳格指定と、スクリプト側での等倍スケーリングの保証を含定ます）

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
│   │   ├── pipeline.py      # フェーズ全体のフロー制御
│   │   └── config.py        # 内部グリッド・Excel出力設定
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
| `markitdown` | PDF全体の論理的なテキスト構造をMarkdown形式で抽出 (Microsoft製) |
| `pdfplumber` | PDF内のテキスト、表、罫線の詳細な座標情報を抽出 |
| `openpyxl` | Excelファイルの生成、セルの結合・スタイリング、印刷範囲設定 |
---

## 開発メモ

- `src/` と `data/` はコンテナにボリュームマウントされており、ホスト側の変更が即時反映されます。
