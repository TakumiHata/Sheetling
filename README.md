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
| Phase 1 | スクリプト | `pdfplumber`でPDFを解析し、LLM用の6段階のプロンプトと抽出データを生成 |
| Phase 2 | 人間 + LLM | STEP 1〜6のプロンプトを順にLLMに投入し、最終的にExcel生成用Pythonスクリプトを出力 |
| Phase 3 | スクリプト | 生成されたPythonスクリプト（`_gen.py`）を実行し、方眼Excelを出力 |

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

`data/out/<pdf_name>/` に以下のファイルが生成されます：

| ファイル | 内容 |
|----------|------|
| `*_prompt_step1.txt` | [STEP 1] Chunking・構造抽出用プロンプト |
| `*_prompt_step2.txt` | [STEP 2] 構造化アラインメント用プロンプト |
| `*_prompt_step3.txt` | [STEP 3] Grid Mapping用プロンプト |
| `*_prompt_step4.txt` | [STEP 4] Command Generation用プロンプト |
| `*_prompt_step5.txt` | [STEP 5] Page Fit Validation用プロンプト |
| `*_prompt_step6.txt` | [STEP 6] Code Generation用プロンプト |
| `*_extracted.json` | PDFから直接抽出した生データ（座標や表情報） |
| `*_gen.py` | （空ファイル）Phase 2で生成したPythonコードを保存するためのファイル |

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

#### 方眼サイズ（`--grid-size`）の仕様

指定できるサイズとその詳細は以下の通りです。

- **`small` (約4.0mm角の小サイズ方眼)**
  - A4縦1ページへの最大配置可能数: 横 45列 × 縦 67行
- **`medium` (約6.0mm角の中サイズ方眼)**
  - A4縦1ページへの最大配置可能数: 横 30列 × 縦 45行
- **`large` (約8.0mm角の大サイズ方眼)**
  - A4縦1ページへの最大配置可能数: 横 22列 × 縦 33行

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

```
Sheetling/
├── src/
│   ├── main.py             # 実行エントリポイント
│   ├── parser/
│   │   └── pdf_extractor.py # Phase 1: PDF解析・抽出ロジック
│   ├── templates/
│   │   └── prompts.py       # Phase 1: LLMプロンプトの定義
│   ├── renderer/
│   │   └── excel_writer.py  # Phase 3: JSONからのExcel描画ロジック
│   ├── core/
│   │   ├── pipeline.py      # 全体フローを管理
│   │   └── config.py        # 共通設定
│   └── utils/
│       └── logger.py
├── data/
│   ├── in/                  # 入力PDFディレクトリ
│   └── out/                 # 出力ディレクトリ
├── Dockerfile
└── docker-compose.yml
```

---

## 使用パッケージ

| パッケージ | 用途 | ライセンス |
|-----------|------|-----------|
| `pdfplumber` | PDFからのテキスト・表バウンディングボックス抽出 | MIT |
| `openpyxl` | Excelファイルの生成・セル結合・改ページ設定 | MIT |
| `markitdown` | （必要に応じて用途を記載：ファイルのMarkdown変換など） | MIT |
---

## 開発メモ

- `src/` と `data/` はコンテナにボリュームマウントされており、ホスト側の変更が即時反映されます。
