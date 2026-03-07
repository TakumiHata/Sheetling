# Sheetling

PDFを読み込み、pdfplumber等を使用してテキスト・座標・フォント情報を抽出し、AIが生成したPythonソースコードを実行して高精度な3シートExcelファイルを生成するツールです。

## プロジェクト構成

- `src/`: アプリケーションのソースコード
  - `main.py`: パイプラインのエントリポイント
  - `core/`: 解析・生成の主要ロジック
    - `extractor.py`: pdfplumberによるPDF解析（テキスト・座標・フォント・色・罫線）
    - `image_converter.py`: pdf2imageによるPDF→画像変換
    - `prompts.py`: AI向けプロンプト生成（Pythonコード出力指示）
    - `executor.py`: AI出力Pythonコード実行 + 3シートExcel生成
    - `pipeline.py`: 各モジュールを統合するパイプライン
    - `schema.py`: PDF抽出結果のPydantic型定義
    - `config.py`: 方眼サイズ等の設定値
  - `utils/`: ロガーなどのユーティリティ
- `data/`: データの入出力ディレクトリ
  - `in/`: 処理対象の入力PDFファイル
  - `out/`: 生成ファイル群（JSONやプロンプト、Excelなど）
- `tests/`: テストファイル群

## セットアップ

### Docker コンテナの起動

```bash
docker compose up -d --build
```

## 実行フロー

### Phase 1: PDF解析 & プロンプト生成

`data/in/` にPDFファイルを配置し、以下を実行します。

```bash
docker compose exec app python -m src.main
```

実行後、`data/out/<pdf_name>/` に以下が生成されます:
- `*_prompt.txt` — LLMに渡すプロンプト
- `*.json` — 抽出データ（テキスト・座標・フォント・色）
- `*.md` — 人間可読な抽出結果
- `*_meta.json` — フォント・カラー・画像パス情報（Phase 3で使用）
- `*_page*.png` — PDF画像（3シート目用）

### Phase 2: LLMによるPythonコード生成（手動）

1. `*_prompt.txt` の内容をコピーし、LLM（ChatGPT, Geminiなど）に送信
2. LLMが返したPythonコードを `data/out/<pdf_name>/<pdf_name>_gen.py` として保存

### Phase 3: Pythonコード実行 & 3シートExcel生成

再度メインスクリプトを実行します。

```bash
docker compose exec app python -m src.main
```

3シート構成のExcelファイル (`<pdf_name>.xlsx`) が生成されます:

| シート | 内容 |
|--------|------|
| 変換結果 | AIがPDFレイアウトを再現したExcel |
| フォント・カラー情報 | 使用フォント一覧 + カラーコード一覧 |
| PDF画像 | PDFの各ページを画像として添付 |

## 開発ガイド

- ホスト側の `src/` および `data/` ディレクトリはコンテナ内にマウントされているため、ホスト側での編集が即座に反映されます。
- 方眼サイズの設定は `src/core/config.py` を参照してください（デフォルト: 4.96pt / 1.75mm）。
