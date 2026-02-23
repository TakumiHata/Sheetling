# Sheetling

ACCESSからエクスポートされたPDFを読み込み、Doclingを使用してテキスト・座標情報を抽出後、LLMを用いて高精度なExcelファイルを生成するツールです。

## プロジェクト構成

- `src/`: アプリケーションのソースコード
  - `main.py`: パイプラインのエントリポイント
  - `core/`: 解析・生成の主要ロジック
- `data/`: データの入出力ディレクトリ
  - `01_input_pdf/`: 処理対象のPDFを配置
  - `02_inter_md/`: 抽出されたMarkdown（中間ファイル）
  - `03_inter_json/`: 抽出された座標データ（中間ファイル）
  - `04_output_excel/`: 生成されたExcelファイル

## セットアップ

### 1. 環境変数の設定

`.env.example` を `.env` にコピーし、必要なAPIキーを設定してください。

```bash
cp .env.example .env
```

### 2. Docker コンテナの起動

以下のコマンドを実行して、開発環境（コンテナ）を起動します。

```bash
docker compose up -d --build
```

## 実行方法

### パイプラインの実行

`data/01_input_pdf/` にPDFファイルを配置した後、以下のコマンドを実行します。

```bash
docker compose exec app env PYTHONPATH=. python3 src/main.py
```

## 開発ガイド

- ホスト側の `src/` および `data/` ディレクトリはコンテナ内にマウントされているため、ホスト側での編集が即座に反映されます。
- Docling の実行に必要なシステムライブラリ（libgl1等）は Dockerfile 内でインストールされています。
