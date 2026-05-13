# Sheetling

PDFを解析し、**任意のPDFレイアウトを維持したままA4/A3方眼Excelに変換**するツールです。
テーブルの列ズレのない高精度な方眼紙レイアウトへの変換を、LLMを最小限に抑えた完全自動パイプラインで実現します。

---

## コマンド一覧

| コマンド | 内容 |
|---------|------|
| `auto` | PDF → Excel 自動変換。`1pt`・`2pt` の2サイズを同時出力。視覚的検証素材も生成 |
| `correct` | ビジョンLLMの修正指示（`_visual_corrections.json`）を適用してExcelを再生成 |
| `check` | `data/in/` 内の全PDFをスキャン画像/テキストPDFに分類し、結果CSVを出力 |

---

## 実行コマンド

### セットアップ（初回のみ）

```bash
pip install -r requirements.txt
```

### auto：PDF → Excel 自動変換

`data/in/` に対象のPDFを置いて実行します。

```bash
# 全PDF一括処理
python -m src.main auto

# 特定PDFのみ
python -m src.main auto --pdf data/in/test1/sample.pdf
```

### check：PDF種別判定

```bash
python -m src.main check
```

`data/doc/pdflist_check.csv` に結果を出力します。

| 判定 | 説明 |
|------|------|
| 通常PDF（テキストあり） | 少なくとも1ページにテキストが含まれる |
| スキャンPDF（画像） | 全ページにテキストがない（画像のみ） |
| エラー | ファイルが破損している等で読み取り不可 |

> [!NOTE]
> `auto` コマンドで変換できるのは「通常PDF（テキストあり）」のみです。スキャンPDFは事前にOCR処理が必要です。

### correct：視覚的検証の修正を適用

```bash
# 特定PDFを指定（推奨）
python -m src.main correct --pdf data/in/test1/sample.pdf

# 全対象を一括処理
python -m src.main correct
```

---

## 出力ファイル

`auto` 実行後、`data/out/<relative_path>/` に以下が生成されます。

> 出力先は `data/in/` からの相対パスで決まります。例: `data/in/test1/sample.pdf` → `data/out/test1/`

```text
data/out/<relative_path>/
├── <pdf_name>_extracted.json              # PDFから抽出した生データ（グリッド座標付き）
├── <pdf_name>_1pt_layout.json             # レイアウトJSON（1pt サイズ）
├── <pdf_name>_2pt_layout.json             # レイアウトJSON（2pt サイズ）
├── <pdf_name>_1pt_grid_params.json        # グリッドパラメータ（1pt サイズ）
├── <pdf_name>_2pt_grid_params.json        # グリッドパラメータ（2pt サイズ）
├── <pdf_name>_Python版_1pt.xlsx           # 生成されたExcelファイル（1pt）
├── <pdf_name>_Python版_2pt.xlsx           # 生成されたExcelファイル（2pt）
├── <pdf_name>.pdf                         # 元PDFのコピー（比較用）
└── prompts/
    ├── 1pt/
    │   └── page_1/
    │       ├── <pdf_name>_page1.png                     # PDFページ画像
    │       ├── <pdf_name>_excel_page1.png               # 罫線プレビュー画像
    │       ├── <pdf_name>_visual_review_page1.txt       # 視覚的検証プロンプト
    │       └── <pdf_name>_visual_corrections_page1.json # LLM修正指示（ユーザーが編集）
    └── 2pt/
        └── page_1/
            └── ...
```

| ファイル | 説明 |
|---------|------|
| `_extracted.json` | pdfplumberで抽出したテキスト・罫線・矩形の生データ。グリッド座標付き |
| `_layout.json` | Excelに描画するレイアウト情報（テキスト位置・罫線・フォント等） |
| `_grid_params.json` | コンテンツ境界・セルサイズ・用紙サイズなどグリッド計算パラメータ |
| `_Python版_1pt.xlsx` | 1ptグリッド（高密度・列幅3.48mm）のExcelファイル |
| `_Python版_2pt.xlsx` | 2ptグリッド（中密度・列幅6.18mm）のExcelファイル |
| `*.pdf` | 元PDFのコピー。生成ExcelとのPDF比較用 |
| `_page{N}.png` | PDFページ画像。視覚的検証に使用 |
| `_excel_page{N}.png` | 罫線プレビュー画像。コンテンツ範囲外はグレー表示 |
| `_visual_review_page{N}.txt` | ビジョンLLMへ投入する検証プロンプト |
| `_visual_corrections_page{N}.json` | LLMの修正指示を記述するJSON（`correct` コマンドが読み込む） |

---

## グリッドサイズ仕様

`auto` コマンドは `1pt`・`2pt` の2サイズを常に同時出力します。

### A4

| グリッド | 列幅 (mm) | 行高 (mm) | 縦 (cols × rows) | 横 (cols × rows) | フォント |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **1pt** | 3.48 | 6.44 | 47 × 39 | 70 × 25 | 7pt MS Gothic |
| **2pt** | 6.18 | 6.44 | 29 × 39 | 44 × 25 | 6pt MS Gothic |

### A3

| グリッド | 列幅 (mm) | 行高 (mm) | 縦 (cols × rows) | 横 (cols × rows) | フォント |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **1pt** | 3.48 | 6.44 | 70 × 57 | 104 × 39 | 7pt MS Gothic |
| **2pt** | 6.18 | 6.44 | 44 × 57 | 65 × 39 | 6pt MS Gothic |

用紙サイズ（A4/A3）と向き（縦/横）はPDFの寸法から自動検出します。

---

## 視覚的検証（任意）

`auto` 実行後、ビジョンLLMで罫線の再現精度をさらに高めることができます。

1. 社内AIチャット（画像入力対応）に以下を投入する：
   - `prompts/{grid_size}/page_{N}/<pdf_name>_page{N}.png`
   - `prompts/{grid_size}/page_{N}/<pdf_name>_excel_page{N}.png`
   - `prompts/{grid_size}/page_{N}/<pdf_name>_visual_review_page{N}.txt` の内容
2. LLMの出力した修正指示JSONを `_visual_corrections_page{N}.json` に保存する
3. `correct` コマンドで修正を適用する

**`_visual_corrections.json` の形式：**

```json
{
  "corrections": [
    {"action": "add_border",    "page": 1, "row": 3, "end_row": 8, "col": 2, "end_col": 15,
                                "borders": {"top": true, "bottom": true, "left": true, "right": true}},
    {"action": "remove_border", "page": 1, "row": 3, "end_row": 5, "col": 2, "end_col": 8}
  ]
}
```

---

## テスト

```bash
python -m pytest tests/ -v
```

110テストケースで以下をカバーしています：

| テストファイル | 対象モジュール |
|-------------|-------------|
| `test_font.py` | フォント正規化、罫線スタイルマッピング |
| `test_text.py` | 日本語判定、テキスト結合、水平ギャップ分割 |
| `test_grid.py` | コンテンツ境界検出、座標付与、線統合、用紙サイズ検出 |
| `test_layout.py` | テキスト要素生成、視覚行分割、重複排除、レイアウトJSON生成 |
| `test_excel.py` | Excel描画（テキスト・罫線・複数ページ・フォント色） |
| `test_preview.py` | プレビュー画像生成、セルメトリクス計算 |
| `test_pdf_extractor.py` | 矩形包含除去、色変換、座標丸め、エッジ統合 |
| `test_pipeline.py` | corrections適用（add/remove/fix）、エラー処理 |

---

## プロジェクト構成

```text
Sheetling/
├── src/
│   ├── main.py                # CLI エントリポイント（auto / correct / check）
│   ├── core/
│   │   ├── pipeline.py        # パイプラインオーケストレーション
│   │   ├── grid.py            # グリッド座標計算・用紙サイズ検出
│   │   └── layout.py          # レイアウトJSON生成
│   ├── parser/
│   │   └── pdf_extractor.py   # PDFデータ抽出（pdfplumber）
│   ├── renderer/
│   │   ├── excel.py           # Excel描画（openpyxl）
│   │   └── preview.py         # 罫線プレビュー画像生成（Pillow）
│   ├── templates/
│   │   └── prompts.py         # 視覚的検証プロンプト・グリッドサイズ設定
│   └── utils/
│       ├── logger.py          # ログ出力管理
│       ├── font.py            # フォント名正規化・罫線スタイルマッピング
│       └── text.py            # テキスト結合・日本語判定・水平ギャップ分割
├── tests/                     # pytest テストスイート（110テスト）
├── data/
│   ├── in/                    # 入力PDFディレクトリ
│   ├── out/                   # 出力ディレクトリ（解析結果・Excel）
│   └── doc/                   # check コマンドの出力ディレクトリ
├── Dockerfile
├── docker-compose.yml
└── requirements.txt
```

---

## 開発者向けドキュメント

コードを読む・改修する際は `docs/` を参照してください。

| ドキュメント | 内容 |
|------------|------|
| [アーキテクチャ](docs/architecture.md) | パイプライン全体のデータフロー・主要関数リファレンス・データ構造 |
| [グリッドシステム](docs/grid-system.md) | コンテンツ境界ベースのグリッド計算・座標変換ロジック・GRID_SIZES 定数 |
| [テーブル検出とテキスト配置](docs/table-detection.md) | pdfplumber パラメータ・word 優先フォールバック戦略 |
| [correct ワークフロー](docs/correct-workflow.md) | corrections JSON 仕様・検証プロンプト設計・安全機構の詳細 |
| [チューニングガイド](docs/tuning-guide.md) | GRID_SIZES 調整・PDF種別ごとの注意点・トラブルシューティング |

---

## 使用パッケージ

| パッケージ | バージョン | 用途 |
|-----------|-----------|------|
| `pdfplumber` | 0.11.9 | PDF内のテキスト・表・罫線の座標情報を抽出 |
| `openpyxl` | 3.1.5 | Excelファイルの生成・セルスタイリング・印刷範囲設定 |
| `Pillow` | *(Dockerイメージ提供)* | 罫線プレビュー画像の生成 |
| `pytest` | 9.0.3 | テストフレームワーク |
