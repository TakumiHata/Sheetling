# Sheetling

PDFを解析し、**任意のPDFレイアウトを維持したままA4/A3方眼Excelに変換**するツールです。
テーブルの列ズレのない高精度な方眼紙レイアウトへの変換を、完全自動パイプラインで実現します。

---

## コマンド一覧

| コマンド | 内容 |
|---------|------|
| `auto` | PDF → Excel 自動変換。`1pt`・`2pt` の2サイズを同時出力。目視確認用プレビュー画像も生成 |
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
└── <pdf_name>.pdf                         # 元PDFのコピー（比較用）
```

| ファイル | 説明 |
|---------|------|
| `_extracted.json` | pdfplumberで抽出したテキスト・罫線・矩形の生データ。グリッド座標付き |
| `_layout.json` | Excelに描画するレイアウト情報（テキスト位置・罫線・フォント等） |
| `_grid_params.json` | コンテンツ境界・セルサイズ・用紙サイズなどグリッド計算パラメータ |
| `_Python版_1pt.xlsx` | 1ptグリッド（高密度・列幅3.48mm）のExcelファイル |
| `_Python版_2pt.xlsx` | 2ptグリッド（中密度・列幅6.18mm）のExcelファイル |
| `*.pdf` | 元PDFのコピー。生成ExcelとのPDF比較用 |

---

## グリッドサイズ仕様

`auto` コマンドは `1pt`・`2pt` の2サイズを常に同時出力します。

### A4

| グリッド | 列幅 (mm) | 行高 (mm) | 縦 (cols × rows) | 横 (cols × rows) | フォント |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **1pt** | 3.48 | 6.44 | 58 × 44 | 85 × 30 | 7pt MS 明朝 |
| **2pt** | 6.18 | 6.44 | 35 × 44 | 52 × 30 | 6pt MS 明朝 |

### A3

| グリッド | 列幅 (mm) | 行高 (mm) | 縦 (cols × rows) | 横 (cols × rows) | フォント |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **1pt** | 3.48 | 6.44 | 88 × 64 | 125 × 45 | 7pt MS 明朝 |
| **2pt** | 6.18 | 6.44 | 51 × 64 | 72 × 45 | 6pt MS 明朝 |

用紙サイズ（A4/A3）と向き（縦/横）はPDFの寸法から自動検出します。

---

## テスト

```bash
python -m pytest tests/ -v
```

118テストケースで以下をカバーしています：

| テストファイル | 対象モジュール |
|-------------|-------------|
| `test_font.py` | フォント正規化、罫線スタイルマッピング |
| `test_text.py` | 日本語判定、テキスト結合、水平ギャップ分割 |
| `test_grid.py` | コンテンツ境界検出、座標付与、線統合、用紙サイズ検出 |
| `test_layout.py` | テキスト要素生成、視覚行分割、重複排除、レイアウトJSON生成 |
| `test_excel.py` | Excel描画（テキスト・罫線・複数ページ・フォント色） |
| `test_pdf_extractor.py` | 矩形包含除去、色変換、座標丸め、エッジ統合 |
| `test_edges.py` | エッジ分解・集約・スパンフィルタ |
| `test_pipeline.py` | パイプライン補助関数 |

---

## プロジェクト構成

```text
Sheetling/
├── src/
│   ├── main.py                # CLI エントリポイント（auto / check）
│   ├── core/
│   │   ├── pipeline.py        # パイプラインファサード（SheetlingPipeline クラス）
│   │   ├── auto_layout_service.py  # auto パイプライン実装
│   │   ├── edges.py           # エッジ単位罫線モデル（分解・集約・スパンフィルタ）
│   │   ├── grid.py            # グリッド座標計算・用紙サイズ検出
│   │   ├── grid_config.py     # GRID_SIZES 定数（A4/A3 × 1pt/2pt のセル寸法・Excel設定）
│   │   ├── constants.py       # 共有の数値定数（tolerance・閾値など）
│   │   └── layout.py          # レイアウトJSON生成
│   ├── parser/
│   │   └── pdf_extractor.py   # PDFデータ抽出（pdfplumber）
│   ├── renderer/
│   │   └── excel.py           # Excel描画（openpyxl）
│   └── utils/
│       ├── logger.py          # ログ出力管理
│       ├── font.py            # フォント名正規化・罫線スタイルマッピング
│       └── text.py            # テキスト結合・日本語判定・水平ギャップ分割
├── tests/                     # pytest テストスイート（118テスト）
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
| [チューニングガイド](docs/tuning-guide.md) | GRID_SIZES 調整・PDF種別ごとの注意点・トラブルシューティング |

---

## 使用パッケージ

| パッケージ | バージョン | 用途 |
|-----------|-----------|------|
| `pdfplumber` | 0.11.9 | PDF内のテキスト・表・罫線の座標情報を抽出 |
| `openpyxl` | 3.1.5 | Excelファイルの生成・セルスタイリング・印刷範囲設定 |
| `pytest` | 9.0.3 | テストフレームワーク |
