# Sheetling

PDFを解析し、LLMが生成したJSONコマンドを実行することで、**任意のPDFレイアウトを維持したままA4方眼Excelに変換**するツールです。  
多段パイプライン方式（解析 → 構造化 → 座標計算 → 描画）により、高精度な方眼紙レイアウト（1列約2mm、1行約5mm）への変換を実現します。

---

## 仕組み（3フェーズ構成）

```
PDF → [Phase 1] 解析・プロンプト生成 → [Phase 2] LLMでJSONコマンド生成（手動） → [Phase 3] Excel描画
```

| フェーズ | 実行者 | 内容 |
|----------|--------|------|
| Phase 1 | スクリプト | `pdfplumber`でPDFを解析し、LLM用の3段階プロンプトと抽出データを生成 |
| Phase 2 | 人間 + LLM | プロンプトを順にLLMに投入し、出力されたJSONコマンドを保存 |
| Phase 3 | スクリプト | 生成されたJSONコマンドを`openpyxl`で実行し、方眼Excelを出力 |

---

## セットアップ

Dockerを使用して環境を構築します。

```bash
docker compose up -d --build
```

---

## 実行手順

### Phase 1：PDF解析 & プロンプト生成

`data/in/` に対象のPDFを置いて、以下のコマンドを実行します：

```bash
docker compose exec app python -m src.main extract
```
※特定のPDFのみ実行する場合は `--pdf data/in/sample.pdf` を付与します。

`data/out/<pdf_name>/` に以下のファイルが生成されます：

| ファイル | 内容 |
|----------|------|
| `*_prompt_step1.txt` | LLMに投入するChunking用プロンプト |
| `*_prompt_step2.txt` | LLMに投入するGrid Mapping用プロンプト |
| `*_prompt_step3.txt` | LLMに投入するCommand Generation用プロンプト |
| `*_extracted.json` | PDFから直接抽出した生データ（座標や表情報） |
| `*_commands.json` | （空ファイル）Phase 2で生成したJSONを保存するためのファイル |

### Phase 2：LLMによる描画コマンド生成（手動）

1. まず `*_prompt_step1.txt` の内容をLLM（Gemini 1.5 Pro等）に入力します。
2. LLMが出力したJSONをコピーし、`*_prompt_step2.txt` 内の「`[ここにSTEP1の出力（JSON部分のみ）を貼り付けてください]`」の部分を書き換えて再度LLMに入力します。
3. 同様に、STEP2の出力結果（JSON）を使って `*_prompt_step3.txt` を書き換え、LLMに入力します。
4. 最終ステップ（STEP 3）で出力されたJSON配列をコピーし、`data/out/<pdf_name>/<pdf_name>_commands.json` に貼り付けて保存します。

### Phase 3：Excel生成

以下のコマンドを実行してExcelを描画します：

```bash
docker compose exec app python -m src.main generate
```

`data/out/<pdf_name>/<pdf_name>.xlsx` が生成されます。  
生成されるExcelファイルは、A4方眼紙レイアウト（列幅2mm・行高5mm相当）上に適切にセル結合と値の入力が行われた状態になります。

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

## 座標変換の仕様（Grid Mapping）

ExcelのA4方眼紙レイアウト（1列約2mm、1行約5mm）に合わせ、PDFのポイント(pt)からExcelの1始まりインデックスへ変換します。

| 軸 | 変換式 |
|----|--------|
| 列（X） | `floor((x * 0.3527) / 2.0) + 1` |
| 行（Y） | `floor((y * 0.3527) / 5.0) + 1` |

※この変換式は `prompts.py` 内のプロンプトを通じてLLMに指示され、座標計算が行われます。また、ページ境界での改ページ（`Break`）は `excel_writer.py` で自動的に設定されます。

---

## 使用パッケージ

| パッケージ | 用途 | ライセンス |
|-----------|------|-----------|
| `pdfplumber` | PDFからのテキスト・表バウンディングボックス抽出 | MIT |
| `openpyxl` | Excelファイルの生成・セル結合・改ページ設定 | MIT |

---

## 開発メモ

- `src/` と `data/` はコンテナにボリュームマウントされており、ホスト側の変更が即時反映されます。
