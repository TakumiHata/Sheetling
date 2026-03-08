# Sheetling

PDFを解析し、LLMが生成したPythonコードを実行することで、**任意のPDFを方眼Excelに変換**するツールです。  
テキスト・座標・フォント・色・罫線情報をPDFから抽出し、Excelの方眼（120列×可変行）上に忠実に再現します。

---

## 仕組み（3フェーズ構成）

```
PDF → [Phase 1] 解析・プロンプト生成 → [Phase 2] LLMでコード生成（手動） → [Phase 3] Excel生成
```

| フェーズ | 実行者 | 内容 |
|----------|--------|------|
| Phase 1 | スクリプト | pdfplumberでPDFを解析、LLM向けプロンプトを生成 |
| Phase 2 | 人間 + LLM | プロンプトをLLMに投入し、生成コードを保存 |
| Phase 3 | スクリプト | 生成コードを実行し、2シートExcelを出力 |

---

## セットアップ

```bash
docker compose up -d --build
```

---

## 実行手順

### Phase 1：PDF解析 & プロンプト生成

`data/in/` にPDFを置いて実行：

```bash
docker compose exec app python -m src.extract
```

`data/out/<pdf_name>/` に以下が生成されます：

| ファイル | 内容 |
|----------|------|
| `*_prompt.txt` | LLMに渡すプロンプト（座標変換式込み） |
| `*.json` | 抽出データ（テキスト・座標・フォント・色・罫線） |
| `*.md` | 人間可読な抽出結果 |
| `*_meta.json` | フォント・カラー・ページ数情報（Phase 3で使用） |

### Phase 2：LLMによるコード生成（手動）

1. `*_prompt.txt` の内容を LLM（ChatGPT、Gemini等）に送信
2. 返されたPythonコードを `data/out/<pdf_name>/<pdf_name>_gen.py` として保存

> **ポイント**: LLMは `page_width` を使ったx座標スケーリングと、累積 `top` 座標をそのまま行番号変換する形式でコードを出力します。

### Phase 3：Excel生成

```bash
docker compose exec app python -m src.generate
```

`data/out/<pdf_name>/<pdf_name>.xlsx`（2シート構成）が生成されます：

| シート | 内容 |
|--------|------|
| 変換結果 | PDFレイアウトを再現した方眼Excel |
| フォント・カラー情報 | 使用フォント一覧 + カラーコード一覧（プレビュー付き） |

---

## プロジェクト構成

```
Sheetling/
├── src/
│   ├── extract.py          # Phase 1 エントリポイント
│   ├── generate.py         # Phase 3 エントリポイント
│   └── core/
│       ├── extractor.py    # PDF解析（テキスト・座標・フォント・色・罫線）
│       ├── prompts.py      # LLM向けプロンプト生成
│       ├── executor.py     # 生成コード実行 + Excel出力
│       ├── pipeline.py     # フェーズ統合パイプライン
│       └── config.py       # 方眼サイズ等の設定（デフォルト: 4.65pt/マス、120列）
├── data/
│   ├── in/                 # 処理対象PDFを置く場所
│   └── out/                # 生成ファイルの出力先
├── Dockerfile
└── docker-compose.yml
```

---

## 座標変換の仕様

| 軸 | 変換式 |
|----|--------|
| 列（X） | `floor(x0 / page_width × 120) + 1` — PDFの実際の幅で正規化 |
| 行（Y） | `floor(top / 4.65) + 1` — 累積ページオフセット込みの絶対座標を使用 |

- **`top`** は `extractor.py` が前ページまでの実際の高さ（`page.height`）を累積加算したもの
- **ページ区切り**は `executor.py` が `page_heights` を元に自動計算（固定行数ではなく実測値ベース）
- そのため、A4以外のPDF（ACCESSレポート等）でも正しいページ分割が行われます

---

## 使用パッケージ

| パッケージ | 用途 | ライセンス |
|-----------|------|-----------|
| `pdfplumber` | テキスト・座標・罫線を高精度に抽出 | MIT |
| `openpyxl` | Excelファイルの生成・スタイル設定 | MIT |

---

## 開発メモ

- `src/` と `data/` はコンテナにボリュームマウントされており、ホスト側の変更が即時反映されます
- 方眼の設定値は `src/core/config.py` で変更可能（`unit_pt`, `target_cols`, `target_rows`）
