# Sheetling アーキテクチャ設計

PDF→Excel変換を「PDF解析→AI(LLM)がPythonソース出力→実行で3シートExcel生成」の3フェーズで行うシステム。

## 全体アーキテクチャ図

```mermaid
sequenceDiagram
    participant User as ユーザー
    participant Phase1 as "1. PDF解析・プロンプト生成<br/>(Python)"
    participant LLM as "2. LLM<br/>(ブラウザから手動)"
    participant Phase3 as "3. Excel生成<br/>(Pythonコード実行)"

    User->>Phase1: PDFファイルを配置
    
    rect rgb(240, 248, 255)
    Note over Phase1: 【Phase 1】PDF解析 & プロンプト生成
    Phase1->>Phase1: pdfplumberでテキスト・座標・<br/>フォント・色・罫線を抽出
    Phase1->>Phase1: pdf2imageでPDF→画像変換
    Phase1->>Phase1: プロンプトを生成
    Phase1-->>User: プロンプトファイル (*_prompt.txt) を出力
    end
    
    rect rgb(255, 240, 245)
    Note over LLM: 【Phase 2】LLMがPythonコード生成（手動）
    User->>LLM: プロンプトをブラウザ上で入力
    LLM->>LLM: openpyxlを使ったPythonコードを生成
    LLM-->>User: Pythonソース (*_gen.py)
    end
    
    rect rgb(240, 255, 240)
    Note over Phase3: 【Phase 3】Pythonコード実行 + 3シートExcel
    User->>Phase3: Pythonソースを配置してパイプライン再実行
    Phase3->>Phase3: AI生成コード実行（1シート目）
    Phase3->>Phase3: フォント・カラー一覧生成（2シート目）
    Phase3->>Phase3: PDF画像添付（3シート目）
    Phase3-->>User: 3シートExcelファイル (.xlsx)
    end
```

## データフロー図

```mermaid
flowchart TD
    A[入力: PDFファイル] -->|pdfplumber| B[テキスト・座標・フォント・色・罫線]
    A -->|pdf2image| C[PDF画像 PNG]
    
    subgraph phase1["Phase 1: PDF解析 & プロンプト生成"]
        B --> D[抽出JSON + Markdown]
        D --> E["プロンプト (*_prompt.txt)"]
        B --> F["メタデータ (*_meta.json)<br/>フォント・色・画像パス"]
    end

    E -->|"ユーザーがLLMに手動入力"| G

    subgraph phase2["Phase 2: LLMがPythonコード生成"]
        G["LLM (ChatGPT/Gemini等)"] -->|"openpyxlコード生成"| H["Pythonソース (*_gen.py)"]
    end

    H --> I
    F --> I
    C --> I

    subgraph phase3["Phase 3: コード実行 + 3シートExcel"]
        I["executor.py"] --> J["1シート目: 変換結果"]
        I --> K["2シート目: フォント・カラー情報"]
        I --> L["3シート目: PDF画像"]
    end
    
    J & K & L --> M["出力: 3シートExcel (.xlsx)"]
```

## 各モジュールの役割

1. **`src/core/extractor.py`**
   - 【Phase 1】`pdfplumber` を使用してPDFからテキスト・座標(bbox)・フォント名・サイズ・色・罫線・矩形を抽出。JSON/Markdown形式で出力。

2. **`src/core/image_converter.py`**
   - 【Phase 1】`pdf2image`（poppler）でPDFを画像(PNG)に変換。3シート目用。

3. **`src/core/prompts.py`**
   - 【Phase 1】AIにopenpyxlベースのPythonコード（`generate(wb, ws)` 関数）を出力させるプロンプトを生成。座標変換ルールと利用可能なライブラリの情報を含む。

4. **`src/core/pipeline.py`**
   - 【Phase 1/3】各モジュールを統合。Phase1では解析→プロンプト生成、Phase3ではAIコード実行→3シートExcel生成。

5. **`src/core/executor.py`**
   - 【Phase 3】AI出力のPythonソースを`exec()`で実行し、1シート目を描画。2シート目（フォント・カラー一覧）と3シート目（PDF画像）を自動付加して3シートExcelを出力。

6. **`src/core/schema.py`**
   - 【共通】PDF抽出結果の型定義（Pydanticモデル）。

7. **`src/core/config.py`**
   - 【共通】方眼サイズ(4.96pt)、列数(120)、行数(176)等の設定値。

## 3シート構成

| シート | 名前 | 内容 | 生成元 |
|--------|------|------|--------|
| 1 | 変換結果 | AIがPDFレイアウトを再現 | AI生成Pythonコード |
| 2 | フォント・カラー情報 | 使用フォント一覧 + カラーコード一覧 | 自動（extractor抽出データ） |
| 3 | PDF画像 | PDFの各ページ画像 | 自動（pdf2image） |
