# スキャンPDF対応 設計書

## 概要

スキャンPDF（画像PDFとも呼ぶ）は、紙を複合機でスキャンしてPDF化したもので、
PDF内部に埋め込まれた**ラスタ画像**が本体である。
pdfplumber が依存するベクタ情報（テキストストリーム・パス）が存在しないため、
現行の `extract_pdf_data()` では何も取れない。

このドキュメントでは以下を整理する：

1. 罫線（borders）の取得可否と手法
2. テキスト取得および配置への絞り込み設計
3. 既存コードの流用範囲
4. 新規実装が必要なコンポーネント
5. 推奨方針と工数感

---

## 現行パイプラインとの差分

| 処理 | 通常PDF（ベクタ） | スキャンPDF |
|------|-----------------|------------|
| テキスト抽出 | pdfplumber `extract_words()` | OCR エンジン（外部依存） |
| テーブル検出 | pdfplumber `extract_tables()` | 不可（または画像処理） |
| 罫線（rects/edges）| pdfplumber `rects` / `lines` | 不可（または OpenCV） |
| グリッド計算 | `compute_grid_coords()` そのまま | **流用可**（入力形式さえ合わせる） |
| レイアウト生成 | `generate_layout()` そのまま | **流用可**（extracted_data スキーマを維持） |
| Excel描画 | `render_layout_to_xlsx()` そのまま | **流用可**（layout JSON は共通） |

---

## 罫線の取得可否

### 結論：**技術的には可能。ただし精度・工数のトレードオフが大きい**

スキャンPDFの罫線はラスタ画像上のピクセル列・行として存在する。
OpenCV の直線検出（Hough 変換）で座標を取り出すことは原理的に可能。

### 取得手順（実装した場合）

```
PDF ページ → png/jpeg 画像 (pdf2image)
    │
    ▼ OpenCV 前処理
グレースケール変換 → 二値化（Otsu / 適応的閾値）→ 細線化 or モルフォロジ処理
    │
    ▼ HoughLinesP
線分リスト（pixel座標）→ 水平/垂直に分類
    │
    ▼ 座標変換
pixel → PDF pt (= pixel × 72 / scan_dpi)
    │
    ▼ 既存パイプライン
h_edges / v_edges 形式に変換 → 以降は通常PDFと同一処理
```

### 精度上の問題点

| 問題 | 説明 |
|------|------|
| スキャンノイズ | 印刷の滲み・紙のシワが偽線を生む |
| 傾き・歪み | 1〜2° の傾きで水平線が複数セグメントに分断される |
| 罫線の欠損 | 薄い印刷や低解像度スキャン（150dpi以下）で検出漏れ |
| スタンプ・手書き | 斜め線・円弧が誤検出される |
| パラメータ感度 | 解像度・用紙種・プリンタによってチューニングが変わる |

### 実用上の判断

- **300dpi 以上の清書PDFなら罫線検出は実用レベルに近い**
- 傾き補正（`cv2.getRotationMatrix2D`）と長さフィルタ（既存 `filter_short_runs()`）を組み合わせることでノイズを大幅に抑制できる
- ただし「スタンプあり」「低品質スキャン」「印刷済みフォームへの手書き」では安定しない

---

## テキストのみモード（推奨先行実装）

罫線検出を省略し、OCR テキストの取得と配置のみを行うモード。
精度・工数・安定性のバランスが最も良い。

### データフロー

```
スキャンPDF
    │
    ▼
pdf2image.convert_from_path()      ← PDF → PIL.Image (per page)
    │  DPI: 300 推奨（150だと OCR 精度が落ちる）
    ▼
OCR エンジン（後述）
    │  出力: [{text, x0, top, x1, bottom, conf}]
    ▼
scan_extractor.py
    │  ・pixel → PDF pt 変換 (×72/dpi)
    │  ・低信頼スコアの語を除外（conf < 閾値）
    │  ・pdfplumber 互換の words リストに変換
    │  ・rects / h_edges / v_edges は空リストで渡す
    ▼
extracted_data（既存スキーマと同一）
    │
    ▼ 以降は既存パイプラインをそのまま使用
setup_grid_params()
compute_grid_coords()
generate_layout()          ← テキスト要素のみ生成される
render_layout_to_xlsx()    ← 罫線なし Excel が出力される
```

### 新規実装：`src/parser/scan_extractor.py`

pdfplumber 版の `extract_pdf_data()` と同じ戻り値スキーマを満たすモジュール。

```python
def extract_scan_pdf_data(pdf_path: str, dpi: int = 300, conf_threshold: float = 50.0) -> dict:
    """スキャンPDFをOCRで解析し、extract_pdf_data() と同じ構造で返す。

    Args:
        pdf_path: 対象PDFパス
        dpi: ラスタライズ解像度（300推奨）
        conf_threshold: OCR 信頼スコア閾値（0〜100）

    Returns:
        {
          'pages': [{
            'page_number': int,
            'width': float,   # PDF pt単位
            'height': float,
            'words': [...],   # pdfplumber互換
            'table_bboxes': [],
            'table_cells': [],
            'table_data': [],
            'table_data_raw': [],
            'table_col_x_positions': [],
            'table_row_y_positions': [],
            'rects': [],       # テキストのみモードでは常に空
            'h_edges': [],
            'v_edges': [],
          }]
        }
    """
```

**words エントリの変換ルール**

```python
# OCR 出力（pixel座標）
ocr_word = {'text': 'ABC', 'left': 120, 'top': 80, 'width': 60, 'height': 20, 'conf': 85.0}

# PDF pt 変換（DPI=300）
scale = 72 / dpi  # = 0.24
word = {
    'text': ocr_word['text'],
    'x0':   ocr_word['left'] * scale,
    'top':  ocr_word['top']  * scale,
    'x1':  (ocr_word['left'] + ocr_word['width'])  * scale,
    'bottom': (ocr_word['top'] + ocr_word['height']) * scale,
    'fontname': '',   # スキャンPDFではフォント情報なし
    'font_size': None,
}
```

### OCR エンジン選定

| エンジン | 精度（日本語） | コスト | 依存 | 推奨シーン |
|---------|------------|-------|------|-----------|
| **RapidOCR** ★採用 | ○ | 無料 | pip のみ | Python 3.13対応・環境依存なし |
| **Tesseract + pytesseract** | △（読みやすいフォントなら実用的） | 無料 | OS パッケージ必要 | ローカル・オフライン |
| **PaddleOCR** | ○ | 無料 | paddlepaddle（重い） | Python 3.11 環境限定 |
| **Google Vision API** | ◎ | 従量課金 | API キー | 高精度優先 |
| **Azure Document Intelligence** | ◎ | 従量課金 | API キー | テーブル構造も取れる場合あり |

**採用：RapidOCR**
- `pip install rapidocr onnxruntime` のみで導入可能
- 日本語モデルあり（`LangRec.JAPAN`）、初回実行時に自動ダウンロード
- 商用利用可（Apache 2.0）
- Python 3.13 対応、Windows/Linux 共通で動作
- PaddleOCR と同等モデルを ONNX Runtime 上で実行するため paddlepaddle 不要

### OCR 出力 → words 変換（RapidOCR）

```python
from rapidocr import RapidOCR
from rapidocr.utils.typings import LangRec

def _ocr_page(image: PIL.Image, dpi: int) -> list:
    ocr = RapidOCR(params={"Rec.lang_type": LangRec.JAPAN})
    result = ocr(np.array(image))
    if result is None or result.boxes is None:
        return []
    scale = 72 / dpi
    words = []
    for box, text, conf in zip(result.boxes, result.txts, result.scores):
        if conf < CONF_THRESHOLD:
            continue
        xs = [float(p[0]) for p in box]
        ys = [float(p[1]) for p in box]
        words.append({
            'text': text,
            'x0': min(xs) * scale,
            'top': min(ys) * scale,
            'x1': max(xs) * scale,
            'bottom': max(ys) * scale,
            'fontname': '',
        })
    return words
```

---

## 罫線検出モード（オプション拡張）

テキストのみモードで安定稼働を確認した後に追加する。

### 追加実装：`_detect_borders_from_image(image, dpi)`

```python
import cv2
import numpy as np

def _detect_borders_from_image(image: PIL.Image, dpi: int) -> tuple[list, list]:
    """画像から水平・垂直罫線を検出し h_edges / v_edges 形式で返す（PDF pt 単位）。"""
    scale = 72 / dpi
    gray = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2GRAY)

    # 前処理：二値化 → 反転（罫線を白、背景を黒に）
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)

    # 水平線カーネル：幅 page_width/30 程度の長いカーネルで横線のみ残す
    h_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (image.width // 30, 1))
    h_lines_img = cv2.morphologyEx(binary, cv2.MORPH_OPEN, h_kernel)

    # 垂直線カーネル：高さ page_height/30 程度の縦カーネルで縦線のみ残す
    v_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, image.height // 30))
    v_lines_img = cv2.morphologyEx(binary, cv2.MORPH_OPEN, v_kernel)

    # HoughLinesP で線分検出
    h_segs = cv2.HoughLinesP(h_lines_img, 1, np.pi/180, threshold=80,
                              minLineLength=image.width//20, maxLineGap=10)
    v_segs = cv2.HoughLinesP(v_lines_img, 1, np.pi/180, threshold=80,
                              minLineLength=image.height//20, maxLineGap=10)

    h_edges = _segments_to_h_edges(h_segs, scale)
    v_edges = _segments_to_v_edges(v_segs, scale)
    return h_edges, v_edges
```

検出後は `extracted_data` の `h_edges` / `v_edges` に設定するだけで、
既存の `_collect_edge_border_elements()` と `filter_short_runs()` がそのまま動く。

---

## 既存コードの流用マップ

```
src/
├── parser/
│   ├── pdf_extractor.py     ✗ スキャンPDFには使えない（pdfplumber依存）
│   └── scan_extractor.py    ★ 新規実装（同じ戻り値スキーマを持つ）
│
├── core/
│   ├── grid.py              ✅ 全関数そのまま流用
│   │     setup_grid_params()       — PDF寸法から用紙判定（変更不要）
│   │     compute_grid_coords()     — グリッド座標変換（入力スキーマ共通のため変更不要）
│   │
│   ├── layout.py            ✅ generate_layout() そのまま流用
│   │     _collect_text_elements()  — words リストがあれば動く
│   │     _collect_edge_border_elements() — h_edges/v_edges が空なら何も出ない（正常動作）
│   │
│   ├── edges.py             ✅ filter_short_runs() 罫線モード時に流用
│   ├── border_layout.py     ✅ 罫線モード時に流用
│   ├── text_layout.py       ✅ 全関数そのまま流用
│   ├── table_layout.py      △ テーブル検出ができないため事実上スキップ
│   ├── auto_layout_service.py △ ScanLayoutService として類似クラスを作る
│   ├── grid_config.py       ✅ そのまま流用
│   └── constants.py         ✅ そのまま流用
│
├── renderer/
│   └── excel.py             ✅ render_layout_to_xlsx() そのまま流用
│
└── utils/
    ├── font.py              △ normalize_font_name は不要（フォント情報なし）
    ├── text.py              ✅ join_word_texts / split_by_horizontal_gap 流用
    └── logger.py            ✅ そのまま流用
```

---

## 新規実装コンポーネント

### 1. `src/parser/scan_extractor.py`（必須）

```python
def extract_scan_pdf_data(pdf_path: str, dpi: int = 300) -> dict
    # pdf2image で各ページを画像化
    # OCR エンジンでテキスト抽出
    # pixel → PDF pt 変換
    # extracted_data スキーマで返す
```

**テキストのみモード**では `rects=[]`, `h_edges=[]`, `v_edges=[]` を返すだけ。
**罫線検出モード**では上記に加えて `_detect_borders_from_image()` を呼ぶ。

### 2. `src/core/scan_layout_service.py`（必須）

`AutoLayoutService` を参考に `ScanLayoutService` を実装。
`extract_pdf_data()` の代わりに `extract_scan_pdf_data()` を呼ぶ以外は
ほぼ同一の処理フロー。

```python
class ScanLayoutService:
    def run(self, pdf_path: str, in_base_dir: str = "data/in",
            grid_size: str = "small", mode: str = "text_only") -> dict:
        ...
```

`mode` パラメータ：
- `"text_only"` — OCR テキストのみ（罫線なし Excel）
- `"with_borders"` — 罫線検出を追加（OpenCV 依存、オプション）

### 3. `src/main.py` への `scan` コマンド追加（必須）

```python
parser.add_argument(
    "command", choices=["auto", "scan", "check"],
)
```

---

## 依存ライブラリの追加

### テキストのみモード（最小構成）

```
pypdfium2         # PDF → 画像レンダリング（pip のみ・システム依存なし）
rapidocr          # RapidOCR 本体
onnxruntime       # RapidOCR の推論エンジン
Pillow            # 画像処理
```

外部バイナリのインストールは不要。初回実行時に日本語モデル（約15MB）を自動ダウンロード。

### 罫線検出モード（追加・Phase 2）

```
opencv-python     # 画像処理・Hough 変換
```

---

## パッケージ安全性評価

サプライチェーン攻撃対策の観点から、使用パッケージごとにリスクを評価する。

### 評価サマリ

| パッケージ | メンテナー形態 | ライセンス | サプライチェーンリスク | 備考 |
|-----------|-------------|---------|------------------|------|
| **pypdfium2** | コミュニティ（PDFium は Google） | Apache-2.0 / BSD | 中 | pdfplumber の依存先でもある |
| **rapidocr** | RapidAI チーム | Apache 2.0 | 低〜中 | PaddleOCR モデルを ONNX 化したもの |
| **onnxruntime** | Microsoft | MIT | 低 | 業界標準・採用実績多数 |
| **Pillow** | チーム（4名） | MIT-CMU | 極低 | Trusted Publishing・OpenSSF 認定 |

---

### pypdfium2

**リスク：中（エンジン自体は Google Chrome と同一）**

- エンジン：PDFium（Google Chrome が使用する PDF レンダラー。BSD ライセンス）
- Pythonラッパー：pypdfium2-team（コミュニティ・ボランティア運営）
- ライセンス：Apache-2.0 / BSD ✅ 商用配布に制限なし
- リリース数：110回（活発に維持。最終更新 2026年5月）
- 採用実績：langchain・docling・nougat・**pdfplumber**（Sheetling の既存依存）
- 動作形態：ローカル完結。pip のみでインストール可（システム依存なし）

pdfplumber がすでに pypdfium2 を使用しているため、Sheetling の依存チェーンにすでに間接的に含まれている。追加のサプライチェーンリスクは最小限。

---

### rapidocr

**リスク：低〜中**

- メンテナー：RapidAI チーム（複数名）
- ライセンス：Apache 2.0 ✅
- 実態：PaddleOCR の学習済みモデルを ONNX 形式にエクスポートし、ONNX Runtime で推論するラッパー
- モデルは初回実行時に modelscope.cn から自動ダウンロード（日本語モデル約15MB）
- Python 3.8〜3.13 対応・Windows/Linux/macOS 共通

**注意**：モデルのダウンロード元が中国のクラウドサービス（modelscope.cn）であるため、エアギャップ環境では事前にモデルファイルを取得して配置する必要がある。

---

### onnxruntime

**リスク：低**

- メンテナー：Microsoft（OSS）
- ライセンス：MIT ✅
- GitHub スター：16,000超
- 採用実績：Azure ML・Hugging Face・PyTorch など業界標準
- Python 3.13 対応・各 OS 向けホイール提供

---

### Pillow

**リスク：極低（最も安全なパッケージのひとつ）**

- メンテナー：4名のチーム体制
- Trusted Publishing 有効（GitHub Actions 経由で署名付きリリース）
- OpenSSF Best Practices バッジ取得（OSS セキュリティ基準を満たすことの認定）
- Sigstore によるビルド透明性エントリ（インストール前に改ざん検証可能）
- ライセンス：MIT-CMU（PIL Software License）
- 月間ダウンロード：業界トップクラス（Pythonで画像を扱う事実上の標準）

---

### サプライチェーン攻撃への追加対策

#### 1. バージョン固定（必須）

`requirements.txt` でバージョンを固定することで、意図しない更新による悪意あるバージョンの混入を防ぐ。

```
pypdfium2==5.8.0
rapidocr==3.8.1
onnxruntime==1.26.0
Pillow==12.2.0
```

#### 2. エアギャップ環境でのモデル配布

インターネット接続のない環境では、モデルファイルを事前に取得して配置する。

```
# モデルのデフォルト保存先
{site-packages}/rapidocr/models/
  ├── ch_PP-OCRv4_det_mobile.onnx   # テキスト検出
  ├── ch_ppocr_mobile_v2.0_cls_mobile.onnx  # 向き分類
  └── japan_PP-OCRv4_rec_mobile.onnx  # 日本語認識
```

#### 3. 社内 PyPI ミラー（高セキュリティ現場向け）

インターネット接続のない環境では、事前に承認済みのパッケージのみを配置した社内ミラーを使用することでサプライチェーン攻撃を根本的に排除できる。

---

### 最終判断

現行の `pypdfium2 + rapidocr + onnxruntime + Pillow` 構成は、高セキュリティ現場における実用的な構成として妥当。ただし以下を事前に確認すること。

| 確認事項 | 対応 |
|---------|------|
| エアギャップ環境か | Yes → モデルファイルを事前配布・社内ミラー構築 |
| rapidocr の更新通知 | GitHub の Watch 設定で新リリースを監視 |

---

## 実装フェーズと工数感

| フェーズ | 内容 | 工数目安 |
|---------|------|---------|
| **Phase 1: テキストのみ** | `scan_extractor.py` + `ScanLayoutService` + `scan` コマンド | 3〜5日 |
| **Phase 2: 罫線検出** | `_detect_borders_from_image()` + チューニング | 5〜10日 |

Phase 1 だけでも「スキャンPDFのテキストをグリッドに配置したExcel」として十分実用的。
Phase 2 は品質が用紙・スキャン条件に強く依存するため、対象ドキュメントを絞った上で判断する。

---

## Q&A サマリ

**Q: スキャンPDFから罫線は取得できるか？**
A: 技術的には可能（OpenCV Hough 変換）。ただし精度はスキャン品質に依存し、ノイズ対策・パラメータチューニングが必要。先行して「テキストのみモード」で価値を確認してから投資判断を推奨する。

**Q: テキスト取得・配置のみに絞った変換は可能か？**
A: 可能。OCR 出力を pdfplumber 互換の words リストに変換すれば、グリッド計算・テキスト配置・Excel 描画は**既存コードをそのまま流用**できる。最大の変更点は `extract_pdf_data()` の置き換えのみ。

**Q: 既存コードの何割が流用できるか？**
A: テキストのみモードでは renderer・core/grid・core/layout・utils の大半（全体の約 70%）がそのまま動く。新規実装は `scan_extractor.py`（~150行）と `scan_layout_service.py`（~80行）程度。
