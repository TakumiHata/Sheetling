from dataclasses import dataclass, field


@dataclass
class GridConfig:
    """内部処理用仮想グリッドの設定"""
    unit_pt: float = 12.0  # 1セルの物理サイズ(pt) / 方眼の大きさ
    target_cols: int = 50  # A4印字可能領域幅(約595pt)に対する基準列数
    target_rows: int = 70  # A4高さ(約842pt)に対する基準行数


@dataclass
class ExcelConfig:
    """生成するExcelファイルの設定"""
    # A4の印字可能領域（約559pt）に120列を確実に収めるための調整値
    # 559pt / 120 = 約 4.65pt
    row_height_pt: float = 12.0
    # 12.0ptの正方形を作るための安全な文字幅設定
    col_width_chars: float = 2.0


@dataclass
class AppConfig:
    """アプリケーション全体の設定"""
    grid: GridConfig = field(default_factory=GridConfig)
    excel: ExcelConfig = field(default_factory=ExcelConfig)


# グローバルに利用可能な設定インスタンス
config = AppConfig()
