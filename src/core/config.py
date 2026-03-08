from dataclasses import dataclass, field


@dataclass
class GridConfig:
    """内部処理用仮想グリッドの設定"""
    unit_pt: float = 4.65  # 1セルの物理サイズ(pt) / 1.75mm。型ヒント(float)付き
    target_cols: int = 120 # A4幅(595pt)に対する基準列数
    target_rows: int = 176 # A4高さに対する基準行数


@dataclass
class ExcelConfig:
    """生成するExcelファイルの設定"""
    # A4の印字可能領域（約559pt）に120列を確実に収めるための調整値
    # 559pt / 120 = 約 4.65pt
    row_height_pt: float = 4.65
    # 4.65ptの正方形を作るための安全な文字幅設定
    col_width_chars: float = 0.75


@dataclass
class AppConfig:
    """アプリケーション全体の設定"""
    grid: GridConfig = field(default_factory=GridConfig)
    excel: ExcelConfig = field(default_factory=ExcelConfig)


# グローバルに利用可能な設定インスタンス
config = AppConfig()
