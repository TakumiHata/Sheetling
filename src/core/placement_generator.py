"""
配置命令リスト生成モジュール

JSONの要素データから、LLMがそのまま使える確定的な place_cell 引数リストを生成する。
座標変換（半開区間→閉区間）、テーブル構造へのスナップ、重複検証を Python 側で行い、
LLM の責務を「命令リストのコード化」のみに限定する。
"""
from dataclasses import dataclass, field
from src.utils.logger import get_logger

logger = get_logger(__name__)


@dataclass
class PlacementCommand:
    """place_cell に渡す引数を保持するデータクラス"""
    category: str  # "rect", "text_outside", "text_table"
    r1: int
    c1: int
    r2: int
    c2: int
    value: str = ""
    font_size: float = 10.0
    font_bold: bool = False
    alignment: str = "left"  # "left", "center", "right"
    comment: str = ""  # デバッグ用コメント


@dataclass
class TableStructure:
    """テーブルの列構造・行構造を保持する"""
    # 列区間のリスト: [(data_col_start, data_col_end), ...]
    col_ranges: list[tuple[int, int]] = field(default_factory=list)
    # 行区間のリスト: [(data_row_start, data_row_end), ...]
    row_ranges: list[tuple[int, int]] = field(default_factory=list)
    # テーブル領域
    table_row_min: int = 0
    table_row_max: int = 0
    table_col_min: int = 0
    table_col_max: int = 0
    # 罫線位置
    v_cols: list[int] = field(default_factory=list)
    h_rows: list[int] = field(default_factory=list)


@dataclass
class LineElement:
    """line要素の元座標を保持するデータクラス（罫線描画用）"""
    row_start: int
    row_end: int
    col_start: int
    col_end: int
    orientation: str  # "horizontal" or "vertical"


@dataclass
class PlacementResult:
    """配置命令リスト生成の結果"""
    commands: list[PlacementCommand] = field(default_factory=list)
    table_structures: list[TableStructure] = field(default_factory=list)
    line_elements: list[LineElement] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


class PlacementGenerator:
    """
    JSONの要素データから確定的な配置命令リストを生成する。

    処理フロー:
    1. line要素からテーブル構造を検出
    2. rect要素の座標変換（半開→閉）
    3. text要素の分類（テーブル内 / テーブル外）
    4. テーブル内要素のスナップ
    5. テーブル外要素の座標変換
    6. 全命令の重複検証・解消
    """

    def generate(self, json_data: dict) -> PlacementResult:
        """
        JSON全体から配置命令リストを生成する。

        Args:
            json_data: 圧縮済みJSON（pages配列を持つ）

        Returns:
            PlacementResult: 配置命令リストと付随情報
        """
        result = PlacementResult()

        for page_data in json_data.get("pages", []):
            elements = page_data.get("elements", [])

            # Step 1: line要素からテーブル構造を検出
            table_struct = self._detect_table_structure(elements)
            result.table_structures.append(table_struct)

            # Step 2: line要素の元座標を収集（罫線描画用）
            line_elems = self._collect_line_elements(elements)
            result.line_elements.extend(line_elems)

            # Step 3: rect要素の配置命令生成
            rect_commands = self._process_rects(elements)
            result.commands.extend(rect_commands)

            # Step 4-6: text要素の配置命令生成
            text_commands = self._process_texts(elements, table_struct)
            result.commands.extend(text_commands)

        # Step 6: 重複検証・解消
        overlap_warnings = self._validate_and_resolve_overlaps(result.commands)
        result.warnings.extend(overlap_warnings)

        logger.info(
            f"配置命令リスト生成完了: "
            f"{len(result.commands)}件の命令, "
            f"{len(result.warnings)}件の警告"
        )

        return result

    def _collect_line_elements(self, elements: list) -> list[LineElement]:
        """line要素の元座標を収集する（罫線描画用）。

        _detect_table_structure とは別に、各lineの向き・範囲をそのまま保持する。
        CodeGenerator がこの情報を使って draw_line を個別に生成する。
        """
        result = []
        lines = [e for e in elements if e.get("type") == "line"]
        seen = set()  # 重複排除

        for line in lines:
            bbox = line.get("grid_bbox", {})
            rs = bbox.get("row_start", 0)
            re = bbox.get("row_end", 0)
            cs = bbox.get("col_start", 0)
            ce = bbox.get("col_end", 0)

            # 縦線: 列幅が1以下で行の高さがある
            if ce - cs <= 1 and re - rs > 1:
                orientation = "vertical"
            # 横線: 行幅が1以下で列の幅がある
            elif re - rs <= 1 and ce - cs > 1:
                orientation = "horizontal"
            else:
                continue

            key = (rs, re, cs, ce, orientation)
            if key in seen:
                continue
            seen.add(key)

            result.append(LineElement(
                row_start=rs,
                row_end=re,
                col_start=cs,
                col_end=ce,
                orientation=orientation,
            ))

        return result

    def _detect_table_structure(self, elements: list) -> TableStructure:
        """line要素からテーブルの列・行構造を検出する。"""
        lines = [e for e in elements if e.get("type") == "line"]
        if not lines:
            return TableStructure()

        v_cols = set()
        h_rows = set()

        for line in lines:
            bbox = line.get("grid_bbox", {})
            rs = bbox.get("row_start", 0)
            re = bbox.get("row_end", 0)
            cs = bbox.get("col_start", 0)
            ce = bbox.get("col_end", 0)

            # 縦線: 列幅が1以下で行の高さがある
            if ce - cs <= 1 and re - rs > 1:
                v_cols.add(cs)
            # 横線: 行幅が1以下で列の幅がある
            elif re - rs <= 1 and ce - cs > 1:
                h_rows.add(rs)

        if not v_cols or not h_rows:
            return TableStructure()

        v_cols_sorted = sorted(v_cols)
        h_rows_sorted = sorted(h_rows)

        # 列区間の計算
        col_ranges = []
        for i in range(len(v_cols_sorted) - 1):
            left = v_cols_sorted[i]
            right = v_cols_sorted[i + 1]
            data_start = left + 1
            data_end = right - 1
            if data_start <= data_end:
                col_ranges.append((data_start, data_end))

        # 行区間の計算
        row_ranges = []
        for i in range(len(h_rows_sorted) - 1):
            top = h_rows_sorted[i]
            bottom = h_rows_sorted[i + 1]
            data_start = top + 1
            data_end = bottom - 1
            if data_start <= data_end:
                row_ranges.append((data_start, data_end))

        return TableStructure(
            col_ranges=col_ranges,
            row_ranges=row_ranges,
            table_row_min=h_rows_sorted[0],
            table_row_max=h_rows_sorted[-1],
            table_col_min=v_cols_sorted[0],
            table_col_max=v_cols_sorted[-1],
            v_cols=v_cols_sorted,
            h_rows=h_rows_sorted,
        )

    def _process_rects(self, elements: list) -> list[PlacementCommand]:
        """rect要素から配置命令を生成する。"""
        commands = []
        for elem in elements:
            if elem.get("type") != "rect":
                continue
            bbox = elem.get("grid_bbox", {})

            cmd = PlacementCommand(
                category="rect",
                r1=bbox["row_start"],
                c1=bbox["col_start"],
                r2=bbox["row_end"] - 1,  # 半開→閉
                c2=bbox["col_end"] - 1,  # 半開→閉
                comment=f"rect",
            )
            commands.append(cmd)
        return commands

    def _process_texts(self, elements: list, table: TableStructure) -> list[PlacementCommand]:
        """text要素を分類し、配置命令を生成する。

        テーブル内のテキストについては占有管理を行い、
        同一セル (row_range, col_range) に複数テキストが割り当てられるのを防ぐ。
        重複が発生した場合、先着テキストを優先し、後続はテーブル外扱いにする。
        """
        commands = []
        texts = [e for e in elements if e.get("type") == "text"]

        # テーブルセルの占有管理: key = (r1, c1, r2, c2), value = 占有テキスト
        occupied_cells: dict[tuple[int, int, int, int], str] = {}

        for elem in texts:
            bbox = elem.get("grid_bbox", {})
            text = elem.get("text", "")
            font_size = elem.get("font_size", 10.0)

            if self._is_inside_table(bbox, table):
                cmd = self._snap_to_table(elem, table)
                if cmd:
                    cell_key = (cmd.r1, cmd.c1, cmd.r2, cmd.c2)
                    if cell_key in occupied_cells:
                        # 同一セルが既に占有されている → テーブル外扱い（元のgrid_bboxで閉区間変換）
                        logger.debug(
                            f"セル占有による退避: '{text}' → セル {cell_key} は "
                            f"'{occupied_cells[cell_key]}' が占有済み"
                        )
                        cmd = self._convert_outside_text(elem)
                    else:
                        occupied_cells[cell_key] = text
            else:
                cmd = self._convert_outside_text(elem)

            if cmd:
                commands.append(cmd)

        return commands

    def _is_inside_table(self, bbox: dict, table: TableStructure) -> bool:
        """text要素がテーブル領域内にあるかを判定する。"""
        if not table.col_ranges:
            return False

        row_start = bbox.get("row_start", 0)
        col_start = bbox.get("col_start", 0)

        return (
            table.table_row_min <= row_start <= table.table_row_max
            and table.table_col_min <= col_start <= table.table_col_max
        )

    def _snap_to_table(self, elem: dict, table: TableStructure) -> PlacementCommand | None:
        """テーブル内のtext要素を列構造・行構造にスナップする。

        bboxの中心点を使ってスナップ先を決定し、r1==r2, c1/c2を同一列区間に
        収めることで、隣接行・列へのまたがりを防止する。
        """
        bbox = elem.get("grid_bbox", {})
        text = elem.get("text", "")
        font_size = elem.get("font_size", 10.0)

        col_start = bbox.get("col_start", 0)
        col_end = bbox.get("col_end", col_start + 1)
        row_start = bbox.get("row_start", 0)
        row_end = bbox.get("row_end", row_start + 1)

        # 中心点を計算（半開区間なのでcol_end/row_endはそのまま使う）
        col_center = (col_start + col_end) / 2.0
        row_center = (row_start + row_end) / 2.0

        # 列スナップ: 中心点が含まれる列区間を探す（整数変換して検索）
        snapped_c1, snapped_c2 = self._find_column_range(col_start, table)
        if snapped_c1 is None:
            # col_startで見つからない場合、中心点で再検索
            snapped_c1, snapped_c2 = self._find_column_range(int(col_center), table)
        if snapped_c1 is None:
            snapped_c1, snapped_c2 = self._nearest_column_range(int(col_center), table)
            if snapped_c1 is None:
                logger.warning(f"テーブル内text '{text}' の列スナップに失敗 (col={col_start})")
                return self._convert_outside_text(elem)

        # 行スナップ: 中心点が含まれる行区間を探す
        snapped_r1, snapped_r2 = self._find_row_range(int(row_center), table)
        if snapped_r1 is None:
            # 中心点で見つからない場合、最近傍を使用
            snapped_r1, snapped_r2 = self._nearest_row_range(int(row_center), table)
            if snapped_r1 is None:
                logger.warning(f"テーブル内text '{text}' の行スナップに失敗 (row={row_start})")
                return self._convert_outside_text(elem)

        alignment = self._guess_alignment(text, font_size)

        return PlacementCommand(
            category="text_table",
            r1=snapped_r1,
            c1=snapped_c1,
            r2=snapped_r2,
            c2=snapped_c2,
            value=text,
            font_size=font_size,
            font_bold=font_size >= 11.0,
            alignment=alignment,
            comment=f"table text (snapped from row={row_start}, col={col_start}, center=({row_center:.0f},{col_center:.0f}))",
        )

    def _convert_outside_text(self, elem: dict) -> PlacementCommand:
        """テーブル外のtext要素を閉区間に変換する。"""
        bbox = elem.get("grid_bbox", {})
        text = elem.get("text", "")
        font_size = elem.get("font_size", 10.0)

        alignment = self._guess_alignment(text, font_size)

        return PlacementCommand(
            category="text_outside",
            r1=bbox["row_start"],
            c1=bbox["col_start"],
            r2=bbox["row_end"] - 1,  # 半開→閉
            c2=bbox["col_end"] - 1,  # 半開→閉
            value=text,
            font_size=font_size,
            font_bold=font_size >= 18.0,
            alignment=alignment,
            comment=f"outside text",
        )

    def _find_column_range(self, col: int, table: TableStructure) -> tuple[int | None, int | None]:
        """col が含まれる列区間を返す。"""
        for c_start, c_end in table.col_ranges:
            if c_start <= col <= c_end:
                return c_start, c_end
        return None, None

    def _nearest_column_range(self, col: int, table: TableStructure) -> tuple[int | None, int | None]:
        """最も近い列区間を返す。"""
        if not table.col_ranges:
            return None, None
        best = None
        best_dist = float("inf")
        for c_start, c_end in table.col_ranges:
            mid = (c_start + c_end) / 2
            dist = abs(col - mid)
            if dist < best_dist:
                best_dist = dist
                best = (c_start, c_end)
        return best if best else (None, None)

    def _find_row_range(self, row: int, table: TableStructure) -> tuple[int | None, int | None]:
        """row が含まれる行区間を返す。"""
        for r_start, r_end in table.row_ranges:
            if r_start <= row <= r_end:
                return r_start, r_end
        return None, None

    def _nearest_row_range(self, row: int, table: TableStructure) -> tuple[int | None, int | None]:
        """最も近い行区間を返す。"""
        if not table.row_ranges:
            return None, None
        best = None
        best_dist = float("inf")
        for r_start, r_end in table.row_ranges:
            mid = (r_start + r_end) / 2
            dist = abs(row - mid)
            if dist < best_dist:
                best_dist = dist
                best = (r_start, r_end)
        return best if best else (None, None)

    def _guess_alignment(self, text: str, font_size: float) -> str:
        """テキスト内容とフォントサイズからアライメントを推定する。"""
        text_stripped = text.strip()

        # 金額・数値 → 右寄せ
        if self._is_numeric_value(text_stripped):
            return "right"

        # 大きいフォント（タイトル・ヘッダー） → 中央寄せ
        if font_size >= 18.0:
            return "center"

        # ヘッダー的なキーワード → 中央寄せ
        header_keywords = ["No.", "摘", "要", "数量", "標準価格", "見積価格", "合計金額",
                           "合", "計", "備", "考", "承認", "担当営業"]
        if text_stripped in header_keywords:
            return "center"

        # その他 → 左寄せ
        return "left"

    def _is_numeric_value(self, text: str) -> bool:
        """カンマ区切り数値かどうかを判定する。"""
        cleaned = text.replace(",", "").replace("\\", "").replace("¥", "").strip()
        if not cleaned:
            return False
        try:
            float(cleaned)
            return True
        except ValueError:
            return False

    def _validate_and_resolve_overlaps(self, commands: list[PlacementCommand]) -> list[str]:
        """
        全配置命令ペアの矩形重複をチェックし、解消を試みる。

        重複解消の戦略:
        - rect と text が重複 → 正常（rect が背景、textが前面）→ 無視
        - text 同士が重複 → 後の要素の座標を調整
        """
        warnings = []

        # rect は背景なので、text との重複は許容する
        non_rect_commands = [
            (i, cmd) for i, cmd in enumerate(commands)
            if cmd.category != "rect"
        ]

        for idx_a in range(len(non_rect_commands)):
            i, cmd_a = non_rect_commands[idx_a]
            for idx_b in range(idx_a + 1, len(non_rect_commands)):
                j, cmd_b = non_rect_commands[idx_b]

                if self._rects_overlap(cmd_a, cmd_b):
                    # 重複を解消する
                    resolved = self._resolve_overlap(cmd_a, cmd_b)
                    if resolved:
                        commands[j] = resolved
                        non_rect_commands[idx_b] = (j, resolved)
                        warnings.append(
                            f"重複を解消: '{cmd_b.value}' の座標を調整 "
                            f"({cmd_b.r1},{cmd_b.c1},{cmd_b.r2},{cmd_b.c2}) → "
                            f"({resolved.r1},{resolved.c1},{resolved.r2},{resolved.c2})"
                        )
                    else:
                        warnings.append(
                            f"⚠ 重複解消不能: '{cmd_a.value}' と '{cmd_b.value}' "
                            f"({cmd_a.r1},{cmd_a.c1},{cmd_a.r2},{cmd_a.c2}) vs "
                            f"({cmd_b.r1},{cmd_b.c1},{cmd_b.r2},{cmd_b.c2})"
                        )

        return warnings

    def _rects_overlap(self, a: PlacementCommand, b: PlacementCommand) -> bool:
        """2つの矩形が重複するかを判定する。"""
        return not (
            a.r2 < b.r1 or b.r2 < a.r1 or  # 行が離れている
            a.c2 < b.c1 or b.c2 < a.c1      # 列が離れている
        )

    def _resolve_overlap(
        self, keep: PlacementCommand, adjust: PlacementCommand
    ) -> PlacementCommand | None:
        """
        重複を解消するために adjust 側の座標を調整する。

        戦略:
        1. 同じ行で列が重複 → adjust の列開始を keep の列終了+1 に
        2. 同じ列で行が重複 → adjust の行開始を keep の行終了+1 に
        3. それ以外 → 解消不能
        """
        # 同じ行範囲で列が重複する場合
        if adjust.r1 >= keep.r1 and adjust.r1 <= keep.r2:
            # adjust を keep の右に押し出す
            new_c1 = keep.c2 + 1
            if new_c1 <= adjust.c2:
                return PlacementCommand(
                    category=adjust.category,
                    r1=adjust.r1, c1=new_c1,
                    r2=adjust.r2, c2=adjust.c2,
                    value=adjust.value,
                    font_size=adjust.font_size,
                    font_bold=adjust.font_bold,
                    alignment=adjust.alignment,
                    comment=f"{adjust.comment} (overlap resolved: c1 {adjust.c1}→{new_c1})",
                )

        # 同じ列範囲で行が重複する場合
        if adjust.c1 >= keep.c1 and adjust.c1 <= keep.c2:
            new_r1 = keep.r2 + 1
            if new_r1 <= adjust.r2:
                return PlacementCommand(
                    category=adjust.category,
                    r1=new_r1, c1=adjust.c1,
                    r2=adjust.r2, c2=adjust.c2,
                    value=adjust.value,
                    font_size=adjust.font_size,
                    font_bold=adjust.font_bold,
                    alignment=adjust.alignment,
                    comment=f"{adjust.comment} (overlap resolved: r1 {adjust.r1}→{new_r1})",
                )

        return None


def format_placement_commands(result: PlacementResult) -> str:
    """
    配置命令リストをLLM向けのMarkdownテキストにフォーマットする。

    出力形式:
    各命令を place_cell の引数として直接記載する。
    LLMはこのリストをそのまま main() 関数内のコードに展開するだけでよい。
    """
    lines = []

    # --- rect 配置 ---
    rect_cmds = [c for c in result.commands if c.category == "rect"]
    if rect_cmds:
        lines.append("### rect要素（背景色）の配置命令")
        lines.append("")
        lines.append("以下の `place_cell` 呼び出しを「--- 2. rect要素 ---」セクションにそのまま記述してください。")
        lines.append("")
        for i, cmd in enumerate(rect_cmds, 1):
            lines.append(
                f'{i}. `place_cell(ws, {cmd.r1}, {cmd.c1}, {cmd.r2}, {cmd.c2})`'
            )
        lines.append("")

    # --- text 配置（テーブル外） ---
    outside_cmds = [c for c in result.commands if c.category == "text_outside"]
    if outside_cmds:
        lines.append("### text要素（テーブル外）の配置命令")
        lines.append("")
        lines.append("以下の `place_cell` 呼び出しを「--- 3. text要素 ---」セクションに記述してください。")
        lines.append("Font名は `Meiryo` を使用してください。")
        lines.append("")
        for i, cmd in enumerate(outside_cmds, 1):
            bold_str = ", bold=True" if cmd.font_bold else ""
            align_map = {"left": "left", "center": "center", "right": "right"}
            align = align_map.get(cmd.alignment, "left")
            # 値のエスケープ
            escaped_value = cmd.value.replace('"', '\\"')
            lines.append(
                f'{i}. `place_cell(ws, {cmd.r1}, {cmd.c1}, {cmd.r2}, {cmd.c2}, '
                f'value="{escaped_value}", '
                f'font=Font(name="Meiryo", size={cmd.font_size}{bold_str}), '
                f'alignment=Alignment(horizontal="{align}", vertical="center", wrap_text=True))`'
            )
        lines.append("")

    # --- text 配置（テーブル内） ---
    table_cmds = [c for c in result.commands if c.category == "text_table"]
    if table_cmds:
        lines.append("### text要素（テーブル内）の配置命令")
        lines.append("")
        lines.append("以下の `place_cell` 呼び出しを「--- 3. text要素 ---」セクションのテーブル部分に記述してください。")
        lines.append("Font名は `Meiryo` を使用してください。")
        lines.append("")
        for i, cmd in enumerate(table_cmds, 1):
            bold_str = ", bold=True" if cmd.font_bold else ""
            align_map = {"left": "left", "center": "center", "right": "right"}
            align = align_map.get(cmd.alignment, "left")
            escaped_value = cmd.value.replace('"', '\\"')
            lines.append(
                f'{i}. `place_cell(ws, {cmd.r1}, {cmd.c1}, {cmd.r2}, {cmd.c2}, '
                f'value="{escaped_value}", '
                f'font=Font(name="Meiryo", size={cmd.font_size}{bold_str}), '
                f'alignment=Alignment(horizontal="{align}", vertical="center", wrap_text=True))`'
            )
        lines.append("")

    # --- テーブル罫線 ---
    if result.table_structures:
        for ts in result.table_structures:
            if ts.v_cols and ts.h_rows:
                lines.append("### テーブル罫線の描画命令")
                lines.append("")
                lines.append("以下の `draw_table_borders` 呼び出しを「--- 4. テーブル罫線 ---」セクションに記述してください。")
                lines.append("")
                lines.append(
                    f'`draw_table_borders(ws, '
                    f'v_cols={ts.v_cols}, '
                    f'h_rows={ts.h_rows}, '
                    f'row_min={ts.table_row_min}, '
                    f'row_max={ts.table_row_max}, '
                    f'col_min={ts.table_col_min}, '
                    f'col_max={ts.table_col_max})`'
                )
                lines.append("")

    # --- 警告 ---
    if result.warnings:
        lines.append("### ⚠ 座標調整ログ")
        lines.append("")
        for w in result.warnings:
            lines.append(f"- {w}")
        lines.append("")

    return "\n".join(lines)
