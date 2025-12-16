import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox

import ttkbootstrap as tb
from ttkbootstrap import ttk
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter


def pick_file(title: str, filetypes, parent=None):
    if parent is None:
        root = tk.Tk()
        root.withdraw()
        path = filedialog.askopenfilename(title=title, filetypes=filetypes)
        root.destroy()
        return path
    return filedialog.askopenfilename(parent=parent, title=title, filetypes=filetypes)


def pick_save_path(title: str, defaultextension: str, filetypes, parent=None):
    if parent is None:
        root = tk.Tk()
        root.withdraw()
        path = filedialog.asksaveasfilename(
            title=title,
            defaultextension=defaultextension,
            filetypes=filetypes,
        )
        root.destroy()
        return path
    return filedialog.asksaveasfilename(
        parent=parent,
        title=title,
        defaultextension=defaultextension,
        filetypes=filetypes,
    )


def _save_workbook_atomic(wb, out_path: str, parent=None) -> str:
    out_path = os.path.abspath(out_path)
    out_dir = os.path.dirname(out_path)
    if out_dir and not os.path.isdir(out_dir):
        os.makedirs(out_dir, exist_ok=True)

    base, ext = os.path.splitext(out_path)
    ext = ext or ".xlsx"
    tmp_path = f"{base}.tmp{ext}"

    while True:
        try:
            wb.save(tmp_path)
            os.replace(tmp_path, out_path)
            return out_path
        except PermissionError as e:
            if os.path.exists(tmp_path):
                try:
                    os.remove(tmp_path)
                except Exception:
                    pass
            if parent is None:
                raise
            retry = messagebox.askretrycancel(
                "保存に失敗",
                "出力先ファイルに書き込めません。\n"
                "Excelで出力先ファイルを開いている場合は閉じてから「再試行」を押してください。\n\n"
                f"出力先:\n{out_path}\n\n"
                f"詳細:\n{e}",
                parent=parent,
            )
            if not retry:
                raise
        finally:
            if os.path.exists(tmp_path):
                try:
                    os.remove(tmp_path)
                except Exception:
                    pass


def _try_extract_int(value):
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        try:
            return int(value)
        except Exception:
            return None
    if isinstance(value, str):
        s = value.strip()
        if not s:
            return None
        try:
            return int(s)
        except Exception:
            m = re.search(r"\d+", s)
            if not m:
                return None
            try:
                return int(m.group(0))
            except Exception:
                return None
    return None


def build_request_formulas(xlsx_path: str, out_path: str, cfg: dict, *, parent=None):
    sheet_name = cfg["sheet_name"]
    measure_no_col = cfg.get("measure_no_col", "A")
    measure_row_min = int(cfg.get("measure_row_min", 9))
    measure_row_max = int(cfg.get("measure_row_max", 74))
    measure_row_step = int(cfg.get("measure_row_step", 2))
    summary_row_min = int(cfg.get("summary_row_min", 9))
    summary_row_max = int(cfg.get("summary_row_max", 159))
    summary_row_step = int(cfg.get("summary_row_step", 2))
    formula_arg_sep = str(cfg.get("formula_arg_sep", ",")).strip() or ","
    flag_col_start = column_index_from_string("I")
    flag_col_end = column_index_from_string("SN")

    tool_start_row = int(cfg.get("tool_start_row", 75))
    tool_name_col = cfg.get("tool_name_col", "B")
    tool_row_step = int(cfg.get("tool_row_step", 2))

    tools = cfg["tools"]
    tool_to_measure_nos = cfg["tool_to_measure_nos"]

    wb = load_workbook(xlsx_path)
    wb_values = load_workbook(xlsx_path, data_only=True)
    if sheet_name not in wb.sheetnames:
        raise ValueError(
            f"シート '{sheet_name}' が見つかりません。存在: {wb.sheetnames}"
        )

    ws = wb[sheet_name]
    ws_values = wb_values[sheet_name]

    # 1) 測定No(A列の整数) → 行番号
    no_col = column_index_from_string(measure_no_col)
    measure_no_to_row = {}
    max_r = min(measure_row_max, ws.max_row or measure_row_max)
    for r in range(measure_row_min, max_r + 1):
        v = ws_values.cell(r, no_col).value
        if v is None:
            v = ws.cell(r, no_col).value
        if v is None:
            continue
        no = _try_extract_int(v)
        if no is None:
            continue
        measure_no_to_row[no] = r

    # 2) 工具名を書き込み & 工具名 → 行番号
    tool_row = {}
    tool_name_c = column_index_from_string(tool_name_col)
    r = tool_start_row
    for tool in tools:
        ws.cell(r, tool_name_c).value = tool
        tool_row[tool] = r
        r += tool_row_step

    # 3) 測定行ごとに参照すべき工具行（逆引き）
    measure_row_to_tool_rows = {}
    missing_nos = []
    for tool, nos in tool_to_measure_nos.items():
        if tool not in tool_row:
            # tools にない工具名が map にいたら無視（またはエラーにしたければ raise）
            continue
        tr = tool_row[tool]
        for no in nos:
            n = _try_extract_int(no)
            if n is None:
                continue
            mr = measure_no_to_row.get(n)
            if mr is None:
                # 指定した測定Noがシートに無い場合は無視（またはログ化）
                missing_nos.append((tool, n))
                continue
            measure_row_to_tool_rows.setdefault(mr, []).append(tr)

    # 4) 依頼関数を投入
    written = 0
    for col_idx in range(flag_col_start, flag_col_end + 1):  # SN列まで含める
        col_letter = get_column_letter(col_idx)
        # 指定列の1行目（例: I1）に集計式を投入
        # 修正版: =SUMPRODUCT((I9:I159<>"")*1,(MOD(ROW(I9:I159)-ROW(I9),2)=0)*1,NOT(ISFORMULA(I9:I159))*1)
        # ISFORMULAが使えない場合の代替: =SUMPRODUCT((I9:I159<>"")*1,(MOD(ROW(I9:I159)-ROW(I9),2)=0)*1,LEFT(CELL("format",I9:I159),1)="="*0)
        if summary_row_step <= 0:
            raise ValueError("summary_row_step は 1 以上を指定してください。")
        if summary_row_min <= 0 or summary_row_max <= 0:
            raise ValueError(
                "summary_row_min/summary_row_max は 1 以上を指定してください。"
            )
        if summary_row_min > summary_row_max:
            raise ValueError(
                "summary_row_min は summary_row_max 以下を指定してください。"
            )

        # ISFORMULA関数を完全に削除し、単純な数式に変更
        # 空でないセルかつ指定ステップの行をカウント
        # 測定行の範囲（measure_row_min〜measure_row_max）のみをカウントするように修正
        sumproduct_args = [
            f'--({col_letter}{measure_row_min}:{col_letter}{measure_row_max}<>"")',
            f"--(MOD(ROW({col_letter}{measure_row_min}:{col_letter}{measure_row_max})-ROW({col_letter}{measure_row_min}),{measure_row_step})=0)",
        ]

        # SUMPRODUCT関数を使用
        sumproduct_formula = f"=SUMPRODUCT({formula_arg_sep.join(sumproduct_args)})"

        # 1行目と2行目に別々の数式を設定
        # 1行目: ROW(J9) を基準
        ws.cell(1, col_idx).value = sumproduct_formula

        # 2行目: ROW(J10) を基準にした数式を作成
        row2_base_row = measure_row_min + 1  # 9 + 1 = 10
        sumproduct_args_row2 = [
            f'--({col_letter}{measure_row_min}:{col_letter}{measure_row_max}<>"")',
            f"--(MOD(ROW({col_letter}{measure_row_min}:{col_letter}{measure_row_max})-ROW({col_letter}{row2_base_row}),{measure_row_step})=0)",
        ]
        sumproduct_formula_row2 = (
            f"=SUMPRODUCT({formula_arg_sep.join(sumproduct_args_row2)})"
        )
        ws.cell(2, col_idx).value = sumproduct_formula_row2

        # デバッグ用に生成した数式をログ出力（実際の使用ではコメントアウトしてもよい）
        print(f"列 {col_letter} 1行目: {sumproduct_formula}")
        print(f"列 {col_letter} 2行目: {sumproduct_formula_row2}")

        for mr, tool_rows in measure_row_to_tool_rows.items():
            # 例: OR(N$75<>"",N$77<>"")
            conds = ",".join([f'{col_letter}${tr}<>""' for tr in tool_rows])
            ws.cell(mr, col_idx).value = f'=IF(OR({conds}),"依頼","")'
            written += 1

    if written == 0:
        raise ValueError(
            "依頼セルが1件も書き込まれませんでした。\n\n"
            f"- 読み取れた測定No件数: {len(measure_no_to_row)}\n"
            f"- 工具件数: {len(tools)}\n"
            f"- 逆引き対象の測定行件数: {len(measure_row_to_tool_rows)}\n"
            "- 出力列: I〜SN（固定）\n"
            f"- 指定Noが見つからない例(先頭10件): {missing_nos[:10]}\n\n"
            "原因候補:\n"
            "1) 測定No列/行範囲が実シートと違う\n"
            "2) 測定Noが数式で、キャッシュ値が未保存（Excelで一度保存してから再実行）\n"
            "3) tool_to_measure_nos のNoがシートに存在しない"
        )

    try:
        wb.calculation.calcMode = "auto"
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.calcOnSave = True
        wb.calculation.forceFullCalc = True
    except Exception:
        pass

    return _save_workbook_atomic(wb, out_path, parent=parent)


def _parse_int_list(text: str):
    if not text.strip():
        return []
    parts = [p.strip() for p in text.replace("、", ",").replace(" ", ",").split(",")]
    nums = []
    for p in parts:
        if not p:
            continue
        nums.append(int(p))
    return nums


def write_measurement_not_required(
    xlsx_path: str,
    out_path: str,
    cfg: dict,
    target_row: int = 69,
    target_nos: list = None,
    *,
    parent=None,
):
    """
    指定したB列に"測定不要"を書き込み、指定したNo.の行のI列に条件付きで"-"を書き込む

    Args:
        xlsx_path: 入力Excelファイルのパス
        out_path: 出力Excelファイルのパス
        cfg: 設定辞書（sheet_name, measure_no_col, measure_row_min, measure_row_max を含む）
        target_row: B列に"測定不要"を書き込む行番号（デフォルト: 69）
        target_nos: I列に"-"を書き込む対象のNo.リスト（A列から検索）
        parent: 親ウィンドウ（エラーメッセージ表示用）

    Returns:
        保存されたファイルのパス
    """
    if target_nos is None:
        target_nos = []

    sheet_name = cfg.get("sheet_name", "工程内検査シート")
    measure_no_col = cfg.get("measure_no_col", "A")
    measure_row_min = int(cfg.get("measure_row_min", 9))
    measure_row_max = int(cfg.get("measure_row_max", 74))

    wb = load_workbook(xlsx_path)
    wb_values = load_workbook(xlsx_path, data_only=True)

    if sheet_name not in wb.sheetnames:
        raise ValueError(
            f"シート '{sheet_name}' が見つかりません。存在: {wb.sheetnames}"
        )

    ws = wb[sheet_name]
    ws_values = wb_values[sheet_name]

    # 1) 測定No(A列の整数) → 行番号のマッピングを作成
    no_col = column_index_from_string(measure_no_col)
    measure_no_to_row = {}
    max_r = min(measure_row_max, ws.max_row or measure_row_max)
    for r in range(measure_row_min, max_r + 1):
        v = ws_values.cell(r, no_col).value
        if v is None:
            v = ws.cell(r, no_col).value
        if v is None:
            continue
        no = _try_extract_int(v)
        if no is None:
            continue
        measure_no_to_row[no] = r

    # 2) 指定行のB列に"測定不要"を書き込み
    b_col = column_index_from_string("B")
    ws.cell(target_row, b_col).value = "測定不要"

    # 3) 指定したNo.の行のI列に条件付きで"-"を書き込む
    # 条件: 指定行（target_row）のI列に値が入ると、全ての指定No.の行のI列に"-"が入る
    i_col = column_index_from_string("I")
    target_i_cell = get_column_letter(i_col) + str(target_row)

    written_count = 0
    for no in target_nos:
        row = measure_no_to_row.get(no)
        if row is None:
            # 指定したNo.が見つからない場合はスキップ（または警告を出す）
            continue

        # 条件付き数式: 指定行のI列に値が入ると"-"を表示
        # =IF(INDIRECT("I"&69)<>"","-","")
        formula = f'=IF({target_i_cell}<>"","-","")'
        ws.cell(row, i_col).value = formula
        written_count += 1

    if written_count == 0 and target_nos:
        raise ValueError(
            f"指定したNo.の行が見つかりませんでした。\n"
            f"- 指定したNo.: {target_nos}\n"
            f"- 読み取れた測定No件数: {len(measure_no_to_row)}\n"
            f"- 測定No範囲: {measure_row_min}〜{measure_row_max}行目"
        )

    return _save_workbook_atomic(wb, out_path, parent=parent)


class ConfigEditor(tb.Window):
    def __init__(self):
        super().__init__(themename="darkly")
        self.title("検査シート 設定作成")
        self.geometry("820x640")
        try:
            self.state("zoomed")
        except Exception:
            pass

        self.vars = {
            "sheet_name": tk.StringVar(value="工程内検査シート"),
            "measure_no_col": tk.StringVar(value="A"),
            "measure_row_min": tk.IntVar(value=9),
            "measure_row_max": tk.IntVar(value=74),
            "measure_row_step": tk.IntVar(value=2),
            "summary_row_min": tk.IntVar(value=9),
            "summary_row_max": tk.IntVar(value=159),
            "summary_row_step": tk.IntVar(value=2),
            "formula_arg_sep": tk.StringVar(value=","),
            "tool_start_row": tk.IntVar(value=75),
            "tool_name_col": tk.StringVar(value="B"),
            "tool_row_step": tk.IntVar(value=2),
            "not_required_row": tk.StringVar(value="69"),
            "not_required_nos": tk.StringVar(value=""),
        }

        self.selected_xlsx = tk.StringVar(value="")
        self.preview_title = tk.StringVar(value="プレビュー (未読み込み)")

        self._build_ui()

    def _build_ui(self):
        header_frame = ttk.Frame(self)
        header_frame.pack(fill="x", padx=10, pady=(10, 0))
        ttk.Button(header_frame, text="ヘルプ", command=self._show_help).pack(side="right")

        main = ttk.Frame(self, padding=10)
        main.pack(fill="both", expand=True)

        # 元Excel読み込み＆プレビュー
        source_frame = ttk.LabelFrame(main, text="元Excelとプレビュー", padding=10)
        source_frame.pack(fill="both", expand=True)

        src_row = ttk.Frame(source_frame)
        src_row.pack(fill="x", pady=(0, 8))
        ttk.Label(src_row, text="元Excelファイル").pack(side="left")
        ttk.Entry(src_row, textvariable=self.selected_xlsx, width=60, state="readonly").pack(side="left", padx=6)
        ttk.Button(src_row, text="Excelを選択", command=self._load_preview).pack(side="left")
        ttk.Button(src_row, text="プレビュー更新", command=self._render_preview).pack(side="left", padx=6)
        ttk.Label(src_row, textvariable=self.preview_title).pack(side="right")

        preview_container = ttk.Frame(source_frame)
        preview_container.pack(fill="both", expand=True)
        self.preview_columns = ("A", "B", "G", "K")
        self.preview_tree = ttk.Treeview(
            preview_container, columns=self.preview_columns, show="headings", height=8
        )
        vsb = ttk.Scrollbar(preview_container, orient="vertical", command=self.preview_tree.yview)
        self.preview_tree.configure(yscrollcommand=vsb.set)
        self.preview_tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="left", fill="y")

        basic = ttk.LabelFrame(main, text="基本設定", padding=10)
        basic.pack(fill="x", pady=(10, 0))

        basic_inner = ttk.Frame(basic)
        basic_inner.pack(fill="x")
        basic_inner.columnconfigure(0, weight=1)
        basic_inner.columnconfigure(1, weight=1)

        basic_left = ttk.Frame(basic_inner)
        basic_left.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        basic_right = ttk.LabelFrame(basic_inner, text="測定不要書き込み設定", padding=10)
        basic_right.grid(row=0, column=1, sticky="nsew")

        def add_field(row, col, label, key, width=12):
            col_offset = col * 2
            ttk.Label(basic_left, text=label).grid(row=row, column=col_offset, sticky="w", padx=(0, 8), pady=3)
            ttk.Entry(basic_left, textvariable=self.vars[key], width=width).grid(row=row, column=col_offset + 1, sticky="w", padx=(0, 20), pady=3)

        add_field(0, 0, "シート名", "sheet_name", width=25)
        add_field(1, 0, "測定No列", "measure_no_col")
        add_field(2, 0, "測定行(min)", "measure_row_min")
        add_field(3, 0, "測定行(max)", "measure_row_max")
        add_field(4, 0, "測定行ステップ", "measure_row_step")
        add_field(5, 0, "集計行(min)", "summary_row_min")
        add_field(6, 0, "集計行(max)", "summary_row_max")
        add_field(7, 0, "集計行ステップ", "summary_row_step")
        add_field(0, 1, "数式区切り(, / ;)", "formula_arg_sep")
        add_field(1, 1, "工具開始行", "tool_start_row")
        add_field(2, 1, "工具名列", "tool_name_col")
        add_field(3, 1, "工具行ステップ", "tool_row_step")

        ttk.Label(basic, text="出力列: I～SN（固定）").pack(anchor="w", pady=3)
        ttk.Label(basic_right, text="B列に書き込む行番号:").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=5)
        ttk.Entry(basic_right, textvariable=self.vars["not_required_row"], width=15).grid(row=0, column=1, sticky="w", padx=(0, 30), pady=5)
        ttk.Label(basic_right, text="I列に'-'を入れるNo.(カンマ区切り):").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=5)
        ttk.Entry(basic_right, textvariable=self.vars["not_required_nos"], width=40).grid(row=1, column=1, sticky="ew", padx=(0, 10), pady=5)
        basic_right.columnconfigure(1, weight=1)

        tools_frame = ttk.LabelFrame(main, text="工具と測定No対応", padding=10)
        tools_frame.pack(fill="both", expand=True, pady=(10, 0))
        self.tools_tree = ttk.Treeview(tools_frame, columns=("tool", "nos"), show="headings", height=12)
        self.tools_tree.heading("tool", text="工具名")
        self.tools_tree.heading("nos", text="測定No(カンマ区切り)")
        self.tools_tree.column("tool", width=220, anchor="w")
        self.tools_tree.column("nos", width=420, anchor="w")
        self.tools_tree.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(tools_frame, orient="vertical", command=self.tools_tree.yview)
        self.tools_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        btns = ttk.Frame(main)
        btns.pack(fill="x", pady=(10, 0))
        ttk.Button(btns, text="工具追加", command=self._add_tool_dialog).pack(side="left")
        ttk.Button(btns, text="選択編集", command=self._edit_selected_tool).pack(side="left", padx=5)
        ttk.Button(btns, text="選択削除", command=self._delete_selected_tool).pack(side="left")
        ttk.Button(btns, text="この設定でExcel生成", command=self._run_build).pack(side="right")

        if not self.tools_tree.get_children():
            self._insert_tool("前挽き(サンプル)", "1, 5, 10")

    def _load_preview(self):
        path = filedialog.askopenfilename(
            parent=self,
            title="元の検査シート（xlsx）を選択",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not path:
            return
        self.selected_xlsx.set(path)
        self._render_preview()

    def _render_preview(self):
        path = self.selected_xlsx.get().strip()
        if not path:
            messagebox.showinfo("読み込み待ち", "先にExcelファイルを選択してください。", parent=self)
            return

        sheet_name = self.vars["sheet_name"].get().strip() or "工程内検査シート"
        try:
            wb = load_workbook(path, data_only=True)
        except Exception as e:
            messagebox.showerror("読み込み失敗", f"Excelを開けませんでした。\n{e}", parent=self)
            return

        if sheet_name not in wb.sheetnames:
            messagebox.showerror(
                "シートなし",
                f"シート「{sheet_name}」が見つかりません。\nシート名を確認して再度プレビューしてください。",
                parent=self,
            )
            return

        ws = wb[sheet_name]
        preview_col_indices = [column_index_from_string(c) for c in self.preview_columns]

        if not ws.max_row or not ws.max_column:
            messagebox.showinfo("シートが空です", "表示できるデータがありません。", parent=self)
            return

        self.preview_tree.configure(columns=self.preview_columns)
        for col in self.preview_columns:
            self.preview_tree.heading(col, text=col)
            self.preview_tree.column(col, width=120, anchor="w")

        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)

        row_count = 0
        for r in range(1, (ws.max_row or 0) + 1):
            a_val = ws.cell(r, preview_col_indices[0]).value
            if a_val is None or (isinstance(a_val, str) and not a_val.strip()):
                break

            row_vals = []
            for c in preview_col_indices:
                v = ws.cell(r, c).value
                row_vals.append("" if v is None else str(v))

            self.preview_tree.insert("", "end", values=row_vals)
            row_count += 1

        if row_count == 0:
            messagebox.showinfo(
                "データなし",
                "A列が空のため表示できる行がありません。",
                parent=self,
            )
            return

        self.preview_title.set(f"{sheet_name} プレビュー（{row_count}行: A列が空になるまで）")

    def _insert_tool(self, tool_name: str, nos_text: str):
        self.tools_tree.insert("", "end", values=(tool_name, nos_text))

    def _tool_dialog(self, title, init_tool="", init_nos=""):
        win = tk.Toplevel(self)
        win.title(title)
        win.transient(self)
        win.grab_set()

        tool_var = tk.StringVar(value=init_tool)
        nos_var = tk.StringVar(value=init_nos)

        frm = ttk.Frame(win, padding=10)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="工具名").grid(row=0, column=0, sticky="w", pady=4)
        tool_entry = ttk.Entry(frm, textvariable=tool_var, width=30)
        tool_entry.grid(row=0, column=1, sticky="w", pady=4)

        ttk.Label(frm, text="測定No(カンマ区切り)").grid(
            row=1, column=0, sticky="w", pady=4
        )
        ttk.Entry(frm, textvariable=nos_var, width=40).grid(
            row=1, column=1, sticky="w", pady=4
        )

        result = {"ok": False}

        def on_ok():
            tool = tool_var.get().strip()
            if not tool:
                messagebox.showwarning(
                    "入力不足", "工具名を入力してください。", parent=win
                )
                return
            try:
                _parse_int_list(nos_var.get())
            except Exception:
                messagebox.showwarning(
                    "入力エラー",
                    "測定Noは整数のカンマ区切りで入力してください。",
                    parent=win,
                )
                return
            result["ok"] = True
            result["tool"] = tool
            result["nos"] = nos_var.get().strip()
            win.destroy()

        def on_cancel():
            win.destroy()

        bfrm = ttk.Frame(frm)
        bfrm.grid(row=2, column=0, columnspan=2, sticky="e", pady=(8, 0))
        ttk.Button(bfrm, text="OK", command=on_ok).pack(side="right")
        ttk.Button(bfrm, text="キャンセル", command=on_cancel).pack(
            side="right", padx=5
        )

        tool_entry.focus_set()
        self.wait_window(win)
        return result

    def _add_tool_dialog(self):
        res = self._tool_dialog("工具追加")
        if res.get("ok"):
            self._insert_tool(res["tool"], res["nos"])

    def _edit_selected_tool(self):
        sel = self.tools_tree.selection()
        if not sel:
            messagebox.showinfo(
                "選択なし", "編集する行を選択してください。", parent=self
            )
            return
        item = sel[0]
        tool, nos = self.tools_tree.item(item, "values")
        res = self._tool_dialog("工具編集", init_tool=tool, init_nos=nos)
        if res.get("ok"):
            self.tools_tree.item(item, values=(res["tool"], res["nos"]))

    def _delete_selected_tool(self):
        sel = self.tools_tree.selection()
        if not sel:
            return
        if not messagebox.askyesno(
            "削除確認", "選択した工具を削除しますか？", parent=self
        ):
            return
        for item in sel:
            self.tools_tree.delete(item)

    def _gather_cfg(self):
        try:
            tools = []
            tool_to_measure_nos = {}
            for item in self.tools_tree.get_children():
                tool, nos_text = self.tools_tree.item(item, "values")
                tools.append(tool)
                tool_to_measure_nos[tool] = _parse_int_list(nos_text)

            if not tools:
                raise ValueError("工具が1件もありません。")

            cfg = {
                "sheet_name": self.vars["sheet_name"].get().strip(),
                "measure_no_col": self.vars["measure_no_col"].get().strip().upper(),
                "measure_row_min": int(self.vars["measure_row_min"].get()),
                "measure_row_max": int(self.vars["measure_row_max"].get()),
                "measure_row_step": int(self.vars["measure_row_step"].get()),
                "summary_row_min": int(self.vars["summary_row_min"].get()),
                "summary_row_max": int(self.vars["summary_row_max"].get()),
                "summary_row_step": int(self.vars["summary_row_step"].get()),
                "formula_arg_sep": self.vars["formula_arg_sep"].get().strip(),
                "tool_start_row": int(self.vars["tool_start_row"].get()),
                "tool_name_col": self.vars["tool_name_col"].get().strip().upper(),
                "tool_row_step": int(self.vars["tool_row_step"].get()),
                "tools": tools,
                "tool_to_measure_nos": tool_to_measure_nos,
            }
            if not cfg["sheet_name"]:
                raise ValueError("シート名が空です。")
            return cfg
        except Exception as e:
            raise ValueError(f"設定の取得に失敗: {e}")

    def _run_build(self):
        try:
            cfg = self._gather_cfg()
        except Exception as e:
            messagebox.showerror("失敗", str(e), parent=self)
            return

        xlsx = self.selected_xlsx.get().strip()
        if not xlsx:
            messagebox.showinfo(
                "ファイル未選択",
                "先に「Excelを選択」で元ファイルを読み込んでください。",
                parent=self,
            )
            self._load_preview()
            xlsx = self.selected_xlsx.get().strip()
            if not xlsx:
                return
        out_path = pick_save_path(
            "出力先（生成したxlsx）を保存",
            ".xlsx",
            [("Excel", "*.xlsx")],
            parent=self,
        )
        if not out_path:
            return

        try:
            saved_path = build_request_formulas(xlsx, out_path, cfg, parent=self)
        except Exception as e:
            messagebox.showerror("失敗", str(e), parent=self)
            return

        # 測定不要書き込み設定が入力されている場合は実行
        try:
            not_required_nos_text = self.vars["not_required_nos"].get().strip()
            if not_required_nos_text:
                try:
                    target_row = int(self.vars["not_required_row"].get().strip())
                    if target_row < 1:
                        raise ValueError("行番号は1以上を指定してください。")
                except ValueError as e:
                    messagebox.showerror(
                        "入力エラー", f"行番号の入力が不正です。\n{e}", parent=self
                    )
                    return

                try:
                    target_nos = _parse_int_list(not_required_nos_text)
                except Exception as e:
                    messagebox.showerror(
                        "入力エラー",
                        f"測定Noの入力が不正です。\n{e}",
                        parent=self,
                    )
                    return

                if target_nos:
                    # 同じファイルに対して測定不要書き込みを実行
                    saved_path = write_measurement_not_required(
                        saved_path,  # 既に生成されたファイルを使用
                        out_path,  # 同じ出力先に上書き
                        cfg,
                        target_row=target_row,
                        target_nos=target_nos,
                        parent=self,
                    )
        except Exception as e:
            # 測定不要書き込みでエラーが発生しても、メイン処理は成功しているので警告のみ
            messagebox.showwarning(
                "警告",
                f"Excel生成は完了しましたが、測定不要書き込みでエラーが発生しました:\n{e}",
                parent=self,
            )
            return

        messagebox.showinfo("完了", f"生成しました:\n{saved_path}", parent=self)

    def _run_write_not_required(self):
        """測定不要書き込み機能の実行"""
        try:
            cfg = self._gather_cfg()
        except Exception as e:
            messagebox.showerror("失敗", str(e), parent=self)
            return

        # UI内の入力フィールドから値を取得
        try:
            target_row = int(self.vars["not_required_row"].get().strip())
            if target_row < 1:
                raise ValueError("行番号は1以上を指定してください。")
        except ValueError as e:
            messagebox.showerror(
                "入力エラー", f"行番号の入力が不正です。\n{e}", parent=self
            )
            return

        try:
            target_nos = _parse_int_list(self.vars["not_required_nos"].get())
        except Exception as e:
            messagebox.showerror(
                "入力エラー",
                f"測定Noの入力が不正です。\n{e}",
                parent=self,
            )
            return

        if not target_nos:
            messagebox.showwarning(
                "入力不足", "I列に'-'を入れるNo.を入力してください。", parent=self
            )
            return

        xlsx = self.selected_xlsx.get().strip()
        if not xlsx:
            messagebox.showinfo(
                "ファイル未選択",
                "先に「Excelを選択」で元ファイルを読み込んでください。",
                parent=self,
            )
            self._load_preview()
            xlsx = self.selected_xlsx.get().strip()
            if not xlsx:
                return
        out_path = pick_save_path(
            "出力先（生成したxlsx）を保存",
            ".xlsx",
            [("Excel", "*.xlsx")],
            parent=self,
        )
        if not out_path:
            return

        try:
            saved_path = write_measurement_not_required(
                xlsx,
                out_path,
                cfg,
                target_row=target_row,
                target_nos=target_nos,
                parent=self,
            )
        except Exception as e:
            messagebox.showerror("失敗", str(e), parent=self)
            return
        messagebox.showinfo("完了", f"書き込み完了:\n{saved_path}", parent=self)

    def _show_help(self):
        """ヘルプモーダルを表示"""
        help_window = tk.Toplevel(self)
        help_window.title("使い方")
        help_window.transient(self)
        help_window.grab_set()
        help_window.geometry("600x500")

        # スクロール可能なフレーム
        canvas = tk.Canvas(help_window, highlightthickness=0, bd=0)
        scrollbar = ttk.Scrollbar(help_window, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # コンテンツ
        content_frame = ttk.Frame(scrollable_frame, padding=20)
        content_frame.pack(fill="both", expand=True)

        # バージョン情報
        version_frame = ttk.LabelFrame(content_frame, text="バージョン情報", padding=10)
        version_frame.pack(fill="x", pady=(0, 15))

        ttk.Label(version_frame, text="Version: 0.5.0", font=("", 10, "bold")).pack(
            anchor="w"
        )
        ttk.Label(version_frame, text="Latest: 2025-12-15", font=("", 10)).pack(
            anchor="w", pady=(5, 0)
        )
        ttk.Label(
            version_frame, text="created by DIP Dpertment/A・T", font=("", 10)
        ).pack(anchor="w", pady=(5, 0))

        # 使い方
        usage_frame = ttk.LabelFrame(content_frame, text="使い方", padding=10)
        usage_frame.pack(fill="both", expand=True)

        usage_text = """
【基本設定】
1. シート名、測定No列、行範囲などを設定します
2. デフォルト値が設定されているので、必要に応じて変更してください

【工具と測定No対応】
1. 「工具追加」ボタンで工具を追加します
2. 工具名と対応する測定Noをカンマ区切りで入力します
  例: 1, 5, 10, 15
3. 「選択編集」「選択削除」で既存の工具を編集・削除できます

【測定不要書き込み設定】
1. B列に「測定不要」を書き込む行番号を指定します
2. I列に「-」を入れる測定Noをカンマ区切りで指定します
3. この設定は任意です（空欄の場合はスキップされます）

【Excel生成】
1. 「この設定でExcel生成」ボタンをクリックします
2. 元のExcelファイルを選択します
3. 出力先のファイル名を指定します
4. 処理が完了すると、指定した場所にExcelファイルが生成されます

【生成される内容】
- 指定した工具に対応する測定Noの行に「依頼」という文字が自動入力されます
- I列からSN列までの各列に対して処理が実行されます
- 測定不要設定が指定されている場合は、該当する行に「測定不要」や「-」が書き込まれます
        """

        usage_label = ttk.Label(
            usage_frame,
            text=usage_text.strip(),
            justify="left",
            font=("", 9),
        )
        usage_label.pack(anchor="w", padx=5)

        # 注意事項
        note_frame = ttk.LabelFrame(content_frame, text="注意事項", padding=10)
        note_frame.pack(fill="x", pady=(15, 0))

        note_text = """
• 元のExcelファイルは開いていない状態で実行してください
• 出力先ファイルが既に存在する場合は上書きされます
• 測定Noは整数で指定してください
• シート名が存在しない場合はエラーになります
        """

        note_label = ttk.Label(
            note_frame,
            text=note_text.strip(),
            justify="left",
            font=("", 9),
            foreground="firebrick",
        )
        note_label.pack(anchor="w", padx=5)

        # 閉じるボタン
        btn_frame = ttk.Frame(content_frame)
        btn_frame.pack(fill="x", pady=(15, 0))
        ttk.Button(btn_frame, text="閉じる", command=help_window.destroy).pack(
            side="right"
        )

        # スクロールバーとキャンバスの配置
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # マウスホイールでスクロール（Windows用）
        def _on_mousewheel(event):
            if event.delta:
                # Windows
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            else:
                # Linux
                if event.num == 4:
                    canvas.yview_scroll(-1, "units")
                elif event.num == 5:
                    canvas.yview_scroll(1, "units")

        # Windows用
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        # Linux用
        canvas.bind_all("<Button-4>", _on_mousewheel)
        canvas.bind_all("<Button-5>", _on_mousewheel)

        # キャンバスのスクロール領域を更新
        def update_scroll_region(event=None):
            canvas.update_idletasks()
            canvas.configure(scrollregion=canvas.bbox("all"))

        def configure_canvas_width(event):
            canvas_width = event.width
            canvas.itemconfig(canvas.find_all()[0], width=canvas_width)

        scrollable_frame.bind("<Configure>", update_scroll_region)
        canvas.bind("<Configure>", configure_canvas_width)

        help_window.focus_set()


def main():
    app = ConfigEditor()
    app.mainloop()


if __name__ == "__main__":
    main()
