import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import ttkbootstrap as tb
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

from .excel_ops import (
    AUTO_DATA_MAX_ITEMS,
    LOCKED_BASIC_SETTINGS,
    NOT_REQUIRED_ROW_DEFAULT,
    _derive_auto_data_start_row,
    _derive_layout_rows,
    _normalize_measure_no_key,
    _parse_int_list,
    _try_extract_int,
    build_request_formulas,
    write_measurement_not_required,
)
from .ui_helpers import LoadingDialog, pick_save_path


class ConfigEditor(tb.Window):
    def __init__(self):
        super().__init__(themename="darkly")
        self.title("検査シート 設定作成")
        self.geometry("820x640")
        try:
            self.state("zoomed")
        except Exception:
            pass

        measure_row_min_default = LOCKED_BASIC_SETTINGS["measure_row_min"]
        measure_row_step_default = LOCKED_BASIC_SETTINGS["measure_row_step"]
        not_required_row_default = NOT_REQUIRED_ROW_DEFAULT
        measure_row_max_default, tool_start_default = _derive_layout_rows(
            not_required_row_default,
            measure_row_min_default,
        )
        summary_row_min_default = LOCKED_BASIC_SETTINGS["summary_row_min"]
        summary_row_max_default = measure_row_max_default
        summary_row_step_default = LOCKED_BASIC_SETTINGS["summary_row_step"]
        tool_row_step_default = LOCKED_BASIC_SETTINGS["tool_row_step"]

        self.vars = {
            "sheet_name": tk.StringVar(value="工程内検査シート"),
            "measure_no_col": tk.StringVar(value=LOCKED_BASIC_SETTINGS["measure_no_col"]),
            "measure_row_min": tk.IntVar(value=measure_row_min_default),
            "measure_row_max": tk.IntVar(value=measure_row_max_default),
            "measure_row_step": tk.IntVar(value=measure_row_step_default),
            "summary_row_min": tk.IntVar(value=summary_row_min_default),
            "summary_row_max": tk.IntVar(value=summary_row_max_default),
            "summary_row_step": tk.IntVar(value=summary_row_step_default),
            "formula_arg_sep": tk.StringVar(value=LOCKED_BASIC_SETTINGS["formula_arg_sep"]),
            "tool_start_row": tk.IntVar(value=tool_start_default),
            "tool_name_col": tk.StringVar(value=LOCKED_BASIC_SETTINGS["tool_name_col"]),
            "tool_row_step": tk.IntVar(value=tool_row_step_default),
            "not_required_row": tk.StringVar(value=str(not_required_row_default)),
            "not_required_nos": tk.StringVar(value=""),
        }

        self.auto_map_measure_no_var = tk.StringVar(value="")
        self.auto_map_data_index_var = tk.StringVar(value="")

        self._bind_basic_setting_sync()
        self._apply_locked_basic_settings()

        self.selected_xlsx = tk.StringVar(value="")
        self.preview_title = tk.StringVar(value="プレビュー (未読み込み)")

        self._build_ui()

    def _build_ui(self):
        header_frame = ttk.Frame(self)
        header_frame.pack(fill="x", padx=10, pady=(10, 0))
        ttk.Button(header_frame, text="ヘルプ", command=self._show_help).pack(side="right")

        body = ttk.Frame(self)
        body.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self.main_canvas = tk.Canvas(body, highlightthickness=0, bd=0)
        self.main_scrollbar = ttk.Scrollbar(
            body,
            orient="vertical",
            command=self.main_canvas.yview,
        )
        self.main_canvas.configure(yscrollcommand=self.main_scrollbar.set)

        self.main_canvas.pack(side="left", fill="both", expand=True)
        self.main_scrollbar.pack(side="right", fill="y")

        main = ttk.Frame(self.main_canvas, padding=10)
        self.main_canvas_window = self.main_canvas.create_window(
            (0, 0),
            window=main,
            anchor="nw",
        )

        def _update_main_scroll_region(event=None):
            self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))

        def _fit_main_width(event):
            self.main_canvas.itemconfigure(self.main_canvas_window, width=event.width)

        main.bind("<Configure>", _update_main_scroll_region)
        self.main_canvas.bind("<Configure>", _fit_main_width)
        self.bind_all("<MouseWheel>", self._on_main_mousewheel, add="+")
        self.bind_all("<Button-4>", self._on_main_mousewheel, add="+")
        self.bind_all("<Button-5>", self._on_main_mousewheel, add="+")

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

        style = ttk.Style(self)
        style.configure(
            "Preview.Treeview",
            rowheight=22,
            borderwidth=1,
            relief="solid",
            bordercolor="#ffffff",
            lightcolor="#ffffff",
            darkcolor="#ffffff",
        )
        style.configure(
            "Preview.Treeview.Heading",
            borderwidth=1,
            relief="solid",
            anchor="center",
        )

        self.preview_tree = ttk.Treeview(
            preview_container,
            columns=self.preview_columns,
            show="headings",
            height=8,
            style="Preview.Treeview",
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

        def add_field(parent, row, label, key, width=12, editable=True):
            ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=(0, 8), pady=3)
            if editable:
                ttk.Entry(parent, textvariable=self.vars[key], width=width).grid(row=row, column=1, sticky="w", padx=(0, 20), pady=3)
                return
            ttk.Label(
                parent,
                text=str(self.vars[key].get()),
                width=width,
            ).grid(row=row, column=1, sticky="w", padx=(0, 20), pady=3)

        add_field(basic_left, 0, "シート名", "sheet_name", width=25)
        ttk.Label(basic_right, text="測定不要書き込み設定の行:").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=5)
        ttk.Entry(
            basic_right,
            textvariable=self.vars["not_required_row"],
            width=15,
        ).grid(row=0, column=1, sticky="w", padx=(0, 30), pady=5)
        ttk.Label(
            basic_right,
            text="L～SR列に'-'を入れるNo.(カンマ区切り):",
        ).grid(row=1, column=0, sticky="w", padx=(0, 10), pady=5)
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

        tools_btns = ttk.Frame(main)
        tools_btns.pack(fill="x", pady=(6, 0))
        ttk.Button(tools_btns, text="工具追加", command=self._add_tool_dialog).pack(side="left")
        ttk.Button(tools_btns, text="選択編集", command=self._edit_selected_tool).pack(side="left", padx=5)
        ttk.Button(tools_btns, text="選択削除", command=self._delete_selected_tool).pack(side="left")

        auto_map_frame = ttk.LabelFrame(main, text="自動測定データ対応（測定No → データ順番）", padding=10)
        auto_map_frame.pack(fill="both", expand=True, pady=(10, 0))

        auto_input = ttk.Frame(auto_map_frame)
        auto_input.pack(fill="x", pady=(0, 8))
        ttk.Label(auto_input, text="測定No").pack(side="left")
        ttk.Entry(auto_input, textvariable=self.auto_map_measure_no_var, width=12).pack(
            side="left", padx=(6, 12)
        )
        ttk.Label(auto_input, text=f"データ順番 (1〜{AUTO_DATA_MAX_ITEMS})").pack(side="left")
        ttk.Entry(auto_input, textvariable=self.auto_map_data_index_var, width=12).pack(
            side="left", padx=(6, 12)
        )
        ttk.Button(auto_input, text="追加", command=self._add_auto_map).pack(side="left")
        ttk.Button(auto_input, text="選択削除", command=self._delete_selected_auto_map).pack(
            side="left", padx=6
        )

        self.auto_map_tree = ttk.Treeview(
            auto_map_frame,
            columns=("measure_no", "data_index"),
            show="headings",
            height=6,
        )
        self.auto_map_tree.heading("measure_no", text="測定No")
        self.auto_map_tree.heading("data_index", text="データ順番")
        self.auto_map_tree.column("measure_no", width=200, anchor="w")
        self.auto_map_tree.column("data_index", width=140, anchor="center")
        self.auto_map_tree.pack(side="left", fill="both", expand=True)
        auto_scrollbar = ttk.Scrollbar(
            auto_map_frame,
            orient="vertical",
            command=self.auto_map_tree.yview,
        )
        self.auto_map_tree.configure(yscrollcommand=auto_scrollbar.set)
        auto_scrollbar.pack(side="right", fill="y")

        btns = ttk.Frame(main)
        btns.pack(fill="x", pady=(10, 0))
        ttk.Button(btns, text="この設定でExcel生成", command=self._run_build).pack(side="right")

        if not self.tools_tree.get_children():
            self._insert_tool("前挽き(サンプル)", "1, 5, 10")

    def _is_in_main_content(self, widget):
        current = widget
        while current is not None:
            if current is self.main_canvas:
                return True
            parent_name = current.winfo_parent()
            if not parent_name:
                break
            try:
                current = current.nametowidget(parent_name)
            except Exception:
                break
        return False

    def _on_main_mousewheel(self, event):
        if not self._is_in_main_content(event.widget):
            return
        if getattr(event, "delta", 0):
            self.main_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        elif getattr(event, "num", None) == 4:
            self.main_canvas.yview_scroll(-1, "units")
        elif getattr(event, "num", None) == 5:
            self.main_canvas.yview_scroll(1, "units")

    def _bind_basic_setting_sync(self):
        self.vars["not_required_row"].trace_add("write", self._sync_not_required_row)
        self.vars["measure_row_min"].trace_add("write", self._sync_measure_row_min)
        self.vars["measure_row_max"].trace_add("write", self._sync_measure_row_max)
        self.vars["measure_row_step"].trace_add("write", self._sync_measure_row_step)

    def _apply_locked_basic_settings(self):
        for key, value in LOCKED_BASIC_SETTINGS.items():
            if self.vars[key].get() != value:
                self.vars[key].set(value)

    def _sync_not_required_row(self, *args):
        try:
            not_required_row = int(self.vars["not_required_row"].get())
        except (tk.TclError, ValueError):
            return
        min_row = self.vars["measure_row_min"].get()
        desired_max, desired_tool_start = _derive_layout_rows(not_required_row, min_row)
        if self.vars["measure_row_max"].get() != desired_max:
            self.vars["measure_row_max"].set(desired_max)
        if self.vars["tool_start_row"].get() != desired_tool_start:
            self.vars["tool_start_row"].set(desired_tool_start)

    def _sync_measure_row_min(self, *args):
        try:
            min_row = self.vars["measure_row_min"].get()
        except tk.TclError:
            return
        self.vars["summary_row_min"].set(min_row)
        if self.vars["measure_row_max"].get() < min_row:
            self.vars["measure_row_max"].set(min_row)

    def _sync_measure_row_max(self, *args):
        try:
            max_row = self.vars["measure_row_max"].get()
        except tk.TclError:
            return
        self.vars["summary_row_max"].set(max_row)

    def _sync_measure_row_step(self, *args):
        try:
            step = self.vars["measure_row_step"].get()
        except tk.TclError:
            return
        if step < 1:
            self.vars["measure_row_step"].set(1)
            return
        self.vars["summary_row_step"].set(step)
        if self.vars["tool_row_step"].get() != step:
            self.vars["tool_row_step"].set(step)

    def _load_preview(self):
        path = filedialog.askopenfilename(
            parent=self,
            title="元の検査シート（xlsx）を選択",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not path:
            return
        self.selected_xlsx.set(path)

        loading = LoadingDialog(self, "読み込み中...", "Excelファイルを読み込んでいます...")
        result = {"success": False, "error": None}

        def load_task():
            try:
                self._render_preview_internal()
                result["success"] = True
            except Exception as e:
                result["error"] = e

        thread = threading.Thread(target=load_task, daemon=True)
        thread.start()

        def check_completion():
            if thread.is_alive():
                self.after(100, check_completion)
                return
            loading.close()
            if result["error"]:
                messagebox.showerror("読み込み失敗", str(result["error"]), parent=self)

        self.after(100, check_completion)

    def _render_preview(self):
        path = self.selected_xlsx.get().strip()
        if not path:
            messagebox.showinfo("読み込み待ち", "先にExcelファイルを選択してください。", parent=self)
            return

        loading = LoadingDialog(self, "更新中...", "プレビューを更新しています...")
        result = {"success": False, "error": None}

        def render_task():
            try:
                self._render_preview_internal()
                result["success"] = True
            except Exception as e:
                result["error"] = e

        thread = threading.Thread(target=render_task, daemon=True)
        thread.start()

        def check_completion():
            if thread.is_alive():
                self.after(100, check_completion)
                return
            loading.close()
            if result["error"]:
                messagebox.showerror("更新失敗", str(result["error"]), parent=self)

        self.after(100, check_completion)

    def _render_preview_internal(self):
        path = self.selected_xlsx.get().strip()
        if not path:
            return

        sheet_name = self.vars["sheet_name"].get().strip() or "工程内検査シート"
        try:
            wb = load_workbook(path, data_only=True)
        except Exception as e:
            raise Exception(f"Excelを開けませんでした。\n{e}")

        if sheet_name not in wb.sheetnames:
            raise Exception(
                f"シート「{sheet_name}」が見つかりません。\nシート名を確認して再度プレビューしてください。"
            )

        ws = wb[sheet_name]
        preview_col_indices = [column_index_from_string(column) for column in self.preview_columns]

        if not ws.max_row or not ws.max_column:
            raise Exception("表示できるデータがありません。")

        col_widths = {"A": 80, "B": 120, "G": 140, "K": 140}

        def update_ui():
            self.preview_tree.configure(columns=self.preview_columns)
            for col in self.preview_columns:
                self.preview_tree.heading(col, text=col, anchor="center")
                self.preview_tree.column(
                    col,
                    width=col_widths.get(col, 120),
                    minwidth=60,
                    anchor="center",
                    stretch=False,
                )

            for item in self.preview_tree.get_children():
                self.preview_tree.delete(item)

            header_row = 10
            group_size = 3
            max_row_value = ws.max_row or 0

            if max_row_value < header_row:
                raise Exception(f"{header_row}行目以降に表示できるデータがありません。")

            header_vals = []
            for col_idx in preview_col_indices:
                value = ws.cell(header_row, col_idx).value
                header_vals.append("" if value is None else str(value))
            self.preview_tree.insert("", "end", values=header_vals)

            data_start_row = header_row + 1
            row_count = 0
            for row_index in range(data_start_row, max_row_value + 1, group_size):
                a_val = ws.cell(row_index, preview_col_indices[0]).value
                if a_val is None or (isinstance(a_val, str) and not a_val.strip()):
                    break

                row_vals = []
                for col_idx in preview_col_indices:
                    value = ws.cell(row_index, col_idx).value
                    row_vals.append("" if value is None else str(value))

                self.preview_tree.insert("", "end", values=row_vals)
                row_count += 1

            if row_count == 0:
                for item in self.preview_tree.get_children():
                    self.preview_tree.delete(item)
                raise Exception("11行目以降の先頭行(A列)が空のため表示できるデータがありません。")

            self.preview_title.set(
                f"{sheet_name} プレビュー（{row_count}行: 10行目ヘッダー＋11行目以降3行刻み先頭行）"
            )

        self.after(0, update_ui)

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

        ttk.Label(frm, text="測定No(カンマ区切り)").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(frm, textvariable=nos_var, width=40).grid(row=1, column=1, sticky="w", pady=4)

        result = {"ok": False}

        def on_ok():
            tool = tool_var.get().strip()
            if not tool:
                messagebox.showwarning("入力不足", "工具名を入力してください。", parent=win)
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
        ttk.Button(bfrm, text="キャンセル", command=on_cancel).pack(side="right", padx=5)

        tool_entry.focus_set()
        self.wait_window(win)
        return result

    def _add_tool_dialog(self):
        result = self._tool_dialog("工具追加")
        if result.get("ok"):
            self._insert_tool(result["tool"], result["nos"])

    def _edit_selected_tool(self):
        selected = self.tools_tree.selection()
        if not selected:
            messagebox.showinfo("選択なし", "編集する行を選択してください。", parent=self)
            return
        item = selected[0]
        tool, nos = self.tools_tree.item(item, "values")
        result = self._tool_dialog("工具編集", init_tool=tool, init_nos=nos)
        if result.get("ok"):
            self.tools_tree.item(item, values=(result["tool"], result["nos"]))

    def _delete_selected_tool(self):
        selected = self.tools_tree.selection()
        if not selected:
            return
        if not messagebox.askyesno("削除確認", "選択した工具を削除しますか？", parent=self):
            return
        for item in selected:
            self.tools_tree.delete(item)

    def _add_auto_map(self):
        measure_no_raw = self.auto_map_measure_no_var.get().strip()
        data_index_raw = self.auto_map_data_index_var.get().strip()

        key = _normalize_measure_no_key(measure_no_raw)
        data_index = _try_extract_int(data_index_raw)

        if key == "":
            messagebox.showwarning("入力エラー", "測定Noを入力してください。", parent=self)
            return
        if data_index is None or data_index < 1 or data_index > AUTO_DATA_MAX_ITEMS:
            messagebox.showwarning(
                "入力エラー",
                f"データ順番は1〜{AUTO_DATA_MAX_ITEMS}の整数で入力してください。",
                parent=self,
            )
            return

        key_str = str(key)
        existing = None
        for item in self.auto_map_tree.get_children():
            values = self.auto_map_tree.item(item, "values")
            if str(values[0]) == key_str:
                existing = item
                break

        if existing is not None:
            self.auto_map_tree.item(existing, values=(key_str, str(data_index)))
        else:
            self.auto_map_tree.insert("", "end", values=(key_str, str(data_index)))

        self.auto_map_measure_no_var.set("")
        self.auto_map_data_index_var.set("")

    def _delete_selected_auto_map(self):
        selected = self.auto_map_tree.selection()
        if not selected:
            return
        for item in selected:
            self.auto_map_tree.delete(item)

    def _gather_cfg(self):
        try:
            self._apply_locked_basic_settings()
            tools = []
            tool_to_measure_nos = {}
            measure_no_to_data_index = {}

            for item in self.tools_tree.get_children():
                tool, nos_text = self.tools_tree.item(item, "values")
                tools.append(tool)
                tool_to_measure_nos[tool] = _parse_int_list(nos_text)

            for item in self.auto_map_tree.get_children():
                measure_no_text, data_index_text = self.auto_map_tree.item(item, "values")
                key = _normalize_measure_no_key(measure_no_text)
                data_index = _try_extract_int(data_index_text)
                if key == "" or data_index is None:
                    continue
                measure_no_to_data_index[key] = data_index

            if not tools:
                raise ValueError("工具が1件もありません。")

            not_required_row_text = self.vars["not_required_row"].get().strip()
            not_required_row = _try_extract_int(not_required_row_text)
            if not_required_row is None:
                raise ValueError("測定不要書き込み設定の行は整数で入力してください。")

            measure_row_min = LOCKED_BASIC_SETTINGS["measure_row_min"]
            measure_row_max, tool_start_row = _derive_layout_rows(not_required_row, measure_row_min)
            if measure_row_max < measure_row_min:
                raise ValueError(
                    f"測定不要書き込み設定の行は {measure_row_min + 1} 以上で入力してください。"
                )
            if tool_start_row < 1:
                raise ValueError("自動計算後の工具開始行が1未満になります。入力値を見直してください。")
            auto_data_start_row = _derive_auto_data_start_row(not_required_row, len(tools))

            cfg = {
                "sheet_name": self.vars["sheet_name"].get().strip(),
                "measure_no_col": LOCKED_BASIC_SETTINGS["measure_no_col"],
                "measure_row_min": measure_row_min,
                "measure_row_max": measure_row_max,
                "measure_row_step": LOCKED_BASIC_SETTINGS["measure_row_step"],
                "summary_row_min": LOCKED_BASIC_SETTINGS["summary_row_min"],
                "summary_row_max": measure_row_max,
                "summary_row_step": LOCKED_BASIC_SETTINGS["summary_row_step"],
                "formula_arg_sep": LOCKED_BASIC_SETTINGS["formula_arg_sep"],
                "tool_start_row": tool_start_row,
                "not_required_row": not_required_row,
                "tool_name_col": LOCKED_BASIC_SETTINGS["tool_name_col"],
                "tool_row_step": LOCKED_BASIC_SETTINGS["tool_row_step"],
                "auto_data_start_row": auto_data_start_row,
                "measure_no_to_data_index": measure_no_to_data_index,
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

        loading = LoadingDialog(self, "生成中...", "Excelファイルを生成しています...")
        result = {"success": False, "error": None, "saved_path": None}

        def build_task():
            try:
                saved_path = build_request_formulas(xlsx, out_path, cfg, parent=self)
                result["saved_path"] = saved_path

                not_required_nos_text = self.vars["not_required_nos"].get().strip()
                if not_required_nos_text:
                    try:
                        target_nos = _parse_int_list(not_required_nos_text)
                        if target_nos:
                            saved_path = write_measurement_not_required(
                                saved_path,
                                out_path,
                                cfg,
                                target_nos=target_nos,
                                parent=self,
                            )
                            result["saved_path"] = saved_path
                    except Exception as e:
                        result["error"] = f"測定不要書き込みでエラーが発生しました:\n{e}"
                        result["success"] = True
                        return

                result["success"] = True
            except Exception as e:
                result["error"] = str(e)

        thread = threading.Thread(target=build_task, daemon=True)
        thread.start()

        def check_completion():
            if thread.is_alive():
                self.after(100, check_completion)
                return
            loading.close()
            if result["success"]:
                if result["error"]:
                    messagebox.showwarning(
                        "警告",
                        f"Excel生成は完了しましたが、{result['error']}",
                        parent=self,
                    )
                else:
                    messagebox.showinfo("完了", f"生成しました:\n{result['saved_path']}", parent=self)
                return
            messagebox.showerror("失敗", result["error"], parent=self)

        self.after(100, check_completion)

    def _run_write_not_required(self):
        try:
            cfg = self._gather_cfg()
        except Exception as e:
            messagebox.showerror("失敗", str(e), parent=self)
            return

        try:
            target_nos = _parse_int_list(self.vars["not_required_nos"].get())
        except Exception as e:
            messagebox.showerror("入力エラー", f"測定Noの入力が不正です。\n{e}", parent=self)
            return

        if not target_nos:
            messagebox.showwarning(
                "入力不足",
                "L～SR列に'-'を入れるNo.を入力してください。",
                parent=self,
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
                target_nos=target_nos,
                parent=self,
            )
        except Exception as e:
            messagebox.showerror("失敗", str(e), parent=self)
            return
        messagebox.showinfo("完了", f"書き込み完了:\n{saved_path}", parent=self)

    def _show_help(self):
        help_window = tk.Toplevel(self)
        help_window.title("使い方")
        help_window.transient(self)
        help_window.grab_set()
        help_window.geometry("600x500")

        canvas = tk.Canvas(help_window, highlightthickness=0, bd=0)
        scrollbar = ttk.Scrollbar(help_window, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        content_frame = ttk.Frame(scrollable_frame, padding=20)
        content_frame.pack(fill="both", expand=True)

        version_frame = ttk.LabelFrame(content_frame, text="バージョン情報", padding=10)
        version_frame.pack(fill="x", pady=(0, 15))

        ttk.Label(version_frame, text="Version: 0.7.0", font=("", 10, "bold")).pack(anchor="w")
        ttk.Label(version_frame, text="Latest: 2026-02-26", font=("", 10)).pack(anchor="w", pady=(5, 0))
        ttk.Label(version_frame, text="created by DIP Dpertment/A・T", font=("", 10)).pack(anchor="w", pady=(5, 0))

        usage_frame = ttk.LabelFrame(content_frame, text="使い方", padding=10)
        usage_frame.pack(fill="both", expand=True)

        usage_text = """
    【基本設定】
    1. 先に「Excelを選択」で元ファイルを読み込み、プレビューで対象シートを確認します
    2. 基本設定では「シート名」だけを必要に応じて調整します
    3. 測定不要書き込み設定の行を入力すると、測定行(max) はその1つ上、工具開始行はその3つ下として内部で自動計算します
    4. 自動測定データ開始行は「測定不要書き込み設定の行 + (工具数 × 3) + 6」で内部自動計算します
    5. 出力列は L～SR 固定です

    【先頭集計式の注意】
    1. 1行目の基準式は =SUMPRODUCT(--(L11:L308<>""),--(MOD(ROW(L11:L308)-ROW(L11),2)=0)) です
    2. 1行目がこの形式でない場合は、L〜SN列の1〜3行目に同パターンの式を補正投入します
    3. 2行目・3行目は1行目を貼り付けた相対参照と同じ内容で設定します

    【工具と測定No対応】
    1. 「工具追加」で工具名と測定No（カンマ区切り）を登録します
    2. 工具行の同じ列に値が入った場合、10行目も「依頼」になります
    3. 登録した測定Noの行に、列ごとに「依頼」判定式が設定されます
    4. 同じ測定Noが自動測定データ対応にもある場合は、自動測定結果参照を優先し、参照結果が空のときだけ「依頼」を表示します

    【注意】
    1. 生成したファイルを開くと「作成されたファイルを修正しますか？」と表示される場合があります
    2. その場合は「はい」を押して進めてください
    3. 書式やレイアウトが変わっている場合は、必要な関数のみを既存のExcelへ貼り付けて使用してください

    【自動測定結果の反映】
    1. 自動測定データ開始行は「測定不要書き込み設定の行 + (工具数 × 3) + 6」で自動計算します
    2. 「測定No → データ順番」を追加すると、対応する測定行のL～SRに自動測定結果参照式を設定します
    3. データ順番が1なら開始行、2なら開始行+1…を同じ列で参照します

    【測定不要書き込み設定（任意）】
    1. 「測定不要書き込み設定の行」には、E列へ「測定不要」を書き込む行番号を入力します
    2. No.欄に測定Noを入れると、生成後に追加処理として測定不要式を設定します
    3. 指定Noの行のL～SRでは、自動測定結果参照と「依頼」を優先し、どちらも空のときだけ同列の「測定不要」入力に応じて「-」を表示します

    【Excel生成の処理順】
    1. 「この設定でExcel生成」を押して出力先を指定します
    2. 依頼式・自動測定参照式を書き込みます
    3. 測定不要Noが指定されている場合は続けて測定不要式を書き込みます
    4. 最後に再計算・保存して完了します
        """

        usage_label = ttk.Label(
            usage_frame,
            text=usage_text.strip(),
            justify="left",
            font=("", 9),
        )
        usage_label.pack(anchor="w", padx=5)

        note_frame = ttk.LabelFrame(content_frame, text="注意事項", padding=10)
        note_frame.pack(fill="x", pady=(15, 0))

        note_text = """
    • 元Excel・出力先Excelは閉じた状態で実行してください（保存失敗の原因になります）
    • 測定Noは整数で指定してください（例: 1,5,10）
    • シート名や測定行範囲が実ファイルと異なると対象行を特定できません
    • 自動測定結果参照の式を使うため、元シートの該当列データ位置を事前に確認してください
        """

        note_label = ttk.Label(
            note_frame,
            text=note_text.strip(),
            justify="left",
            font=("", 9),
            foreground="firebrick",
        )
        note_label.pack(anchor="w", padx=5)

        btn_frame = ttk.Frame(content_frame)
        btn_frame.pack(fill="x", pady=(15, 0))
        ttk.Button(btn_frame, text="閉じる", command=help_window.destroy).pack(side="right")

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def _on_mousewheel(event):
            if event.delta:
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
                return
            if event.num == 4:
                canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                canvas.yview_scroll(1, "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        canvas.bind_all("<Button-4>", _on_mousewheel)
        canvas.bind_all("<Button-5>", _on_mousewheel)

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
