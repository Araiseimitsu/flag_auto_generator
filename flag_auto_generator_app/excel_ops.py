import os
import re
import time
import zipfile

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.formula import ArrayFormula
from tkinter import messagebox


AUTO_DATA_START_ROW_DEFAULT = 230
NOT_REQUIRED_ROW_DEFAULT = 197
AUTO_DATA_MAX_ITEMS = 100
REQUEST_HEADER_ROW = 10
SUMMARY_FORMULA_COL_START = "L"
SUMMARY_FORMULA_COL_END = "SN"
REQUEST_OUTPUT_COL_START = "L"
REQUEST_OUTPUT_COL_END = "SR"
SUMMARY_FORMULA_BASE_START_ROW = 9
SUMMARY_FORMULA_BASE_END_ROW = 308
LOCKED_BASIC_SETTINGS = {
    "measure_no_col": "A",
    "measure_row_min": 11,
    "measure_row_step": 3,
    "summary_row_min": 11,
    "summary_row_step": 3,
    "formula_arg_sep": ",",
    "tool_name_col": "E",
    "tool_row_step": 3,
}


def _derive_layout_rows(not_required_row: int, measure_row_min: int) -> tuple[int, int]:
    measure_row_max = max(not_required_row - 1, measure_row_min)
    tool_start_row = max(not_required_row + 3, 1)
    return measure_row_max, tool_start_row


def _derive_auto_data_start_row(not_required_row: int, tool_count: int) -> int:
    return not_required_row + (tool_count * LOCKED_BASIC_SETTINGS["tool_row_step"]) + 6


def _get_writable_cell(ws, row: int, col: int):
    cell = ws.cell(row, col)
    if not isinstance(cell, MergedCell):
        return cell

    for merged_range in ws.merged_cells.ranges:
        if merged_range.min_row <= row <= merged_range.max_row and merged_range.min_col <= col <= merged_range.max_col:
            return ws.cell(merged_range.min_row, merged_range.min_col)

    return cell


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


def _mark_workbook_for_full_recalc(wb):
    try:
        wb.calculation.calcMode = "auto"
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
    except Exception:
        pass


def _safe_call(func, *args, default=None, **kwargs):
    try:
        return func(*args, **kwargs)
    except Exception:
        return default


def _close_workbook_quietly(wb):
    if wb is None:
        return
    _safe_call(wb.close)


def _restore_package_parts_from_source(source_xlsx_path: str, target_xlsx_path: str):
    prefixes = (
        "xl/drawings/",
        "xl/externalLinks/",
    )

    source_path = os.path.abspath(source_xlsx_path)
    target_path = os.path.abspath(target_xlsx_path)
    temp_path = f"{target_path}.pkgfix"

    try:
        with zipfile.ZipFile(source_path, "r") as src_zip, zipfile.ZipFile(
            target_path, "r"
        ) as dst_zip:
            src_names = set(src_zip.namelist())
            dst_names = set(dst_zip.namelist())
            restore_names = {
                name
                for name in src_names
                if any(name.startswith(prefix) for prefix in prefixes)
            }
            if not restore_names:
                return False

            with zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as out_zip:
                for name in dst_zip.namelist():
                    if any(name.startswith(prefix) for prefix in prefixes):
                        if name in restore_names:
                            out_zip.writestr(name, src_zip.read(name))
                        continue
                    out_zip.writestr(name, dst_zip.read(name))

                for name in sorted(restore_names):
                    if name in dst_names:
                        continue
                    out_zip.writestr(name, src_zip.read(name))

        os.replace(temp_path, target_path)
        return True
    except Exception as e:
        _safe_call(os.remove, temp_path)
        print(f"[info] パッケージ復元をスキップしました: {e}")
        return False


def _force_excel_recalc_and_save(xlsx_path: str):
    try:
        import win32com.client  # type: ignore
    except Exception as e:
        print(f"[warn] Excel強制再計算をスキップしました（pywin32未導入）: {e}")
        return False

    abs_path = os.path.abspath(xlsx_path)
    last_error = None

    for attempt in range(3):
        excel_app = None
        excel_wb = None
        try:
            excel_app = win32com.client.DispatchEx("Excel.Application")
            excel_app.Visible = False
            excel_app.DisplayAlerts = False
            _safe_call(setattr, excel_app, "AskToUpdateLinks", False)

            excel_wb = excel_app.Workbooks.Open(abs_path, UpdateLinks=0, ReadOnly=False)

            _safe_call(setattr, excel_wb, "ForceFullCalculation", True)

            worksheets = _safe_call(lambda: list(excel_wb.Worksheets), default=[])
            for sheet in worksheets:
                _safe_call(setattr, sheet, "EnableCalculation", True)
                _safe_call(sheet.UsedRange.Calculate)

            _safe_call(excel_wb.RefreshAll)
            _safe_call(excel_app.CalculateFull)
            excel_app.CalculateFullRebuild()
            save_as_kwargs = {
                "Filename": abs_path,
                "FileFormat": 51,
                "ConflictResolution": 2,
                "Local": True,
            }
            if _safe_call(excel_wb.SaveAs, **save_as_kwargs) is None:
                excel_wb.Save()
            return True
        except Exception as e:
            last_error = e
            time.sleep(0.4 * (attempt + 1))
        finally:
            if excel_wb is not None:
                _safe_call(excel_wb.Close, SaveChanges=True)
            if excel_app is not None:
                _safe_call(excel_app.Quit)

    print(f"[info] Excel強制再計算をスキップしました（出力ファイルは作成済み）: {last_error}")
    return False


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


def _resolve_measure_no(value, row: int, measure_row_min: int, measure_row_step: int):
    if isinstance(value, str) and value.startswith("="):
        if measure_row_step <= 0:
            return None

        row_offset = row - measure_row_min
        if row_offset < 0:
            return None
        if row_offset % measure_row_step != 0:
            return None

        return (row_offset // measure_row_step) + 1

    return _try_extract_int(value)


def _normalize_measure_no_key(value):
    n = _try_extract_int(value)
    if n is not None:
        return n
    if value is None:
        return ""
    return str(value).strip()


def _normalize_measure_to_index_map(raw_map: dict):
    normalized = {}
    if not isinstance(raw_map, dict):
        return normalized
    for measure_no, data_index in raw_map.items():
        key = _normalize_measure_no_key(measure_no)
        if key == "":
            continue
        index_value = _try_extract_int(data_index)
        if index_value is None or index_value < 1 or index_value > AUTO_DATA_MAX_ITEMS:
            continue
        normalized[key] = index_value
    return normalized


def _build_auto_data_formula(col_letter: str, data_start_row: int, data_index: int, sep: str):
    source_row = data_start_row + data_index - 1
    return (
        f'IFERROR(IF({col_letter}{source_row}=""{sep}""{sep}{col_letter}{source_row}){sep}"")'
    )


def _is_empty_cell_value(value):
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def _to_formula_expression(value):
    if _is_empty_cell_value(value):
        return '""'
    if isinstance(value, str) and value.startswith("="):
        return value[1:]
    if isinstance(value, str):
        escaped = value.replace('"', '""')
        return f'"{escaped}"'
    return str(value)


def _build_not_required_overlay_formula(
    current_value,
    trigger_cell_ref: str,
    sep: str,
):
    base_expression = _to_formula_expression(current_value)
    return (
        f'=IF(IFERROR(LEN(TRIM(({base_expression})&"")){sep}0)>0{sep}'
        f'({base_expression}){sep}'
        f'IF(IFERROR(LEN(TRIM({trigger_cell_ref}&"")){sep}0)>0{sep}"-"{sep}""))'
    )


def _can_overwrite_with_formula(value):
    if _is_empty_cell_value(value):
        return True
    if isinstance(value, str) and value.startswith("="):
        return True
    return False


def _normalize_single_cell_array_formulas_in_column(
    ws,
    col_letter: str,
    *,
    row_start: int = 1,
    row_end: int | None = None,
):
    col_idx = column_index_from_string(col_letter)
    max_row = ws.max_row or row_start
    end_row = max_row if row_end is None else min(row_end, max_row)
    rewritten = 0

    start_row = max(row_start, 1)
    for row in range(start_row, end_row + 1):
        cell = ws.cell(row, col_idx)
        value = cell.value
        if not isinstance(value, ArrayFormula):
            continue

        formula_ref = str(value.ref or "").replace("$", "")
        cell_ref = cell.coordinate.replace("$", "")
        if formula_ref != cell_ref:
            continue

        formula_text = str(value.text or "").strip()
        if not formula_text:
            continue

        cell.value = f"={formula_text.lstrip('=')}"
        rewritten += 1

    return rewritten


def _build_summary_formula(col_letter: str, row_offset: int) -> str:
    start_row = SUMMARY_FORMULA_BASE_START_ROW + row_offset
    end_row = SUMMARY_FORMULA_BASE_END_ROW + row_offset
    return (
        f"=SUMPRODUCT(--({col_letter}{start_row}:{col_letter}{end_row}<>\"\"),"
        f"--(MOD(ROW({col_letter}{start_row}:{col_letter}{end_row})-ROW({col_letter}{start_row}),2)=0))"
    )


def _build_stepped_non_empty_count_formula(
    col_letter: str,
    row_min: int,
    row_max: int,
    row_step: int,
    sep: str,
) -> str:
    target_range = f"{col_letter}{row_min}:{col_letter}{row_max}"
    start_ref = f"{col_letter}{row_min}"
    return (
        f'SUMPRODUCT(--(LEN(TRIM({target_range}&""))>0){sep}'
        f'--(MOD(ROW({target_range})-ROW({start_ref}){sep}{row_step})=0))'
    )


def _build_request_formula(conditions: str, fallback_expression: str = '""') -> str:
    return f'=IF(OR({conditions}),"依頼",{fallback_expression})'


def _build_measure_row_formula(
    request_conditions: str,
    sep: str,
    auto_formula: str | None = None,
) -> str:
    if auto_formula is not None:
        auto_expression = _strip_formula_prefix(auto_formula)
        return (
            f'=IF(LEN(TRIM(({auto_expression})&""))>0{sep}'
            f'({auto_expression}){sep}'
            f'IF(OR({request_conditions}){sep}"依頼"{sep}""))'
        )
    return _build_request_formula(request_conditions)


def _build_request_header_formula(
    col_letter: str,
    sep: str,
    tool_row_min: int | None = None,
    tool_row_max: int | None = None,
    tool_row_step: int | None = None,
) -> str:
    if tool_row_min is None or tool_row_max is None or not tool_row_step:
        return ""
    conditions = (
        f"{_build_stepped_non_empty_count_formula(col_letter, tool_row_min, tool_row_max, tool_row_step, sep)}>0"
    )
    return _build_request_formula(conditions)


def _strip_formula_prefix(formula_text: str) -> str:
    if isinstance(formula_text, str) and formula_text.startswith("="):
        return formula_text[1:]
    return str(formula_text)


def _normalize_formula_text(value) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", "", str(value)).replace("$", "").upper()


def _ensure_summary_formulas(ws):
    summary_col_start = column_index_from_string(SUMMARY_FORMULA_COL_START)
    summary_col_end = column_index_from_string(SUMMARY_FORMULA_COL_END)

    for col_idx in range(summary_col_start, summary_col_end + 1):
        col_letter = get_column_letter(col_idx)
        expected_row1_formula = _build_summary_formula(col_letter, 0)
        row1_cell = ws.cell(1, col_idx)
        if _normalize_formula_text(row1_cell.value) == _normalize_formula_text(
            expected_row1_formula
        ):
            continue

        for row_idx in range(1, 4):
            ws.cell(row_idx, col_idx).value = _build_summary_formula(
                col_letter,
                row_idx - 1,
            )


def build_request_formulas(xlsx_path: str, out_path: str, cfg: dict, *, parent=None):
    sheet_name = cfg["sheet_name"]
    measure_no_col = cfg.get("measure_no_col", "A")
    measure_row_min = int(cfg.get("measure_row_min", 11))
    measure_row_step = int(cfg.get("measure_row_step", 3))
    tool_start_row = int(cfg.get("tool_start_row", 200))
    measure_row_max = int(cfg.get("measure_row_max", tool_start_row - 4))
    formula_arg_sep = str(cfg.get("formula_arg_sep", ",")).strip() or ","
    flag_col_start = column_index_from_string(REQUEST_OUTPUT_COL_START)
    flag_col_end = column_index_from_string(REQUEST_OUTPUT_COL_END)

    tool_name_col = cfg.get("tool_name_col", "E")
    tool_row_step = int(cfg.get("tool_row_step", measure_row_step))
    auto_data_start_row = int(cfg.get("auto_data_start_row", AUTO_DATA_START_ROW_DEFAULT))
    measure_no_to_data_index = _normalize_measure_to_index_map(
        cfg.get("measure_no_to_data_index", {})
    )

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

    no_col = column_index_from_string(measure_no_col)
    measure_no_to_row = {}
    measure_row_to_no = {}
    max_r = min(measure_row_max, ws.max_row or measure_row_max)
    for row_index in range(measure_row_min, max_r + 1):
        value = ws_values.cell(row_index, no_col).value
        if value is None:
            value = ws.cell(row_index, no_col).value
        if value is None:
            continue
        measure_no = _resolve_measure_no(value, row_index, measure_row_min, measure_row_step)
        if measure_no is None:
            continue
        measure_no_to_row[measure_no] = row_index
        measure_row_to_no[row_index] = measure_no

    tool_row = {}
    tool_name_col_index = column_index_from_string(tool_name_col)
    current_tool_row = tool_start_row
    for tool in tools:
        tool_cell = _get_writable_cell(ws, current_tool_row, tool_name_col_index)
        tool_cell.value = tool
        tool_row[tool] = tool_cell.row
        current_tool_row += tool_row_step

    auto_data_label_cell = _get_writable_cell(ws, auto_data_start_row, tool_name_col_index)
    auto_data_label_cell.value = f"測定結果貼付は{auto_data_start_row}行から"

    measure_row_to_tool_rows = {}
    missing_nos = []
    for tool, nos in tool_to_measure_nos.items():
        if tool not in tool_row:
            continue
        tool_row_index = tool_row[tool]
        for no in nos:
            measure_no = _try_extract_int(no)
            if measure_no is None:
                continue
            measure_row = measure_no_to_row.get(measure_no)
            if measure_row is None:
                missing_nos.append((tool, measure_no))
                continue
            measure_row_to_tool_rows.setdefault(measure_row, []).append(tool_row_index)

    all_tool_rows = sorted(tool_row.values())

    _ensure_summary_formulas(ws)

    written = 0
    target_found = 0
    for col_idx in range(flag_col_start, flag_col_end + 1):
        col_letter = get_column_letter(col_idx)
        header_cell = ws.cell(REQUEST_HEADER_ROW, col_idx)
        if _can_overwrite_with_formula(header_cell.value):
            header_formula = _build_request_header_formula(
                col_letter=col_letter,
                sep=formula_arg_sep,
                tool_row_min=all_tool_rows[0] if all_tool_rows else None,
                tool_row_max=all_tool_rows[-1] if all_tool_rows else None,
                tool_row_step=tool_row_step if all_tool_rows else None,
            )
            header_cell.value = header_formula

        for measure_row, tool_rows in measure_row_to_tool_rows.items():
            conditions = formula_arg_sep.join([f'{col_letter}${tool_row_index}<>""' for tool_row_index in tool_rows])
            measure_no = measure_row_to_no.get(measure_row)
            data_index = measure_no_to_data_index.get(measure_no)
            target_found += 1
            target_cell = ws.cell(measure_row, col_idx)
            if not _can_overwrite_with_formula(target_cell.value):
                continue
            if data_index is not None:
                auto_formula = _build_auto_data_formula(
                    col_letter=col_letter,
                    data_start_row=auto_data_start_row,
                    data_index=data_index,
                    sep=formula_arg_sep,
                )
                target_cell.value = _build_measure_row_formula(
                    conditions,
                    formula_arg_sep,
                    auto_formula,
                )
            else:
                target_cell.value = _build_measure_row_formula(conditions, formula_arg_sep)
            written += 1

        for measure_no, data_index in measure_no_to_data_index.items():
            measure_row = measure_no_to_row.get(measure_no)
            if measure_row is None or measure_row in measure_row_to_tool_rows:
                continue
            target_found += 1
            target_cell = ws.cell(measure_row, col_idx)
            if not _can_overwrite_with_formula(target_cell.value):
                continue
            auto_formula = _build_auto_data_formula(
                col_letter=col_letter,
                data_start_row=auto_data_start_row,
                data_index=data_index,
                sep=formula_arg_sep,
            )
            target_cell.value = f"={auto_formula}"
            written += 1

    if target_found == 0:
        raise ValueError(
            "依頼/自動測定データの書き込み対象が1件も見つかりませんでした。\n\n"
            f"- 読み取れた測定No件数: {len(measure_no_to_row)}\n"
            f"- 工具件数: {len(tools)}\n"
            f"- 逆引き対象の測定行件数: {len(measure_row_to_tool_rows)}\n"
            "- 出力列: L～SR（固定）\n"
            f"- 指定Noが見つからない例(先頭10件): {missing_nos[:10]}\n\n"
            "原因候補:\n"
            "1) 測定No列/行範囲が実シートと違う\n"
            "2) 測定Noが数式で、キャッシュ値が未保存（Excelで一度保存してから再実行）\n"
            "3) tool_to_measure_nos のNoがシートに存在しない"
        )

    normalized_count = _normalize_single_cell_array_formulas_in_column(
        ws,
        "B",
        row_start=measure_row_min,
        row_end=measure_row_max,
    )
    if normalized_count:
        print(f"[info] B列の単一セル配列数式を通常数式へ変換: {normalized_count}件")

    _mark_workbook_for_full_recalc(wb)
    saved_path = _save_workbook_atomic(wb, out_path, parent=parent)
    _close_workbook_quietly(wb_values)
    _close_workbook_quietly(wb)
    _restore_package_parts_from_source(xlsx_path, saved_path)
    _force_excel_recalc_and_save(saved_path)
    return saved_path


def _parse_int_list(text: str):
    if not text or not text.strip():
        return []
    parts = re.split(r"[,\s、，;；]+", text.strip())
    nums = []
    for part in parts:
        if not part:
            continue
        value = _try_extract_int(part)
        if value is None:
            raise ValueError(f"測定Noに整数以外の入力が含まれています: '{part}'")
        nums.append(value)
    return nums


def write_measurement_not_required(
    xlsx_path: str,
    out_path: str,
    cfg: dict,
    target_nos: list | None = None,
    *,
    parent=None,
):
    if target_nos is None:
        target_nos = []

    sheet_name = cfg.get("sheet_name", "工程内検査シート")
    measure_no_col = cfg.get("measure_no_col", "A")
    tool_start_row = int(cfg.get("tool_start_row", 200))
    not_required_row = int(cfg.get("not_required_row", tool_start_row - 3))
    measure_row_min = int(cfg.get("measure_row_min", 11))
    measure_row_step = int(cfg.get("measure_row_step", 3))
    measure_row_max = int(cfg.get("measure_row_max", tool_start_row - 4))
    formula_arg_sep = str(cfg.get("formula_arg_sep", ",")).strip() or ","

    wb = load_workbook(xlsx_path)
    wb_values = load_workbook(xlsx_path, data_only=True)

    if sheet_name not in wb.sheetnames:
        raise ValueError(
            f"シート '{sheet_name}' が見つかりません。存在: {wb.sheetnames}"
        )

    ws = wb[sheet_name]
    ws_values = wb_values[sheet_name]

    no_col = column_index_from_string(measure_no_col)
    measure_no_to_row = {}
    max_r = min(measure_row_max, ws.max_row or measure_row_max)
    debug_info = []

    for row_index in range(measure_row_min, max_r + 1):
        value = ws_values.cell(row_index, no_col).value
        cell_obj = ws.cell(row_index, no_col)
        formula = None

        if value is None:
            value = cell_obj.value
            if isinstance(value, str) and value.startswith("="):
                formula = value

        if value is None:
            continue

        measure_no = _resolve_measure_no(value, row_index, measure_row_min, measure_row_step)

        if len(debug_info) < 10:
            debug_info.append(
                f"行{row_index}: 値={value!r}, 型={type(value).__name__}, 数式={formula}, 解決No={measure_no}"
            )

        if measure_no is None:
            continue
        measure_no_to_row[measure_no] = row_index

    target_row = not_required_row

    e_col = column_index_from_string("E")
    target_e_cell = _get_writable_cell(ws, target_row, e_col)
    target_e_cell.value = "測定不要"

    flag_col_start = column_index_from_string("L")
    flag_col_end = column_index_from_string("SR")

    written_count = 0
    for no in target_nos:
        row_index = measure_no_to_row.get(no)
        if row_index is None:
            continue

        for col_idx in range(flag_col_start, flag_col_end + 1):
            col_letter = get_column_letter(col_idx)
            target_cell = f"{col_letter}{target_row}"
            current_cell = ws.cell(row_index, col_idx)
            if not _can_overwrite_with_formula(current_cell.value):
                continue
            current_cell.value = _build_not_required_overlay_formula(
                current_cell.value,
                target_cell,
                formula_arg_sep,
            )
        written_count += 1

    if written_count == 0 and target_nos:
        available_nos = sorted(measure_no_to_row.keys())
        available_nos_str = ", ".join(map(str, available_nos[:20]))
        if len(available_nos) > 20:
            available_nos_str += f" ... (他{len(available_nos) - 20}件)"

        debug_str = "\n".join(debug_info)

        raise ValueError(
            f"指定したNo.の行が見つかりませんでした。\n\n"
            f"【指定したNo.】\n{target_nos}\n\n"
            f"【読み取れた測定No】\n{available_nos_str}\n\n"
            f"【設定】\n"
            f"- 測定No列: {measure_no_col}列\n"
            f"- 測定行範囲: {measure_row_min}〜{measure_row_max}行目\n"
            f"- 読み取れた測定No件数: {len(measure_no_to_row)}件\n\n"
            f"【デバッグ情報（先頭10行）】\n{debug_str}\n\n"
            f"※測定Noが数式の場合、Excelで一度ファイルを開いて保存してから再実行してください。"
        )

    normalized_count = _normalize_single_cell_array_formulas_in_column(
        ws,
        "B",
        row_start=measure_row_min,
        row_end=measure_row_max,
    )
    if normalized_count:
        print(f"[info] B列の単一セル配列数式を通常数式へ変換: {normalized_count}件")

    _mark_workbook_for_full_recalc(wb)
    saved_path = _save_workbook_atomic(wb, out_path, parent=parent)
    _close_workbook_quietly(wb_values)
    _close_workbook_quietly(wb)
    _restore_package_parts_from_source(xlsx_path, saved_path)
    _force_excel_recalc_and_save(saved_path)
    return saved_path
