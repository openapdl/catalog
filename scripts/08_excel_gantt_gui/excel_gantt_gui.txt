from __future__ import annotations


import os
import subprocess
import sys
from pathlib import Path


# 換到其他電腦時，只需修改下一行的 pythonw.exe 絕對路徑。
PYTHONW_EXE = Path(
    r"C:\Users\Desktop\015_python\python3.8\win64\pythonw.exe"
)


def _executable_path_key(path: Path) -> str:
    """Return a case-insensitive absolute identity for a Windows executable."""

    return os.path.normcase(os.path.abspath(str(path)))


def _show_bootstrap_error(detail: str) -> None:
    """Show an error even when the current interpreter has no console."""

    if sys.stderr is not None:
        sys.stderr.write(detail + "\n")
    try:
        import ctypes

        ctypes.windll.user32.MessageBoxW(0, detail, "Excel 甘特圖產生器", 0x10)
    except (AttributeError, OSError):
        pass


def _ensure_configured_python() -> None:
    """Relaunch this file with the configured portable Python when necessary."""

    configured_pythonw = PYTHONW_EXE.expanduser()
    configured_python = configured_pythonw.with_name("python.exe")
    current_key = _executable_path_key(Path(sys.executable))
    configured_keys = {
        _executable_path_key(configured_pythonw),
        _executable_path_key(configured_python),
    }
    if current_key in configured_keys:
        return

    if not configured_pythonw.is_file():
        _show_bootstrap_error(
            "找不到指定的 Python 執行檔：\n{}\n\n"
            "請用文字編輯器開啟 excel_gantt_gui.pyw，修改檔案最上方的 PYTHONW_EXE。".format(configured_pythonw)
        )
        raise SystemExit(1)

    current_is_console_python = Path(sys.executable).name.lower() == "python.exe"
    target_executable = (
        configured_python if current_is_console_python and configured_python.is_file() else configured_pythonw
    )
    script_path = Path(__file__).resolve()
    command = [str(target_executable), str(script_path)] + sys.argv[1:]
    try:
        if current_is_console_python:
            return_code = subprocess.call(command, cwd=str(script_path.parent))
            raise SystemExit(return_code)
        subprocess.Popen(command, cwd=str(script_path.parent))
    except OSError as exc:
        _show_bootstrap_error("無法使用指定的 Python 啟動程式：\n{}\n\n{}".format(target_executable, exc))
        raise SystemExit(1)
    raise SystemExit(0)


if __name__ == "__main__":
    _ensure_configured_python()


import argparse
import calendar
import tkinter as tk
from dataclasses import dataclass
from datetime import date, timedelta
from tkinter import colorchooser, filedialog, messagebox, ttk
from typing import Callable, Iterable

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation


APP_DIR = Path(__file__).resolve().parent
DEFAULT_OUTPUT_NAME = "Excel_Gantt_範例.xlsx"
DEFAULT_PROJECT_MONTH = 10

TITLE_ROW = 1
NOTE_ROW = 2
QUARTER_ROW = 3
MONTH_ROW = 4
PERIOD_ROW = 5
FIRST_TASK_ROW = 6

TASK_COLUMN = 1
START_COLUMN = 2
END_COLUMN = 3
FIRST_TIMELINE_COLUMN = 4
EXTRA_TASK_ROWS = 20
MINIMUM_TASK_ROWS = 30
MAX_TIMELINE_SLOTS = 1200

HEADER_BLUE = "087DB8"
HEADER_DARK_BLUE = "056A9E"
HEADER_LIGHT = "E7EEF4"
GRID_GRAY = "B8C0C7"
ERROR_RED = "FFD9D9"
DEFAULT_AVERAGE_COLOR = "FFFF00"

SCALE_LABELS = {
    "日": "每日顯示一格",
    "週": "每 7 天顯示一個週區塊，底層保留每日精度",
}


@dataclass(frozen=True)
class Task:
    name: str
    start: date
    end: date


def format_date(value: date) -> str:
    return value.strftime("%Y-%m-%d")


def add_months(value: date, month_count: int) -> date:
    absolute_month = value.year * 12 + value.month - 1 + month_count
    year, month_index = divmod(absolute_month, 12)
    return date(year, month_index + 1, 1)


def distribute_tasks_for_range(
    tasks: Iterable[Task],
    overall_start: date,
    overall_end: date,
) -> list[Task]:
    task_list = list(tasks)
    if not task_list:
        return []
    if overall_end < overall_start:
        raise ValueError("總結束日期不可早於總開始日期。")

    total_days = (overall_end - overall_start).days
    base_days, remainder = divmod(total_days, len(task_list))
    first_larger_task = len(task_list) - remainder

    distributed: list[Task] = []
    cursor = overall_start
    for task_index, task in enumerate(task_list):
        duration_days = base_days + (
            1 if remainder and task_index >= first_larger_task else 0
        )
        task_end = cursor + timedelta(days=duration_days)
        distributed.append(Task(task.name, cursor, task_end))
        cursor = task_end
    return distributed


def distribute_tasks_evenly(tasks: Iterable[Task]) -> list[Task]:
    task_list = list(tasks)
    if not task_list:
        return []
    overall_start = min(task.start for task in task_list)
    overall_end = max(task.end for task in task_list)
    return distribute_tasks_for_range(task_list, overall_start, overall_end)


def sample_tasks(reference_day: date | None = None) -> list[Task]:
    reference_day = reference_day or date.today()
    overall_start = date(reference_day.year, DEFAULT_PROJECT_MONTH, 1)
    overall_end = date(
        reference_day.year,
        DEFAULT_PROJECT_MONTH,
        calendar.monthrange(reference_day.year, DEFAULT_PROJECT_MONTH)[1],
    )
    tasks = [
        Task(name, overall_start, overall_end)
        for name in ("Planning", "Research", "Design", "Implementation", "Follow up")
    ]
    return distribute_tasks_for_range(tasks, overall_start, overall_end)


def build_timeline(
    tasks: Iterable[Task],
    scale: str,
    exact_daily: bool = False,
) -> list[date]:
    task_list = list(tasks)
    earliest = min(task.start for task in task_list)
    latest = max(task.end for task in task_list)

    if exact_daily or scale == "日":
        slot_count = (latest - earliest).days + 1
        dates = [earliest + timedelta(days=offset) for offset in range(slot_count)]
    elif scale == "週":
        dates = [earliest]
        while dates[-1] < latest:
            next_date = dates[-1] + timedelta(weeks=1)
            dates.append(min(next_date, latest))
    else:
        raise ValueError(f"不支援的時間刻度：{scale}")

    # 即使使用週刻度，也保留每個工作的精確邊界日期。
    # 如此黃色區塊才能從 B 欄開始日對齊到 C 欄結束日。
    task_boundary_dates = {
        boundary_date
        for task in task_list
        for boundary_date in (task.start, task.end)
    }
    dates = sorted(set(dates) | task_boundary_dates)

    if len(dates) > MAX_TIMELINE_SLOTS:
        raise ValueError(
            f"目前刻度會建立 {len(dates)} 個時間欄，超過上限 {MAX_TIMELINE_SLOTS}。"
            "請改用較大的時間刻度。"
        )
    return dates


def clean_hex_color(value: str) -> str:
    cleaned = value.strip().lstrip("#").upper()
    if len(cleaned) != 6 or any(character not in "0123456789ABCDEF" for character in cleaned):
        raise ValueError("色彩必須是 6 位十六進位色碼。")
    return cleaned


def contrast_text_color(value: str) -> str:
    color = clean_hex_color(value)
    red, green, blue = (int(color[index:index + 2], 16) for index in (0, 2, 4))
    luminance = 0.299 * red + 0.587 * green + 0.114 * blue
    return "#000000" if luminance >= 150 else "#FFFFFF"


def quarter_label(value: date) -> str:
    return f"Q{((value.month - 1) // 3) + 1} {value.year}"


def month_label(value: date) -> str:
    return f"{value.year}/{value.month:02d}"


def group_ranges(labels: list[str]) -> list[tuple[int, int, str]]:
    groups: list[tuple[int, int, str]] = []
    group_start = 0
    for index in range(1, len(labels) + 1):
        if index == len(labels) or labels[index] != labels[group_start]:
            groups.append((group_start, index - 1, labels[group_start]))
            group_start = index
    return groups


def period_ranges(
    timeline_dates: list[date],
    scale: str,
) -> list[tuple[int, int, date]]:
    if scale == "日":
        return [
            (index, index, timeline_date)
            for index, timeline_date in enumerate(timeline_dates)
        ]
    if scale == "週":
        return [
            (
                start_index,
                min(start_index + 6, len(timeline_dates) - 1),
                timeline_dates[start_index],
            )
            for start_index in range(0, len(timeline_dates), 7)
        ]
    raise ValueError(f"不支援的時間刻度：{scale}")


def merge_timeline_headers(
    worksheet,
    row: int,
    labels: list[str],
    fill_color: str,
    font_color: str,
) -> None:
    header_fill = PatternFill("solid", fgColor=fill_color)
    header_font = Font(color=font_color, bold=True, size=11)
    thin_border = Border(
        left=Side(style="thin", color="FFFFFF"),
        right=Side(style="thin", color="FFFFFF"),
        top=Side(style="thin", color="FFFFFF"),
        bottom=Side(style="thin", color="FFFFFF"),
    )

    for start_index, end_index, label in group_ranges(labels):
        start_column = FIRST_TIMELINE_COLUMN + start_index
        end_column = FIRST_TIMELINE_COLUMN + end_index
        for column in range(start_column, end_column + 1):
            cell = worksheet.cell(row=row, column=column)
            cell.fill = header_fill
            cell.border = thin_border
        if end_column > start_column:
            worksheet.merge_cells(
                start_row=row,
                start_column=start_column,
                end_row=row,
                end_column=end_column,
            )
        cell = worksheet.cell(row=row, column=start_column, value=label)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")


def create_excel_gantt(
    output_path: Path,
    project_title: str,
    tasks: list[Task],
    scale: str,
    average_color: str = DEFAULT_AVERAGE_COLOR,
    overall_start: date | None = None,
    overall_end: date | None = None,
) -> None:
    if not tasks:
        raise ValueError("至少需要一項工作。")
    if (overall_start is None) != (overall_end is None):
        raise ValueError("總開始日期與總結束日期必須同時設定。")
    if overall_start is not None and overall_end is not None:
        tasks = distribute_tasks_for_range(tasks, overall_start, overall_end)
    else:
        if any(task.end < task.start for task in tasks):
            raise ValueError("所有工作的結束日期都必須晚於或等於開始日期。")
        tasks = distribute_tasks_evenly(tasks)

    output_path = output_path.resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    average_color = clean_hex_color(average_color)
    # 動態條件格式需要每天都有對應欄位，才能在 Excel 內改日期後精確移動。
    timeline_dates = build_timeline(tasks, scale, exact_daily=True)
    last_timeline_column = FIRST_TIMELINE_COLUMN + len(timeline_dates) - 1
    last_timeline_letter = get_column_letter(last_timeline_column)
    task_row_count = max(MINIMUM_TASK_ROWS, len(tasks) + EXTRA_TASK_ROWS)
    last_task_row = FIRST_TASK_ROW + task_row_count - 1
    last_data_row = FIRST_TASK_ROW + len(tasks) - 1

    workbook = Workbook()
    workbook.calculation.calcMode = "auto"
    workbook.calculation.fullCalcOnLoad = True
    worksheet = workbook.active
    worksheet.title = "甘特圖"
    worksheet.sheet_view.showGridLines = False
    worksheet.sheet_view.zoomScale = 85
    worksheet.freeze_panes = f"{get_column_letter(FIRST_TIMELINE_COLUMN)}{FIRST_TASK_ROW}"

    worksheet.merge_cells(
        start_row=TITLE_ROW,
        start_column=TASK_COLUMN,
        end_row=NOTE_ROW,
        end_column=last_timeline_column,
    )
    title_cell = worksheet.cell(row=TITLE_ROW, column=TASK_COLUMN, value=project_title)
    title_cell.font = Font(name="Microsoft JhengHei", size=26, bold=True, color="111111")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    worksheet.row_dimensions[TITLE_ROW].height = 36
    worksheet.row_dimensions[NOTE_ROW].height = 36

    left_headers = {
        TASK_COLUMN: "工作項目",
        START_COLUMN: "開始日期",
        END_COLUMN: "結束日期",
    }
    header_fill = PatternFill("solid", fgColor=HEADER_BLUE)
    white_border = Border(
        left=Side(style="thin", color="FFFFFF"),
        right=Side(style="thin", color="FFFFFF"),
        top=Side(style="thin", color="FFFFFF"),
        bottom=Side(style="thin", color="FFFFFF"),
    )
    for column, label in left_headers.items():
        worksheet.merge_cells(
            start_row=QUARTER_ROW,
            start_column=column,
            end_row=PERIOD_ROW,
            end_column=column,
        )
        cell = worksheet.cell(row=QUARTER_ROW, column=column, value=label)
        cell.fill = header_fill
        cell.font = Font(name="Microsoft JhengHei", color="FFFFFF", bold=True, size=12)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = white_border

    quarter_labels = [quarter_label(value) for value in timeline_dates]
    month_labels = [month_label(value) for value in timeline_dates]
    merge_timeline_headers(worksheet, QUARTER_ROW, quarter_labels, HEADER_DARK_BLUE, "FFFFFF")
    merge_timeline_headers(worksheet, MONTH_ROW, month_labels, HEADER_BLUE, "FFFFFF")

    period_fill = PatternFill("solid", fgColor=HEADER_LIGHT)
    period_border = Border(
        left=Side(style="thin", color=GRID_GRAY),
        right=Side(style="thin", color=GRID_GRAY),
        top=Side(style="thin", color=GRID_GRAY),
        bottom=Side(style="thin", color=GRID_GRAY),
    )
    for index, _value in enumerate(timeline_dates):
        column = FIRST_TIMELINE_COLUMN + index
        cell = worksheet.cell(row=PERIOD_ROW, column=column)
        cell.fill = period_fill
        cell.font = Font(name="Microsoft JhengHei", size=9, color="222222")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = period_border

    period_number_format = {"日": "d", "週": "mm/dd"}[scale]
    for start_index, end_index, period_start in period_ranges(timeline_dates, scale):
        start_column = FIRST_TIMELINE_COLUMN + start_index
        end_column = FIRST_TIMELINE_COLUMN + end_index
        if end_column > start_column:
            worksheet.merge_cells(
                start_row=PERIOD_ROW,
                start_column=start_column,
                end_row=PERIOD_ROW,
                end_column=end_column,
            )
        cell = worksheet.cell(row=PERIOD_ROW, column=start_column, value=period_start)
        cell.number_format = period_number_format
        cell.fill = period_fill
        cell.font = Font(name="Microsoft JhengHei", size=9, color="222222")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = period_border

    worksheet.row_dimensions[QUARTER_ROW].height = 24
    worksheet.row_dimensions[MONTH_ROW].height = 23
    worksheet.row_dimensions[PERIOD_ROW].height = 24
    worksheet.column_dimensions["A"].width = 30
    worksheet.column_dimensions["B"].width = 14
    worksheet.column_dimensions["C"].width = 14
    timeline_width = {"日": 4.0, "週": 1.2}[scale]
    for column in range(FIRST_TIMELINE_COLUMN, last_timeline_column + 1):
        worksheet.column_dimensions[get_column_letter(column)].width = timeline_width

    grid_side = Side(style="thin", color=GRID_GRAY)
    grid_border = Border(left=grid_side, right=grid_side, top=grid_side, bottom=grid_side)
    centered = Alignment(horizontal="center", vertical="center")

    for offset in range(task_row_count):
        row = FIRST_TASK_ROW + offset
        worksheet.row_dimensions[row].height = 32
        left_fill = PatternFill("solid", fgColor="FFFFFF")
        for column in range(TASK_COLUMN, END_COLUMN + 1):
            cell = worksheet.cell(row=row, column=column)
            cell.fill = left_fill
            cell.border = grid_border
            cell.font = Font(name="Microsoft JhengHei", size=11, color="1F252B")
            cell.alignment = Alignment(horizontal="left" if column == TASK_COLUMN else "center", vertical="center")
        worksheet.cell(row=row, column=START_COLUMN).number_format = "yyyy/mm/dd"
        worksheet.cell(row=row, column=END_COLUMN).number_format = "yyyy/mm/dd"

        for index, _timeline_date in enumerate(timeline_dates):
            column = FIRST_TIMELINE_COLUMN + index
            cell = worksheet.cell(row=row, column=column)
            cell.fill = PatternFill("solid", fgColor="FFFFFF")
            cell.border = Border(
                left=grid_side if index == 0 else Side(),
                right=grid_side if index == len(timeline_dates) - 1 else Side(),
                top=grid_side,
                bottom=grid_side,
            )
            cell.alignment = centered

    for offset, task in enumerate(tasks):
        row = FIRST_TASK_ROW + offset
        worksheet.cell(row=row, column=TASK_COLUMN, value=task.name)
        worksheet.cell(row=row, column=START_COLUMN, value=task.start)
        worksheet.cell(row=row, column=END_COLUMN, value=task.end)

    average_argb = f"FF{average_color}"
    average_fill = PatternFill(
        "solid",
        fgColor=average_argb,
        bgColor=average_argb,
    )
    timeline_body_range = (
        f"{get_column_letter(FIRST_TIMELINE_COLUMN)}{FIRST_TASK_ROW}:"
        f"{last_timeline_letter}{last_task_row}"
    )
    axis_start = timeline_dates[0]
    axis_start_formula = f"DATE({axis_start.year},{axis_start.month},{axis_start.day})"
    timeline_date_formula = (
        f"{axis_start_formula}+COLUMN()-"
        f"COLUMN(${get_column_letter(FIRST_TIMELINE_COLUMN)}${PERIOD_ROW})"
    )
    worksheet.conditional_formatting.add(
        timeline_body_range,
        FormulaRule(
            formula=[
                f'AND($B{FIRST_TASK_ROW}<>"",$C{FIRST_TASK_ROW}<>"",'
                f'{timeline_date_formula}>=$B{FIRST_TASK_ROW},'
                f'{timeline_date_formula}<=$C{FIRST_TASK_ROW})'
            ],
            fill=average_fill,
            stopIfTrue=True,
        ),
    )

    date_validation = DataValidation(
        type="date",
        operator="between",
        formula1="DATE(1900,1,1)",
        formula2="DATE(9999,12,31)",
        allow_blank=True,
    )
    date_validation.errorTitle = "日期格式錯誤"
    date_validation.error = "請輸入有效的 Excel 日期。"
    date_validation.promptTitle = "排程日期"
    date_validation.prompt = (
        "在現有時間軸內改日期，黃色區塊會自動更新；"
        "若要改變總時間範圍，請由 GUI 重新產生。"
    )
    date_validation.showErrorMessage = True
    date_validation.showInputMessage = True
    worksheet.add_data_validation(date_validation)
    date_validation.add(f"B{FIRST_TASK_ROW}:C{last_task_row}")

    invalid_date_argb = f"FF{ERROR_RED}"
    invalid_date_fill = PatternFill(
        "solid",
        fgColor=invalid_date_argb,
        bgColor=invalid_date_argb,
    )
    worksheet.conditional_formatting.add(
        f"C{FIRST_TASK_ROW}:C{last_task_row}",
        FormulaRule(
            formula=[f'AND($B{FIRST_TASK_ROW}<>"",$C{FIRST_TASK_ROW}<$B{FIRST_TASK_ROW})'],
            fill=invalid_date_fill,
        ),
    )

    worksheet.print_area = f"A1:{last_timeline_letter}{last_data_row}"
    worksheet.print_title_rows = f"1:{PERIOD_ROW}"
    worksheet.page_setup.orientation = "landscape"
    worksheet.page_setup.fitToWidth = 1
    worksheet.page_setup.fitToHeight = 0
    worksheet.sheet_properties.pageSetUpPr.fitToPage = True
    worksheet.sheet_properties.outlinePr.summaryBelow = True
    worksheet.sheet_view.selection[0].activeCell = f"A{FIRST_TASK_ROW}"
    worksheet.sheet_view.selection[0].sqref = f"A{FIRST_TASK_ROW}"

    workbook.active = 0
    workbook.save(output_path)


def validate_generated_workbook(
    path: Path,
    expected_task_count: int,
    expected_average_color: str = DEFAULT_AVERAGE_COLOR,
    expected_scale: str | None = None,
) -> None:
    workbook = load_workbook(path, read_only=False, data_only=False)
    if "甘特圖" not in workbook.sheetnames:
        raise RuntimeError("輸出檔缺少必要工作表。")
    worksheet = workbook["甘特圖"]
    expected_title_merge = (
        f"{get_column_letter(TASK_COLUMN)}{TITLE_ROW}:"
        f"{get_column_letter(worksheet.max_column)}{NOTE_ROW}"
    )
    title_merges = {
        str(merged_range)
        for merged_range in worksheet.merged_cells.ranges
        if merged_range.min_row <= NOTE_ROW
    }
    if title_merges != {expected_title_merge}:
        raise RuntimeError("「Gantt Chart」標題區未完整合併成單一儲存格。")
    actual_tasks = sum(
        1
        for row in range(FIRST_TASK_ROW, FIRST_TASK_ROW + expected_task_count)
        if worksheet.cell(row=row, column=TASK_COLUMN).value
    )
    if actual_tasks != expected_task_count:
        raise RuntimeError("輸出檔的工作數量不正確。")
    if len(worksheet.conditional_formatting) < 2:
        raise RuntimeError("輸出檔缺少動態色塊或日期錯誤檢查規則。")

    generated_dates: list[tuple[date, date]] = []
    for row in range(FIRST_TASK_ROW, FIRST_TASK_ROW + expected_task_count):
        start_value = worksheet.cell(row=row, column=START_COLUMN).value
        end_value = worksheet.cell(row=row, column=END_COLUMN).value
        start_date = start_value.date() if hasattr(start_value, "date") else start_value
        end_date = end_value.date() if hasattr(end_value, "date") else end_value
        generated_dates.append((start_date, end_date))

    for index in range(expected_task_count - 1):
        if generated_dates[index][1] != generated_dates[index + 1][0]:
            raise RuntimeError("工作結束日期未與下一項開始日期相接。")
    durations = [(end - start).days for start, end in generated_dates]
    if durations and max(durations) - min(durations) > 1:
        raise RuntimeError("各工作日期未平均分配。")

    first_header_value = worksheet.cell(
        row=PERIOD_ROW,
        column=FIRST_TIMELINE_COLUMN,
    ).value
    first_header_date = (
        first_header_value.date()
        if hasattr(first_header_value, "date")
        else first_header_value
    )
    timeline_slot_count = worksheet.max_column - FIRST_TIMELINE_COLUMN + 1
    timeline_header_dates = [
        first_header_date + timedelta(days=offset)
        for offset in range(timeline_slot_count)
    ]
    last_header_date = timeline_header_dates[-1]
    if first_header_date != generated_dates[0][0]:
        raise RuntimeError("時間軸開始日期與第一項開始日期不一致。")
    if last_header_date != generated_dates[-1][1]:
        raise RuntimeError("時間軸結束日期與最後一項結束日期不一致。")

    if expected_scale is not None:
        expected_period_merges: set[str] = set()
        for start_index, end_index, period_start in period_ranges(
            timeline_header_dates,
            expected_scale,
        ):
            start_column = FIRST_TIMELINE_COLUMN + start_index
            end_column = FIRST_TIMELINE_COLUMN + end_index
            start_letter = get_column_letter(start_column)
            end_letter = get_column_letter(end_column)
            if end_column > start_column:
                expected_period_merges.add(
                    f"{start_letter}{PERIOD_ROW}:{end_letter}{PERIOD_ROW}"
                )
            anchor_value = worksheet.cell(row=PERIOD_ROW, column=start_column).value
            anchor_date = anchor_value.date() if hasattr(anchor_value, "date") else anchor_value
            if anchor_date != period_start:
                raise RuntimeError("時間刻度標題的起始日期不正確。")

        actual_period_merges = {
            str(merged_range)
            for merged_range in worksheet.merged_cells.ranges
            if merged_range.min_row == PERIOD_ROW
            and merged_range.max_row == PERIOD_ROW
            and merged_range.min_col >= FIRST_TIMELINE_COLUMN
        }
        if actual_period_merges != expected_period_merges:
            raise RuntimeError(f"時間刻度未依「{expected_scale}」正確合併顯示。")

    timeline_date_set = set(timeline_header_dates)
    if any(
        start_date not in timeline_date_set or end_date not in timeline_date_set
        for start_date, end_date in generated_dates
    ):
        raise RuntimeError("時間軸缺少工作的開始或結束日期。")

    expected_average_rgb = f"FF{clean_hex_color(expected_average_color)}"
    body_merged_ranges = [
        merged_range
        for merged_range in worksheet.merged_cells.ranges
        if merged_range.min_row >= FIRST_TASK_ROW
        and merged_range.max_row == merged_range.min_row
        and merged_range.min_col >= FIRST_TIMELINE_COLUMN
    ]
    if body_merged_ranges:
        raise RuntimeError("動態色塊區不可使用實際合併儲存格。")

    non_white_cells: list[str] = []
    divided_cells: list[str] = []
    for row in range(FIRST_TASK_ROW, worksheet.max_row + 1):
        for column in range(FIRST_TIMELINE_COLUMN, worksheet.max_column + 1):
            cell = worksheet.cell(row=row, column=column)
            fill_rgb = cell.fill.fgColor.rgb if cell.fill.fgColor.type == "rgb" else None
            if cell.fill.fill_type != "solid" or fill_rgb not in {"00FFFFFF", "FFFFFFFF"}:
                non_white_cells.append(cell.coordinate)
            if (
                (column > FIRST_TIMELINE_COLUMN and cell.border.left.style is not None)
                or (column < worksheet.max_column and cell.border.right.style is not None)
            ):
                divided_cells.append(cell.coordinate)
    if non_white_cells:
        preview = "、".join(non_white_cells[:5])
        raise RuntimeError(f"動態色塊的基礎儲存格不是白色：{preview}")
    if divided_cells:
        preview = "、".join(divided_cells[:5])
        raise RuntimeError(f"動態色塊區仍存在內部垂直分隔線：{preview}")

    dynamic_range = (
        f"{get_column_letter(FIRST_TIMELINE_COLUMN)}{FIRST_TASK_ROW}:"
        f"{get_column_letter(worksheet.max_column)}{worksheet.max_row}"
    )
    axis_start_formula = (
        f"DATE({first_header_date.year},{first_header_date.month},{first_header_date.day})"
    )
    timeline_date_formula = (
        f"{axis_start_formula}+COLUMN()-"
        f"COLUMN(${get_column_letter(FIRST_TIMELINE_COLUMN)}${PERIOD_ROW})"
    )
    dynamic_formula = (
        f'AND($B{FIRST_TASK_ROW}<>"",$C{FIRST_TASK_ROW}<>"",'
        f'{timeline_date_formula}>=$B{FIRST_TASK_ROW},'
        f'{timeline_date_formula}<=$C{FIRST_TASK_ROW})'
    )
    dynamic_rule_found = False
    for conditional_range in worksheet.conditional_formatting:
        if str(conditional_range.sqref) != dynamic_range:
            continue
        for rule in worksheet.conditional_formatting[conditional_range]:
            rule_rgb = None
            if rule.dxf is not None and rule.dxf.fill is not None:
                rule_rgb = rule.dxf.fill.fgColor.rgb
            if rule.formula == [dynamic_formula] and rule_rgb == expected_average_rgb:
                dynamic_rule_found = True
                break
    if not dynamic_rule_found:
        raise RuntimeError("輸出檔缺少可依 B、C 欄日期自動更新的色塊規則。")
    workbook.close()


class MonthCalendarPopup(tk.Toplevel):
    WEEKDAY_LABELS = ("一", "二", "三", "四", "五", "六", "日")

    def __init__(
        self,
        master,
        initial_date: date,
        on_select: Callable[[date], None],
        on_close: Callable[[], None],
    ) -> None:
        super().__init__(master)
        self.selected_date = initial_date
        self.display_month = date(initial_date.year, initial_date.month, 1)
        self.on_select = on_select
        self.on_close = on_close

        self.title("選擇日期")
        self.resizable(False, False)
        self.transient(master.winfo_toplevel())
        self.configure(background="#FFFFFF")
        self.protocol("WM_DELETE_WINDOW", self.close)
        self.bind("<Escape>", lambda _event: self.close())

        self._build_ui()
        self._render_month()
        self.after_idle(self._position_and_focus)

    def _build_ui(self) -> None:
        header = ttk.Frame(self, padding=(8, 8, 8, 5))
        header.pack(fill="x")
        ttk.Button(header, text="‹", width=3, command=lambda: self._change_month(-1)).pack(side="left")
        self.month_title = ttk.Label(
            header,
            anchor="center",
            font=("Microsoft JhengHei", 11, "bold"),
        )
        self.month_title.pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(header, text="›", width=3, command=lambda: self._change_month(1)).pack(side="right")

        weekday_frame = tk.Frame(self, background="#FFFFFF", padx=8)
        weekday_frame.pack(fill="x")
        for column, label in enumerate(self.WEEKDAY_LABELS):
            foreground = "#C24A4A" if column >= 5 else "#5A6570"
            tk.Label(
                weekday_frame,
                text=label,
                width=4,
                background="#FFFFFF",
                foreground=foreground,
                font=("Microsoft JhengHei", 9, "bold"),
            ).grid(row=0, column=column, padx=1, pady=2)

        self.day_frame = tk.Frame(self, background="#FFFFFF", padx=8, pady=2)
        self.day_frame.pack(fill="both")

        footer = ttk.Frame(self, padding=(8, 5, 8, 8))
        footer.pack(fill="x")
        ttk.Button(footer, text="今天", command=self._select_today).pack(side="left")
        ttk.Button(footer, text="取消", command=self.close).pack(side="right")

    def _render_month(self) -> None:
        self.month_title.configure(
            text=f"{self.display_month.year} 年 {self.display_month.month} 月"
        )
        for child in self.day_frame.winfo_children():
            child.destroy()

        weeks = calendar.Calendar(firstweekday=calendar.MONDAY).monthdayscalendar(
            self.display_month.year,
            self.display_month.month,
        )
        while len(weeks) < 6:
            weeks.append([0] * 7)

        today = date.today()
        for row, week in enumerate(weeks):
            for column, day_number in enumerate(week):
                if day_number == 0:
                    tk.Label(
                        self.day_frame,
                        text="",
                        width=4,
                        height=2,
                        background="#FFFFFF",
                    ).grid(row=row, column=column, padx=1, pady=1)
                    continue

                cell_date = date(
                    self.display_month.year,
                    self.display_month.month,
                    day_number,
                )
                is_selected = cell_date == self.selected_date
                is_today = cell_date == today
                is_weekend = column >= 5
                background = "#087DB8" if is_selected else ("#F3F5F7" if is_weekend else "#FFFFFF")
                foreground = "#FFFFFF" if is_selected else ("#C23B3B" if is_today else "#20262C")
                font = ("Microsoft JhengHei", 9, "bold" if is_selected or is_today else "normal")
                button = tk.Button(
                    self.day_frame,
                    text=str(day_number),
                    width=4,
                    height=1,
                    borderwidth=0,
                    relief="flat",
                    background=background,
                    foreground=foreground,
                    activebackground="#B9DDF0",
                    activeforeground="#17212B",
                    font=font,
                    cursor="hand2",
                    command=lambda selected=cell_date: self._select_date(selected),
                )
                button.grid(row=row, column=column, padx=1, pady=1, ipady=3)

    def _change_month(self, offset: int) -> None:
        self.display_month = add_months(self.display_month, offset)
        self._render_month()

    def _select_date(self, selected_date: date) -> None:
        self.on_select(selected_date)
        self.close()

    def _select_today(self) -> None:
        self.on_select(date.today())
        self.close()

    def _position_and_focus(self) -> None:
        self.update_idletasks()
        desired_x = self.master.winfo_rootx()
        desired_y = self.master.winfo_rooty() + self.master.winfo_height()
        maximum_x = max(0, self.winfo_screenwidth() - self.winfo_width())
        maximum_y = max(0, self.winfo_screenheight() - self.winfo_height())
        self.geometry(f"+{min(desired_x, maximum_x)}+{min(desired_y, maximum_y)}")
        self.grab_set()
        self.focus_force()

    def close(self) -> None:
        try:
            if self.grab_current() is self:
                self.grab_release()
        except tk.TclError:
            pass
        self.on_close()
        self.destroy()


class CalendarDatePicker(ttk.Frame):
    def __init__(
        self,
        master,
        initial_date: date,
        on_change: Callable[[date], None] | None = None,
    ) -> None:
        super().__init__(master)
        self.selected_date = initial_date
        self.on_change = on_change
        self.display_value = tk.StringVar(value=format_date(initial_date))
        self.popup: MonthCalendarPopup | None = None

        self.entry = ttk.Entry(
            self,
            textvariable=self.display_value,
            state="readonly",
            width=12,
            cursor="hand2",
        )
        self.entry.pack(side="left")
        self.entry.bind("<Button-1>", self.open_calendar)
        ttk.Button(self, text="▼", width=3, command=self.open_calendar).pack(side="left", padx=(2, 0))

    def open_calendar(self, _event=None):
        if self.popup is not None and self.popup.winfo_exists():
            self.popup.lift()
            self.popup.focus_force()
            return "break"
        self.popup = MonthCalendarPopup(
            self,
            initial_date=self.selected_date,
            on_select=self.set_date,
            on_close=self._popup_closed,
        )
        return "break"

    def _popup_closed(self) -> None:
        self.popup = None

    def get_date(self) -> date:
        return self.selected_date

    def set_date(self, value: date, notify: bool = True) -> None:
        changed = value != self.selected_date
        self.selected_date = value
        self.display_value.set(format_date(value))
        if changed and notify and self.on_change is not None:
            self.on_change(value)


class ExcelGanttApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Excel 甘特圖產生器")
        self.geometry("1050x700")
        self.minsize(920, 620)

        self.tasks = sample_tasks()
        self.initial_overall_start = self.tasks[0].start
        self.initial_overall_end = self.tasks[-1].end
        self.selected_index: int | None = None
        self.average_color = f"#{DEFAULT_AVERAGE_COLOR}"

        self.project_title = tk.StringVar(value="Gantt Chart")
        self.scale = tk.StringVar(value="週")
        self.task_name = tk.StringVar()
        self.scale_help = tk.StringVar(value=SCALE_LABELS[self.scale.get()])
        self.range_text = tk.StringVar()
        self.status_text = tk.StringVar(value="請編輯排程，完成後按「產生 Excel」。")

        self.configure(background="#F4F6F8")
        self._configure_style()
        self._build_ui()
        self._refresh_tree()

    def _configure_style(self) -> None:
        style = ttk.Style(self)
        available_themes = style.theme_names()
        if "vista" in available_themes:
            style.theme_use("vista")
        elif "clam" in available_themes:
            style.theme_use("clam")
        style.configure("Subtitle.TLabel", font=("Microsoft JhengHei", 10), foreground="#607080")
        style.configure("Section.TLabelframe.Label", font=("Microsoft JhengHei", 11, "bold"))
        style.configure("Treeview", font=("Microsoft JhengHei", 10), rowheight=29)
        style.configure("Treeview.Heading", font=("Microsoft JhengHei", 10, "bold"))
        style.configure("Primary.TButton", font=("Microsoft JhengHei", 11, "bold"), padding=(15, 9))

    def _build_ui(self) -> None:
        outer = ttk.Frame(self, padding=18)
        outer.pack(fill="both", expand=True)

        ttk.Label(
            outer,
            text="只需設定整張圖的總期間，各工作會自動平均分配並連續相接。",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(0, 14))

        settings = ttk.LabelFrame(outer, text="專案設定", style="Section.TLabelframe", padding=12)
        settings.pack(fill="x")
        settings.columnconfigure(1, weight=1)
        settings.columnconfigure(5, weight=1)

        ttk.Label(settings, text="圖表標題").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Entry(settings, textvariable=self.project_title, width=32).grid(row=0, column=1, sticky="ew")
        ttk.Label(settings, text="時間刻度").grid(row=0, column=2, sticky="w", padx=(20, 8))
        scale_box = ttk.Combobox(
            settings,
            textvariable=self.scale,
            values=tuple(SCALE_LABELS),
            state="readonly",
            width=8,
        )
        scale_box.grid(row=0, column=3, sticky="w")
        scale_box.bind("<<ComboboxSelected>>", self._on_scale_changed)
        self.color_button = tk.Button(
            settings,
            text="動態色塊顏色",
            command=self._choose_color,
            bg=self.average_color,
            activebackground=self.average_color,
            fg=contrast_text_color(self.average_color),
            relief="flat",
            padx=12,
            pady=5,
            font=("Microsoft JhengHei", 9, "bold"),
        )
        self.color_button.grid(row=0, column=4, padx=(20, 8))
        ttk.Label(settings, textvariable=self.scale_help, style="Subtitle.TLabel").grid(
            row=0,
            column=5,
            sticky="w",
        )
        ttk.Label(settings, text="總開始日期").grid(
            row=1,
            column=0,
            sticky="w",
            padx=(0, 8),
            pady=(10, 0),
        )
        self.project_start_picker = CalendarDatePicker(
            settings,
            self.initial_overall_start,
            on_change=self._on_project_date_changed,
        )
        self.project_start_picker.grid(row=1, column=1, sticky="w", pady=(10, 0))
        ttk.Label(settings, text="總結束日期").grid(
            row=1,
            column=2,
            sticky="w",
            padx=(20, 8),
            pady=(10, 0),
        )
        self.project_end_picker = CalendarDatePicker(
            settings,
            self.initial_overall_end,
            on_change=self._on_project_date_changed,
        )
        self.project_end_picker.grid(row=1, column=3, sticky="w", pady=(10, 0))
        ttk.Label(
            settings,
            text="點日期欄位或 ▼ 開啟整月月曆",
            style="Subtitle.TLabel",
        ).grid(row=1, column=4, columnspan=2, sticky="w", padx=(20, 0), pady=(10, 0))
        ttk.Label(settings, textvariable=self.range_text, style="Subtitle.TLabel").grid(
            row=2,
            column=0,
            columnspan=6,
            sticky="w",
            pady=(8, 0),
        )

        editor = ttk.LabelFrame(outer, text="新增／修改工作", style="Section.TLabelframe", padding=12)
        editor.pack(fill="x", pady=(12, 10))
        editor.columnconfigure(1, weight=1)

        ttk.Label(editor, text="工作項目").grid(row=0, column=0, sticky="w", padx=(0, 8))
        task_entry = ttk.Entry(editor, textvariable=self.task_name)
        task_entry.grid(row=0, column=1, sticky="ew")
        ttk.Button(editor, text="新增", command=self._add_task).grid(row=0, column=2, padx=(18, 4))
        ttk.Button(editor, text="更新選取", command=self._update_task).grid(row=0, column=3, padx=4)
        ttk.Button(editor, text="清除欄位", command=self._clear_editor).grid(row=0, column=4, padx=(4, 0))
        ttk.Label(editor, text="各項日期由上方總期間自動平均分配", style="Subtitle.TLabel").grid(
            row=1,
            column=1,
            columnspan=4,
            sticky="w",
            pady=(7, 0),
        )
        task_entry.bind("<Return>", lambda _event: self._add_task())

        # 先以 bottom 排版保留操作列空間，中央表格只使用剩餘高度。
        actions = ttk.Frame(outer)
        actions.pack(side="bottom", fill="x", pady=(12, 0))
        actions.columnconfigure(4, weight=1)
        ttk.Label(actions, textvariable=self.status_text, style="Subtitle.TLabel").grid(
            row=0,
            column=0,
            columnspan=6,
            sticky="ew",
            pady=(0, 6),
        )
        ttk.Button(actions, text="刪除選取", command=self._delete_task).grid(row=1, column=0, sticky="w")
        ttk.Button(actions, text="上移", command=lambda: self._move_task(-1)).grid(
            row=1,
            column=1,
            sticky="w",
            padx=(8, 0),
        )
        ttk.Button(actions, text="下移", command=lambda: self._move_task(1)).grid(
            row=1,
            column=2,
            sticky="w",
            padx=(8, 0),
        )
        ttk.Button(actions, text="重新載入範例", command=self._load_sample).grid(
            row=1,
            column=3,
            sticky="w",
            padx=(18, 0),
        )
        self.generate_button = ttk.Button(
            actions,
            text="產生 Excel",
            command=self._generate_excel,
            style="Primary.TButton",
        )
        self.generate_button.grid(row=1, column=5, sticky="e")

        table_frame = ttk.Frame(outer)
        table_frame.pack(fill="both", expand=True)
        columns = ("number", "task", "start", "end", "days")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="browse")
        self.tree.heading("number", text="#")
        self.tree.heading("task", text="工作項目")
        self.tree.heading("start", text="開始日期")
        self.tree.heading("end", text="結束日期")
        self.tree.heading("days", text="天數")
        self.tree.column("number", width=48, anchor="center", stretch=False)
        self.tree.column("task", width=430, anchor="w")
        self.tree.column("start", width=130, anchor="center", stretch=False)
        self.tree.column("end", width=130, anchor="center", stretch=False)
        self.tree.column("days", width=80, anchor="center", stretch=False)
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        self.tree.bind("<Double-1>", self._on_tree_select)

    def _on_scale_changed(self, _event=None) -> None:
        self.scale_help.set(SCALE_LABELS[self.scale.get()])

    def _get_project_range(self) -> tuple[date, date]:
        overall_start = self.project_start_picker.get_date()
        overall_end = self.project_end_picker.get_date()
        if overall_end < overall_start:
            raise ValueError("總結束日期不可早於總開始日期。")
        return overall_start, overall_end

    def _redistribute_tasks(self, select_index: int | None = None) -> bool:
        try:
            overall_start, overall_end = self._get_project_range()
        except ValueError as exc:
            self.range_text.set(f"日期設定錯誤：{exc}")
            self.status_text.set("請先修正總開始日期與總結束日期。")
            return False

        self.tasks = distribute_tasks_for_range(self.tasks, overall_start, overall_end)
        self._refresh_tree(select_index=select_index)
        return True

    def _on_project_date_changed(self, _selected_date: date | None = None) -> None:
        selected_index = self.selected_index
        if self._redistribute_tasks(select_index=selected_index):
            self.status_text.set("已依新的總期間自動平均分配各工作。")

    def _choose_color(self) -> None:
        _rgb, chosen = colorchooser.askcolor(
            color=self.average_color,
            title="選擇動態色塊顏色",
            parent=self,
        )
        if chosen:
            self.average_color = chosen
            self.color_button.configure(
                bg=chosen,
                activebackground=chosen,
                fg=contrast_text_color(chosen),
            )

    def _read_editor_task(self) -> Task:
        name = self.task_name.get().strip()
        if not name:
            raise ValueError("請輸入工作項目名稱。")
        overall_start, overall_end = self._get_project_range()
        return Task(name=name, start=overall_start, end=overall_end)

    def _add_task(self) -> None:
        try:
            task = self._read_editor_task()
        except ValueError as exc:
            messagebox.showerror("資料錯誤", str(exc), parent=self)
            return
        self.tasks.append(task)
        self._redistribute_tasks(select_index=len(self.tasks) - 1)
        self._clear_editor()
        self.status_text.set(f"已新增「{task.name}」，並重新平均分配總期間。")

    def _update_task(self) -> None:
        if self.selected_index is None:
            messagebox.showinfo("尚未選取", "請先在清單中選取要修改的工作。", parent=self)
            return
        try:
            task = self._read_editor_task()
        except ValueError as exc:
            messagebox.showerror("資料錯誤", str(exc), parent=self)
            return
        self.tasks[self.selected_index] = task
        self._redistribute_tasks(select_index=self.selected_index)
        self.status_text.set(f"已更新「{task.name}」，並重新平均分配總期間。")

    def _delete_task(self) -> None:
        if self.selected_index is None:
            messagebox.showinfo("尚未選取", "請先選取要刪除的工作。", parent=self)
            return
        task = self.tasks[self.selected_index]
        if not messagebox.askyesno("刪除工作", f"確定刪除「{task.name}」？", parent=self):
            return
        del self.tasks[self.selected_index]
        next_index = min(self.selected_index, len(self.tasks) - 1) if self.tasks else None
        self.selected_index = None
        self._redistribute_tasks(select_index=next_index)
        self._clear_editor()
        self.status_text.set(f"已刪除「{task.name}」，並重新平均分配總期間。")

    def _move_task(self, direction: int) -> None:
        if self.selected_index is None:
            messagebox.showinfo("尚未選取", "請先選取要移動的工作。", parent=self)
            return
        destination = self.selected_index + direction
        if not 0 <= destination < len(self.tasks):
            return
        self.tasks[self.selected_index], self.tasks[destination] = (
            self.tasks[destination],
            self.tasks[self.selected_index],
        )
        self.selected_index = destination
        self._redistribute_tasks(select_index=destination)

    def _load_sample(self) -> None:
        if self.tasks and not messagebox.askyesno(
            "重新載入範例",
            "目前清單會被 Planning～Follow up 範例取代，是否繼續？",
            parent=self,
        ):
            return
        self.tasks = sample_tasks()
        overall_start = self.tasks[0].start
        overall_end = self.tasks[-1].end
        self.project_start_picker.set_date(overall_start, notify=False)
        self.project_end_picker.set_date(overall_end, notify=False)
        self.selected_index = None
        self._redistribute_tasks()
        self._clear_editor()
        self.status_text.set("已重新載入範例排程。")

    def _on_tree_select(self, _event=None) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        index = int(selection[0])
        if not 0 <= index < len(self.tasks):
            return
        self.selected_index = index
        task = self.tasks[index]
        self.task_name.set(task.name)

    def _clear_editor(self) -> None:
        self.selected_index = None
        self.task_name.set("")
        for item in self.tree.selection():
            self.tree.selection_remove(item)

    def _refresh_tree(self, select_index: int | None = None) -> None:
        self.tree.delete(*self.tree.get_children())
        for index, task in enumerate(self.tasks):
            duration = (task.end - task.start).days
            self.tree.insert(
                "",
                "end",
                iid=str(index),
                values=(index + 1, task.name, format_date(task.start), format_date(task.end), duration),
            )
        try:
            overall_start, overall_end = self._get_project_range()
            self.range_text.set(
                f"甘特圖總範圍：{format_date(overall_start)} ～ "
                f"{format_date(overall_end)}（各項目自動平均分配）"
            )
        except ValueError as exc:
            self.range_text.set(f"日期設定錯誤：{exc}")
        self.selected_index = None
        if select_index is not None and 0 <= select_index < len(self.tasks):
            item_id = str(select_index)
            self.tree.selection_set(item_id)
            self.tree.focus(item_id)
            self.tree.see(item_id)
            self._on_tree_select()

    def _generate_excel(self) -> None:
        if self.selected_index is not None or self.task_name.get().strip():
            try:
                pending_task = self._read_editor_task()
            except ValueError as exc:
                messagebox.showerror("資料錯誤", str(exc), parent=self)
                return
            if self.selected_index is None:
                self.tasks.append(pending_task)
                committed_index = len(self.tasks) - 1
            else:
                self.tasks[self.selected_index] = pending_task
                committed_index = self.selected_index
            self._refresh_tree(select_index=committed_index)

        if not self.tasks:
            messagebox.showerror("沒有工作", "至少新增一項工作後才能產生 Excel。", parent=self)
            return
        try:
            overall_start, overall_end = self._get_project_range()
        except ValueError as exc:
            messagebox.showerror("日期錯誤", str(exc), parent=self)
            return
        self.tasks = distribute_tasks_for_range(self.tasks, overall_start, overall_end)
        self._refresh_tree()
        self.status_text.set("已依精確總日期平均分配各工作。")
        title = self.project_title.get().strip() or "Gantt Chart"
        output_value = filedialog.asksaveasfilename(
            parent=self,
            title="儲存 Excel 甘特圖",
            initialdir=str(APP_DIR),
            initialfile=DEFAULT_OUTPUT_NAME,
            defaultextension=".xlsx",
            filetypes=(("Excel 活頁簿", "*.xlsx"),),
        )
        if not output_value:
            return
        output_path = Path(output_value)
        try:
            create_excel_gantt(
                output_path=output_path,
                project_title=title,
                tasks=self.tasks,
                scale=self.scale.get(),
                average_color=self.average_color,
                overall_start=overall_start,
                overall_end=overall_end,
            )
            validate_generated_workbook(
                output_path,
                len(self.tasks),
                self.average_color,
                self.scale.get(),
            )
        except PermissionError:
            messagebox.showerror(
                "無法儲存",
                "檔案可能正在 Excel 中開啟。請關閉該檔案後再試一次。",
                parent=self,
            )
            return
        except (OSError, RuntimeError, ValueError) as exc:
            messagebox.showerror("產生失敗", str(exc), parent=self)
            return

        self.status_text.set(f"已產生：{output_path.name}")
        should_open = messagebox.askyesno(
            "產生完成",
            f"已建立原生 Excel 甘特圖：\n{output_path}\n\n是否立即開啟？",
            parent=self,
        )
        if should_open:
            try:
                os.startfile(output_path)  # type: ignore[attr-defined]
            except OSError as exc:
                messagebox.showwarning("無法開啟", f"檔案已建立，但無法自動開啟：\n{exc}", parent=self)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="用 GUI 建立可在 Excel 內直接修改的甘特圖。")
    parser.add_argument(
        "--sample-output",
        type=Path,
        help="不開啟 GUI，直接產生範例 Excel；供測試或快速建立範本。",
    )
    parser.add_argument("--scale", choices=tuple(SCALE_LABELS), default="週", help="範例的時間刻度。")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if args.sample_output:
        output_path = args.sample_output.resolve()
        tasks = sample_tasks()
        try:
            create_excel_gantt(output_path, "Gantt Chart", tasks, args.scale)
            validate_generated_workbook(output_path, len(tasks), expected_scale=args.scale)
        except (OSError, RuntimeError, ValueError) as exc:
            print(f"錯誤：{exc}", file=sys.stderr)
            return 1
        print(f"已建立：{output_path}")
        return 0

    app = ExcelGanttApp()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
