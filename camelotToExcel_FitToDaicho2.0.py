import os
import glob
import re
import shutil
import subprocess
import sys
import tkinter as tk
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from tkinter import filedialog, messagebox, ttk

import camelot
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTChar


EXPECTED_COLUMN_COUNT = 19
NAME_COLUMN_INDEX = 1
DEFAULT_COLUMNS = (
    "54.95,153.33,176.31,198.44,250.75,280.99,310.50,"
    "341.70,370.73,400.00,429.52,458.55,487.83,517.10,"
    "546.38,575.65,605.17,747.94"
)
DEFAULT_TABLE_AREAS = ""


def find_ghostscript_executable():
    """Find an installed Ghostscript console executable on Windows."""
    bundle_root = getattr(sys, "_MEIPASS", os.path.dirname(__file__))
    bundled_candidates = [
        os.path.join(
            bundle_root, "ghostscript", "bin", "gswin64c.exe"
        ),
        os.path.join(
            bundle_root, "ghostscript", "bin", "gswin32c.exe"
        ),
    ]
    for executable in bundled_candidates:
        if os.path.isfile(executable):
            return executable

    for command in ("gswin64c.exe", "gswin32c.exe", "gs.exe", "gs"):
        executable = shutil.which(command)
        if executable:
            return executable

    search_roots = [
        os.environ.get("ProgramFiles"),
        os.environ.get("ProgramFiles(x86)"),
    ]
    candidates = []
    for root in filter(None, search_roots):
        candidates.extend(
            glob.glob(os.path.join(root, "gs", "gs*", "bin", "gswin64c.exe"))
        )
        candidates.extend(
            glob.glob(os.path.join(root, "gs", "gs*", "bin", "gswin32c.exe"))
        )

    if candidates:
        return sorted(candidates, reverse=True)[0]
    return None


class InstalledGhostscriptBackend:
    """Camelot backend that calls an installed Ghostscript executable directly."""

    def __init__(self, resolution=300):
        self.resolution = resolution

    def convert(self, pdf_path, png_path):
        executable = find_ghostscript_executable()
        if not executable:
            raise OSError(
                "Ghostscriptの実行ファイル（gswin64c.exe）を検出できません。"
                "GhostscriptのbinフォルダをPATHへ登録するか、"
                "標準のProgram Files配下へインストールしてください。"
            )

        environment = os.environ.copy()
        executable_dir = os.path.dirname(executable)
        bundle_root = os.path.dirname(executable_dir)
        bundled_resource = os.path.join(bundle_root, "Resource")
        bundled_lib = os.path.join(bundle_root, "lib")
        if os.path.isdir(bundled_resource):
            environment["GS_LIB"] = os.pathsep.join(
                [bundled_resource, bundled_lib]
            )
            environment["PATH"] = (
                executable_dir
                + os.pathsep
                + environment.get("PATH", "")
            )

        command = [
            executable,
            "-q",
            "-dSAFER",
            "-dBATCH",
            "-dNOPAUSE",
            "-sDEVICE=png16m",
            f"-r{self.resolution}",
            f"-sOutputFile={png_path}",
            pdf_path,
        ]
        completed = subprocess.run(
            command,
            env=environment,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        if completed.returncode != 0:
            detail = (completed.stderr or completed.stdout).strip()
            raise OSError(
                f"Ghostscript変換に失敗しました（終了コード"
                f" {completed.returncode}）。\n{detail}"
            )


@dataclass
class ExtractedTable:
    df: pd.DataFrame
    page: int
    accuracy: float


@dataclass(frozen=True)
class PageInfo:
    branch: str = ""
    division: str = ""
    division_code: str = ""
    group: str = ""

    def missing_fields(self):
        missing = []
        if not self.branch:
            missing.append("支部")
        if not self.division:
            missing.append("分会")
        if not self.division_code:
            missing.append("分会コード")
        if not self.group:
            missing.append("班")
        return missing


@dataclass(frozen=True)
class YearProcessingOptions:
    enabled: bool = False
    fiscal_year: int | None = None
    unresolved_mode: str = "keep"


def parse_list(value):
    return [item.strip() for item in value.split(";") if item.strip()]


def parse_bool(value):
    return value.strip().lower() in {"true", "1", "yes", "on"}


def normalize_pdf_text(value):
    return (
        value.replace("\u3000", " ")
        .replace("（", "(")
        .replace("）", ")")
        .strip()
    )


def to_ascii_digits(value):
    table = str.maketrans("０１２３４５６７８９", "0123456789")
    return value.translate(table)


def two_digit_code(value):
    digits = to_ascii_digits(value).strip()
    if digits.isdigit():
        return digits.zfill(2)
    return digits


def iter_text_objects(layout_object):
    if isinstance(layout_object, LTChar):
        return
    if hasattr(layout_object, "get_text"):
        text = layout_object.get_text()
        if text:
            yield text
            return
    if hasattr(layout_object, "__iter__"):
        for child in layout_object:
            yield from iter_text_objects(child)


def extract_pdf_page_lines(pdf_path):
    pages = {}
    for page_number, page_layout in enumerate(extract_pages(pdf_path), start=1):
        lines = []
        for text in iter_text_objects(page_layout):
            for line in text.splitlines():
                line = normalize_pdf_text(line)
                if line:
                    lines.append(line)
        pages[page_number] = lines
    return pages


def parse_page_info(lines):
    text = normalize_pdf_text(" ".join(lines))
    branch = ""
    division = ""
    division_code = ""
    group = ""

    branch_match = re.search(r"([^\s,、。()]+?)支部", text)
    if branch_match:
        branch = branch_match.group(1).strip()

    group_match = re.search(r"([0-9０-９]{1,2})\s*班", text)
    if group_match:
        group = two_digit_code(group_match.group(1))

    division_code_match = re.search(
        r"([^\s,、。()]+?)分会\s*\(\s*([0-9０-９]{1,2})\s*\)",
        text,
    )
    if division_code_match:
        division = division_code_match.group(1).strip()
        if "支部" in division:
            division = division.rsplit("支部", 1)[1].strip()
        division_code = two_digit_code(division_code_match.group(2))
    else:
        division_match = re.search(r"([^\s,、。()]+?)分会", text)
        if division_match:
            division = division_match.group(1).strip()
            if "支部" in division:
                division = division.rsplit("支部", 1)[1].strip()
            code_end = group_match.start() if group_match else len(text)
            code_area = text[division_match.end():code_end]
            code_match = re.search(r"([0-9０-９]{1,2})", code_area)
            if code_match:
                division_code = two_digit_code(code_match.group(1))

    return PageInfo(
        branch=branch,
        division=division,
        division_code=division_code,
        group=group,
    )


def extract_page_infos(pdf_path):
    return {
        page_number: parse_page_info(lines)
        for page_number, lines in extract_pdf_page_lines(pdf_path).items()
    }


def export_page_text_debug(pdf_path, output_folder):
    os.makedirs(output_folder, exist_ok=True)
    base = os.path.splitext(os.path.basename(pdf_path))[0]
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = os.path.join(
        output_folder, f"{base}_{timestamp}_page_text_debug.txt"
    )
    page_lines = extract_pdf_page_lines(pdf_path)
    with open(output_path, "w", encoding="utf-8") as output_file:
        for page_number in sorted(page_lines):
            output_file.write(f"===== page {page_number} =====\n")
            for line in page_lines[page_number]:
                output_file.write(line + "\n")
            output_file.write("\n")
    return output_path


def validate_coordinate_string(value, expected_count=None):
    parts = [part.strip() for part in value.split(",")]
    if expected_count is not None and len(parts) != expected_count:
        raise ValueError(
            f"{expected_count}個の数値が必要です。現在は{len(parts)}個です。"
        )
    for part in parts:
        float(part)


def iter_chars(layout_object):
    if isinstance(layout_object, LTChar):
        yield layout_object
        return
    if hasattr(layout_object, "__iter__"):
        for child in layout_object:
            yield from iter_chars(child)


def find_column(x_center, boundaries):
    for index in range(len(boundaries) - 1):
        if boundaries[index] <= x_center < boundaries[index + 1]:
            return index
    if x_center == boundaries[-1]:
        return len(boundaries) - 2
    return None


def split_textline_by_column(textline, boundaries):
    """Split a PDF text line at actual column boundaries, character by character."""
    fragments = []
    current_column = None
    current_chars = []
    current_x_values = []
    previous_x1 = None

    def flush():
        if current_column is None or not current_chars:
            return
        text = "".join(current_chars).strip()
        if text:
            fragments.append(
                {
                    "column": current_column,
                    "text": text,
                    "x": min(current_x_values),
                    "y": (textline.y0 + textline.y1) / 2,
                }
            )

    for char in iter_chars(textline):
        text = char.get_text()
        if not text:
            continue
        x_center = (char.x0 + char.x1) / 2
        column = find_column(x_center, boundaries)
        if column is None:
            continue
        if current_column is not None and column != current_column:
            flush()
            current_chars = []
            current_x_values = []
            previous_x1 = None
        current_column = column
        if previous_x1 is not None and char.x0 - previous_x1 > 1.0:
            current_chars.append(" ")
        current_chars.append(text)
        current_x_values.append(char.x0)
        previous_x1 = char.x1

    flush()
    return fragments


def join_fragments(fragments):
    if not fragments:
        return ""

    lines = defaultdict(list)
    for fragment in fragments:
        y_key = round(fragment["y"], 1)
        lines[y_key].append(fragment)

    output_lines = []
    for y_key in sorted(lines, reverse=True):
        same_line = sorted(lines[y_key], key=lambda item: item["x"])
        output_lines.append(" ".join(item["text"] for item in same_line))
    return "\n".join(output_lines)


def name_fragments_to_three_rows(fragments, y_bottom, y_top):
    bands = [[], [], []]
    height = max(y_top - y_bottom, 0.001)

    for fragment in fragments:
        relative_position = (fragment["y"] - y_bottom) / height
        if relative_position >= 2 / 3:
            band_index = 0
        elif relative_position >= 1 / 3:
            band_index = 1
        else:
            band_index = 2
        bands[band_index].append(fragment)

    output = []
    for band in bands:
        ordered = sorted(band, key=lambda item: item["x"])
        output.append(" ".join(item["text"] for item in ordered).strip())
    return output


def get_column_boundaries(table, configured_separators):
    left = min(column[0] for column in table.cols)
    right = max(column[1] for column in table.cols)
    detected = [left] + [column[1] for column in table.cols]

    if len(configured_separators) == EXPECTED_COLUMN_COUNT - 1:
        configured = [left] + configured_separators + [right]
        if all(
            configured[index] < configured[index + 1]
            for index in range(len(configured) - 1)
        ):
            return configured
    return detected


def is_header_row(table, row_index):
    values = [cell.text.replace("\n", "") for cell in table.cells[row_index]]
    return "人数" in values and any("氏" in value for value in values)


def extract_row_fragments(table, row_index, boundaries):
    row_cells = table.cells[row_index]
    y_bottom = min(cell.y1 for cell in row_cells)
    y_top = max(cell.y2 for cell in row_cells)
    by_column = defaultdict(list)

    for textline in table.textlines:
        line_y = (textline.y0 + textline.y1) / 2
        if not (y_bottom <= line_y <= y_top):
            continue
        for fragment in split_textline_by_column(textline, boundaries):
            by_column[fragment["column"]].append(fragment)

    return by_column, y_bottom, y_top


def expand_member_row(table, row_index, boundaries):
    fragments, y_bottom, y_top = extract_row_fragments(
        table, row_index, boundaries
    )
    name_rows = name_fragments_to_three_rows(
        fragments.get(NAME_COLUMN_INDEX, []), y_bottom, y_top
    )
    output_rows = [[""] * EXPECTED_COLUMN_COUNT for _ in range(3)]

    for name_row_index, value in enumerate(name_rows):
        output_rows[name_row_index][NAME_COLUMN_INDEX] = value

    for column_index in range(EXPECTED_COLUMN_COUNT):
        if column_index == NAME_COLUMN_INDEX:
            continue
        output_rows[1][column_index] = join_fragments(
            fragments.get(column_index, [])
        )

    return output_rows


def extract_fragments_in_band(textlines, y_bottom, y_top, boundaries):
    by_column = defaultdict(list)
    for textline in textlines:
        line_y = (textline.y0 + textline.y1) / 2
        if not (y_bottom <= line_y < y_top):
            continue
        for fragment in split_textline_by_column(textline, boundaries):
            by_column[fragment["column"]].append(fragment)
    return by_column


def find_member_row_centers(textlines, boundaries, header_bottom):
    """Use the numbered first column as anchors for all ledger rows."""
    numbered_rows = {}
    for textline in textlines:
        line_y = (textline.y0 + textline.y1) / 2
        if line_y >= header_bottom:
            continue
        fragments = split_textline_by_column(textline, boundaries)
        if len(fragments) != 1 or fragments[0]["column"] != 0:
            continue
        value = fragments[0]["text"].replace(" ", "").strip()
        if not value.isdigit():
            continue
        number = int(value)
        if 1 <= number <= 99:
            numbered_rows[number] = line_y
    anchors = sorted(numbered_rows.items())
    if len(anchors) < 2:
        return sorted(anchors, key=lambda item: item[1], reverse=True)

    per_row_gaps = []
    for index in range(len(anchors) - 1):
        first_number, first_y = anchors[index]
        second_number, second_y = anchors[index + 1]
        number_gap = second_number - first_number
        if number_gap > 0:
            per_row_gaps.append((first_y - second_y) / number_gap)

    if not per_row_gaps:
        return sorted(anchors, key=lambda item: item[1], reverse=True)

    typical_gap = sorted(per_row_gaps)[len(per_row_gaps) // 2]
    first_number, first_y = anchors[0]
    last_number = anchors[-1][0]
    completed_rows = []
    for number in range(first_number, last_number + 1):
        center = numbered_rows.get(
            number, first_y - (number - first_number) * typical_gap
        )
        completed_rows.append((number, center))
    return sorted(completed_rows, key=lambda item: item[1], reverse=True)


def expand_member_band(textlines, y_bottom, y_top, boundaries):
    fragments = extract_fragments_in_band(
        textlines, y_bottom, y_top, boundaries
    )
    name_rows = name_fragments_to_three_rows(
        fragments.get(NAME_COLUMN_INDEX, []), y_bottom, y_top
    )
    output_rows = [[""] * EXPECTED_COLUMN_COUNT for _ in range(3)]

    for name_row_index, value in enumerate(name_rows):
        output_rows[name_row_index][NAME_COLUMN_INDEX] = value

    for column_index in range(EXPECTED_COLUMN_COUNT):
        if column_index == NAME_COLUMN_INDEX:
            continue
        output_rows[1][column_index] = join_fragments(
            fragments.get(column_index, [])
        )
    return output_rows


def normalize_page_from_numbered_grid(page_tables, configured_separators):
    """Reconstruct rows even when lattice misses alternating/broken row boxes."""
    page_tables = sorted(page_tables, key=lambda table: table._bbox[3], reverse=True)
    source_table = page_tables[0]
    boundaries = get_column_boundaries(source_table, configured_separators)
    header_row = None
    header_bottom = None

    for table in page_tables:
        if table.shape[1] != EXPECTED_COLUMN_COUNT:
            continue
        for row_index in range(table.shape[0]):
            if is_header_row(table, row_index):
                header_row = [
                    cell.text.strip() for cell in table.cells[row_index]
                ]
                header_bottom = min(
                    cell.y1 for cell in table.cells[row_index]
                )
                break
        if header_row is not None:
            break

    if header_row is None:
        return None

    numbered_rows = find_member_row_centers(
        source_table.textlines, boundaries, header_bottom
    )
    if not numbered_rows:
        return None

    centers = [center for _, center in numbered_rows]
    if len(centers) >= 2:
        gaps = [
            centers[index] - centers[index + 1]
            for index in range(len(centers) - 1)
        ]
        typical_gap = sorted(gaps)[len(gaps) // 2]
    else:
        typical_gap = 29.5

    output_rows = [header_row]
    for index, (_, center) in enumerate(numbered_rows):
        y_top = (
            (centers[index - 1] + center) / 2
            if index > 0
            else min(header_bottom, center + typical_gap / 2)
        )
        y_bottom = (
            (center + centers[index + 1]) / 2
            if index + 1 < len(centers)
            else center - typical_gap / 2
        )
        output_rows.extend(
            expand_member_band(
                source_table.textlines, y_bottom, y_top, boundaries
            )
        )

    accuracy_values = [
        table.parsing_report["accuracy"]
        for table in page_tables
        if table.shape[1] == EXPECTED_COLUMN_COUNT
    ]
    accuracy = (
        round(sum(accuracy_values) / len(accuracy_values), 2)
        if accuracy_values
        else 0.0
    )
    return ExtractedTable(
        df=pd.DataFrame(output_rows),
        page=int(source_table.page),
        accuracy=accuracy,
    )


def normalize_page_tables(page_tables, configured_separators):
    numbered_grid = normalize_page_from_numbered_grid(
        page_tables, configured_separators
    )
    if numbered_grid is not None:
        return numbered_grid

    page_tables = sorted(page_tables, key=lambda table: table._bbox[3], reverse=True)
    output_rows = []
    header_found = False
    accuracy_values = []

    for table in page_tables:
        if table.shape[1] != EXPECTED_COLUMN_COUNT:
            continue
        accuracy_values.append(table.parsing_report["accuracy"])
        boundaries = get_column_boundaries(table, configured_separators)

        for row_index in range(table.shape[0]):
            if is_header_row(table, row_index):
                if not header_found:
                    header = [
                        cell.text.strip()
                        for cell in table.cells[row_index]
                    ]
                    output_rows.append(header)
                    header_found = True
                continue

            fragments, _, _ = extract_row_fragments(
                table, row_index, boundaries
            )
            if not any(fragments.values()):
                continue
            output_rows.extend(expand_member_row(table, row_index, boundaries))

    if not output_rows:
        return None

    accuracy = (
        round(sum(accuracy_values) / len(accuracy_values), 2)
        if accuracy_values
        else 0.0
    )
    return ExtractedTable(
        df=pd.DataFrame(output_rows),
        page=int(page_tables[0].page),
        accuracy=accuracy,
    )


def extract_daicho(pdf_path, pages, options, configured_separators):
    options = dict(options)
    resolution = int(options.pop("resolution", 300))
    options["backend"] = InstalledGhostscriptBackend(resolution=resolution)
    options["use_fallback"] = False
    raw_tables = camelot.read_pdf(
        pdf_path,
        flavor="lattice",
        pages=pages,
        suppress_stdout=False,
        debug=True,
        **options,
    )

    tables_by_page = defaultdict(list)
    for table in raw_tables:
        if table.shape[1] == EXPECTED_COLUMN_COUNT:
            tables_by_page[int(table.page)].append(table)

    extracted = []
    for page_number in sorted(tables_by_page):
        normalized = normalize_page_tables(
            tables_by_page[page_number], configured_separators
        )
        if normalized is not None:
            extracted.append(normalized)
    return extracted


def clean_cell_value(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def current_fiscal_year(today=None):
    today = today or date.today()
    return today.year if today.month >= 4 else today.year - 1


def parse_month_label(value):
    text = to_ascii_digits(clean_cell_value(value))
    if not text.isdigit():
        return None
    month = int(text)
    if 1 <= month <= 12:
        return month
    return None


def base_year_for_month(fiscal_year, month):
    return fiscal_year if month >= 4 else fiscal_year + 1


def next_business_day_if_weekend(value):
    while value.weekday() >= 5:
        value += timedelta(days=1)
    return value


def monthly_header_date(fiscal_year, month):
    value = date(base_year_for_month(fiscal_year, month), month, 25)
    return next_business_day_if_weekend(value)


def normalize_date_text(value):
    return (
        to_ascii_digits(clean_cell_value(value))
        .replace("／", "/")
        .replace("−", "-")
        .replace("ー", "-")
    )


def extract_month_day(value):
    text = normalize_date_text(value)
    match = re.search(
        r"(?<!\d)(0?[1-9]|1[0-2])\s*/\s*(0?[1-9]|[12]\d|3[01])(?!\d)",
        text,
    )
    if not match:
        return None
    month = int(match.group(1))
    day = int(match.group(2))
    return month, day


def make_payment_date(fiscal_year, month, day, previous_date=None):
    year = base_year_for_month(fiscal_year, month)
    while True:
        value = date(year, month, day)
        if previous_date is None or value >= previous_date:
            return value
        year += 1


def format_date_for_text(value):
    if isinstance(value, date):
        return f"{value.year}/{value.month}/{value.day}"
    return value


def detect_month_columns(header):
    expected = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]
    columns = []
    for column_index, value in enumerate(header):
        month = parse_month_label(value)
        if month in expected:
            columns.append((column_index, month))
    return columns


def resolve_unreadable_payment_text(value, mode):
    original = clean_cell_value(value)
    if not original:
        return original, None
    if mode == "keep":
        return original, None
    if mode == "circle":
        marks = []
        for char in original:
            if char in {"〇", "○"}:
                marks.append("〇")
            elif char in {"-", "－", "−", "‐", "‑"}:
                marks.append("-")
        resolved = "".join(marks)
    elif mode == "clear":
        resolved = ""
    else:
        resolved = original
    if resolved == original:
        return resolved, None
    return resolved, (original, resolved)


def apply_year_processing(df, options, source_name=""):
    if not options.enabled or df.empty:
        return df, [], []

    fiscal_year = options.fiscal_year or current_fiscal_year()
    processed = df.copy()
    warnings = []
    header = processed.iloc[0].tolist()
    month_columns = detect_month_columns(header)
    expected_months = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]
    detected_months = [month for _, month in month_columns]
    if detected_months != expected_months:
        warnings.append(
            "月列が想定順序と一致しません。検出: "
            + ",".join(str(month) for month in detected_months)
        )

    for column_index, month in month_columns:
        processed.iat[0, column_index] = monthly_header_date(
            fiscal_year, month
        )

    for row_index in range(1, len(processed.index)):
        previous_date = None
        for column_index, month in month_columns:
            raw_value = clean_cell_value(processed.iat[row_index, column_index])
            if not raw_value:
                continue
            month_day = extract_month_day(raw_value)
            if month_day is None:
                resolved, removed = resolve_unreadable_payment_text(
                    raw_value, options.unresolved_mode
                )
                processed.iat[row_index, column_index] = resolved
                if removed is not None:
                    original, after = removed
                    after_label = after if after else "空白"
                    warnings.append(
                        f"{source_name} row {row_index + 1} / "
                        f"{month}月列: 「{original}」を削除しました。"
                        f"削除後: {after_label}"
                    )
                continue
            raw_month, raw_day = month_day
            try:
                payment_date = make_payment_date(
                    fiscal_year, raw_month, raw_day, previous_date
                )
            except ValueError:
                warnings.append(
                    f"{source_name} row {row_index + 1} / "
                    f"{month}月列: 日付として扱えません: {raw_value}"
                )
                continue
            processed.iat[row_index, column_index] = payment_date
            previous_date = payment_date

    return processed, warnings, [column for column, _ in month_columns]


def format_dates_for_text_output(df):
    return df.apply(lambda column: column.map(format_date_for_text))


def build_one_sheet_table(table, page_info):
    if table.df.empty:
        return table.df.copy()

    source_rows = table.df.fillna("").values.tolist()
    header = [
        "支部",
        "分会",
        "分会コード",
        "班",
        "組合員番号",
    ] + [clean_cell_value(value) for value in source_rows[0]]
    output_rows = [header]

    for row_index in range(1, len(source_rows), 3):
        if row_index + 1 >= len(source_rows):
            break
        upper_row = source_rows[row_index]
        middle_row = source_rows[row_index + 1]
        member_number = clean_cell_value(upper_row[NAME_COLUMN_INDEX])
        output_rows.append(
            [
                page_info.branch,
                page_info.division,
                page_info.division_code,
                page_info.group,
                member_number,
            ]
            + [clean_cell_value(value) for value in middle_row]
        )

    return pd.DataFrame(output_rows)


def combine_one_sheet_tables(tables, page_infos):
    combined_rows = []
    warnings = []
    for table in tables:
        page_info = page_infos.get(table.page, PageInfo())
        missing = page_info.missing_fields()
        if missing:
            warnings.append(
                f"page {table.page}: "
                + "、".join(missing)
                + "を取得できませんでした。"
            )
        one_sheet = build_one_sheet_table(table, page_info)
        if one_sheet.empty:
            continue
        rows = one_sheet.values.tolist()
        if not combined_rows:
            combined_rows.extend(rows)
        else:
            combined_rows.extend(rows[1:])

    if not combined_rows:
        return pd.DataFrame(), warnings
    return pd.DataFrame(combined_rows), warnings


def export_tables(
    tables, pdf_path, output_folder, output_format, year_options
):
    os.makedirs(output_folder, exist_ok=True)
    base = os.path.splitext(os.path.basename(pdf_path))[0]
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_base = os.path.join(output_folder, f"{base}_{timestamp}")
    page_infos = extract_page_infos(pdf_path)
    combined_df, warnings = combine_one_sheet_tables(tables, page_infos)
    source_name = os.path.basename(pdf_path)
    combined_df, year_warnings, year_month_columns = apply_year_processing(
        combined_df, year_options, source_name=source_name
    )
    warnings.extend(year_warnings)

    if output_format == "excel":
        output_path = output_base + ".xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "all_pages"
        for row in dataframe_to_rows(
            combined_df, index=False, header=False
        ):
            sheet.append(row)
        for row in sheet.iter_rows():
            for cell in row:
                cell.alignment = Alignment(
                    vertical="center", wrap_text=True
                )
                if year_options.enabled and cell.column - 1 in year_month_columns:
                    if cell.row == 1 and isinstance(cell.value, date):
                        cell.number_format = "m"
                    elif isinstance(cell.value, date):
                        cell.number_format = "m/d"
        workbook.save(output_path)
        return [output_path], warnings

    extension = "md" if output_format == "markdown" else output_format
    output_path = f"{output_base}.{extension}"
    text_df = format_dates_for_text_output(combined_df)
    if output_format == "csv":
        text_df.to_csv(
            output_path,
            index=False,
            header=False,
            encoding="utf-8-sig",
        )
    elif output_format == "json":
        text_df.to_json(
            output_path,
            orient="values",
            force_ascii=False,
            indent=2,
        )
    elif output_format == "html":
        text_df.to_html(output_path, index=False, header=False)
    elif output_format == "markdown":
        with open(output_path, "w", encoding="utf-8") as output_file:
            output_file.write(
                text_df.to_markdown(index=False, headers=[])
            )
    return [output_path], warnings


def run_camelot():
    pdf_paths = [
        path.strip() for path in pdf_var.get().split(";") if path.strip()
    ]
    if not pdf_paths:
        messagebox.showerror("入力エラー", "PDFファイルを選択してください。")
        return

    try:
        separators = [
            float(value.strip()) for value in columns_var.get().split(",")
        ]
        if len(separators) != EXPECTED_COLUMN_COUNT - 1:
            raise ValueError(
                "columnsには19列を分ける18個の座標が必要です。"
            )

        table_areas = parse_list(table_areas_var.get())
        for area in table_areas:
            validate_coordinate_string(area, expected_count=4)

        options = {
            "line_scale": int(line_scale_var.get()),
            "line_tol": int(line_tol_var.get()),
            "joint_tol": int(joint_tol_var.get()),
            "threshold_blocksize": int(threshold_blocksize_var.get()),
            "threshold_constant": int(threshold_constant_var.get()),
            "iterations": int(iterations_var.get()),
            "resolution": int(resolution_var.get()),
            "process_background": parse_bool(process_background_var.get()),
            "split_text": False,
            "flag_size": False,
        }
        if table_areas:
            options["table_areas"] = table_areas

        year_options = YearProcessingOptions()
        if year_processing_var.get():
            fiscal_year_text = fiscal_year_var.get().strip()
            if fiscal_year_text:
                if not re.fullmatch(r"\d{4}", fiscal_year_text):
                    raise ValueError("年度は西暦4桁で入力してください。")
                fiscal_year = int(fiscal_year_text)
            else:
                fiscal_year = current_fiscal_year()
            unresolved_mode_labels = {
                "すべて残す": "keep",
                "〇・-だけ残す": "circle",
                "すべて消す": "clear",
            }
            year_options = YearProcessingOptions(
                enabled=True,
                fiscal_year=fiscal_year,
                unresolved_mode=unresolved_mode_labels[
                    unreadable_payment_var.get()
                ],
            )
    except ValueError as exc:
        messagebox.showerror("設定エラー", str(exc))
        return

    output_folder = output_folder_var.get().strip()
    if not output_folder:
        output_folder = os.path.join(
            os.path.dirname(pdf_paths[0]), "result"
        )

    results = []
    warnings = []
    for pdf_path in pdf_paths:
        try:
            debug_path = None
            if page_text_debug_var.get():
                debug_path = export_page_text_debug(pdf_path, output_folder)
            tables = extract_daicho(
                pdf_path,
                pages_var.get().strip() or "all",
                options,
                separators,
            )
            if not tables:
                warnings.append(
                    f"{os.path.basename(pdf_path)}: "
                    "19列の台帳表を検出できませんでした。"
                )
                continue

            output_paths, export_warnings = export_tables(
                tables,
                pdf_path,
                output_folder,
                output_format_var.get(),
                year_options,
            )
            warnings.extend(
                f"{os.path.basename(pdf_path)}: {warning}"
                for warning in export_warnings
            )
            for table in tables:
                results.append(
                    f"{os.path.basename(pdf_path)} page {table.page}: "
                    f"{table.df.shape[0]}行 x {table.df.shape[1]}列 / "
                    f"元表平均精度 {table.accuracy}%"
                )
            results.append(
                f"{os.path.basename(pdf_path)}: "
                f"{len(output_paths)}ファイルにまとめて出力しました。"
            )
            if debug_path:
                results.append(
                    f"{os.path.basename(pdf_path)}: "
                    f"ページ本文確認用ログ {debug_path}"
                )
        except Exception as exc:
            warnings.append(
                f"{os.path.basename(pdf_path)}: "
                f"{type(exc).__name__}: {exc}"
            )

    summary = "\n".join(results) if results else "出力できた表はありません。"
    if warnings:
        summary += "\n\n警告:\n" + "\n".join(warnings)
    summary += f"\n\n出力先:\n{output_folder}"
    messagebox.showinfo("処理結果", summary)


def select_pdf_files():
    paths = filedialog.askopenfilenames(
        title="台帳PDFを選択",
        filetypes=[("PDFファイル", "*.pdf")],
    )
    if paths:
        pdf_var.set("; ".join(paths))
        file_count_var.set(f"{len(paths)}ファイル選択済み")


def select_output_folder():
    path = filedialog.askdirectory(title="出力フォルダを選択")
    if path:
        output_folder_var.set(path)


def reset_optimized_values():
    table_areas_var.set(DEFAULT_TABLE_AREAS)
    columns_var.set(DEFAULT_COLUMNS)
    line_scale_var.set("25")
    line_tol_var.set("2")
    joint_tol_var.set("2")
    threshold_blocksize_var.set("15")
    threshold_constant_var.set("-2")
    iterations_var.set("0")
    resolution_var.set("300")
    process_background_var.set("False")


if len(sys.argv) >= 3 and sys.argv[1] == "--self-test-pdf":
    self_test_options = {
        "line_scale": 25,
        "line_tol": 2,
        "joint_tol": 2,
        "threshold_blocksize": 15,
        "threshold_constant": -2,
        "iterations": 0,
        "resolution": 300,
        "process_background": False,
        "split_text": False,
        "flag_size": False,
    }
    self_test_separators = [
        float(value) for value in DEFAULT_COLUMNS.split(",")
    ]
    self_test_tables = extract_daicho(
        sys.argv[2], "1", self_test_options, self_test_separators
    )
    if not self_test_tables:
        raise SystemExit(2)
    if self_test_tables[0].df.shape[1] != EXPECTED_COLUMN_COUNT:
        raise SystemExit(3)
    raise SystemExit(0)


root = tk.Tk()
root.title("Camelot 台帳19列抽出ツール 2.0")
root.geometry("980x830")
root.minsize(820, 700)

main = ttk.Frame(root, padding=12)
main.pack(fill="both", expand=True)
main.columnconfigure(1, weight=1)

pdf_var = tk.StringVar()
file_count_var = tk.StringVar(value="未選択")
pages_var = tk.StringVar(value="all")
output_folder_var = tk.StringVar()
output_format_var = tk.StringVar(value="excel")
table_areas_var = tk.StringVar(value=DEFAULT_TABLE_AREAS)
columns_var = tk.StringVar(value=DEFAULT_COLUMNS)
line_scale_var = tk.StringVar(value="25")
line_tol_var = tk.StringVar(value="2")
joint_tol_var = tk.StringVar(value="2")
threshold_blocksize_var = tk.StringVar(value="15")
threshold_constant_var = tk.StringVar(value="-2")
iterations_var = tk.StringVar(value="0")
resolution_var = tk.StringVar(value="300")
process_background_var = tk.StringVar(value="False")
page_text_debug_var = tk.BooleanVar(value=False)
year_processing_var = tk.BooleanVar(value=False)
fiscal_year_var = tk.StringVar()
unreadable_payment_var = tk.StringVar(value="すべて残す")

row = 0
ttk.Label(main, text="PDFファイル").grid(
    row=row, column=0, sticky="w", pady=4
)
ttk.Entry(main, textvariable=pdf_var).grid(
    row=row, column=1, sticky="ew", padx=8, pady=4
)
ttk.Button(main, text="参照", command=select_pdf_files).grid(
    row=row, column=2, pady=4
)

row += 1
ttk.Label(main, textvariable=file_count_var, foreground="gray").grid(
    row=row, column=1, sticky="w", padx=8
)

row += 1
ttk.Label(main, text="ページ").grid(
    row=row, column=0, sticky="w", pady=4
)
ttk.Entry(main, textvariable=pages_var, width=20).grid(
    row=row, column=1, sticky="w", padx=8, pady=4
)
ttk.Label(main, text="例: all / 1 / 1,3 / 1-5").grid(
    row=row, column=1, sticky="w", padx=(180, 0)
)

row += 1
ttk.Separator(main).grid(
    row=row, column=0, columnspan=3, sticky="ew", pady=10
)

row += 1
ttk.Label(
    main, text="台帳2.0抽出設定", font=("", 11, "bold")
).grid(row=row, column=0, columnspan=3, sticky="w")

row += 1
ttk.Label(main, text="抽出モード").grid(
    row=row, column=0, sticky="w", pady=4
)
ttk.Label(main, text="lattice（台帳2.0用に固定）").grid(
    row=row, column=1, sticky="w", padx=8
)

row += 1
ttk.Label(main, text="table_areas").grid(
    row=row, column=0, sticky="w", pady=4
)
ttk.Entry(main, textvariable=table_areas_var).grid(
    row=row, column=1, columnspan=2, sticky="ew", padx=8, pady=4
)

row += 1
ttk.Label(main, text="columns").grid(
    row=row, column=0, sticky="w", pady=4
)
ttk.Entry(main, textvariable=columns_var).grid(
    row=row, column=1, columnspan=2, sticky="ew", padx=8, pady=4
)

row += 1
ttk.Label(
    main,
    text=(
        "table_areasは通常空欄を推奨します。行数を固定せず、"
        "検出された19列表をすべて処理します。\n"
        "columnsは文字を19列へ再配置する18本の境界座標です。"
    ),
    foreground="gray",
).grid(row=row, column=1, columnspan=2, sticky="w", padx=8)

row += 1
option_frame = ttk.Frame(main)
option_frame.grid(
    row=row, column=0, columnspan=3, sticky="ew", pady=10
)
option_definitions = [
    ("line_scale", line_scale_var),
    ("line_tol", line_tol_var),
    ("joint_tol", joint_tol_var),
    ("threshold_blocksize", threshold_blocksize_var),
    ("threshold_constant", threshold_constant_var),
    ("iterations", iterations_var),
    ("resolution", resolution_var),
    ("process_background", process_background_var),
]
for index, (label, variable) in enumerate(option_definitions):
    option_row = index // 4
    option_column = (index % 4) * 2
    ttk.Label(option_frame, text=label).grid(
        row=option_row,
        column=option_column,
        sticky="w",
        padx=(0, 4),
        pady=3,
    )
    ttk.Entry(option_frame, textvariable=variable, width=10).grid(
        row=option_row,
        column=option_column + 1,
        sticky="w",
        padx=(0, 16),
        pady=3,
    )

row += 1
ttk.Button(
    main,
    text="最適化済み初期値に戻す",
    command=reset_optimized_values,
).grid(row=row, column=1, sticky="w", padx=8, pady=8)

row += 1
ttk.Checkbutton(
    main,
    text="ページ本文確認用ログを出力する",
    variable=page_text_debug_var,
).grid(row=row, column=1, sticky="w", padx=8, pady=4)

row += 1
year_frame = ttk.Frame(main)
year_frame.grid(row=row, column=1, columnspan=2, sticky="w", padx=8, pady=4)
ttk.Checkbutton(
    year_frame,
    text="年度処理を行う",
    variable=year_processing_var,
).pack(side="left", padx=(0, 14))
ttk.Label(year_frame, text="年度").pack(side="left", padx=(0, 4))
ttk.Entry(year_frame, textvariable=fiscal_year_var, width=8).pack(
    side="left", padx=(0, 14)
)
ttk.Label(year_frame, text="日付なし文字").pack(side="left", padx=(0, 4))
ttk.Combobox(
    year_frame,
    textvariable=unreadable_payment_var,
    values=("すべて残す", "〇・-だけ残す", "すべて消す"),
    width=12,
    state="readonly",
).pack(side="left")

row += 1
ttk.Label(
    main,
    text=(
        "出力規則: 会員1名を中央段の1行にまとめます。"
        "支部・分会・分会コード・班・組合員番号を先頭列へ追加します。"
    ),
    foreground="#245a8d",
).grid(row=row, column=0, columnspan=3, sticky="w", pady=5)

row += 1
ttk.Separator(main).grid(
    row=row, column=0, columnspan=3, sticky="ew", pady=10
)

row += 1
ttk.Label(main, text="出力形式").grid(
    row=row, column=0, sticky="w", pady=4
)
format_frame = ttk.Frame(main)
format_frame.grid(row=row, column=1, columnspan=2, sticky="w", padx=8)
for output_format in ["excel", "csv", "json", "html", "markdown"]:
    ttk.Radiobutton(
        format_frame,
        text=output_format,
        variable=output_format_var,
        value=output_format,
    ).pack(side="left", padx=(0, 12))

row += 1
ttk.Label(main, text="出力フォルダ").grid(
    row=row, column=0, sticky="w", pady=4
)
ttk.Entry(main, textvariable=output_folder_var).grid(
    row=row, column=1, sticky="ew", padx=8, pady=4
)
ttk.Button(main, text="参照", command=select_output_folder).grid(
    row=row, column=2, pady=4
)

row += 1
ttk.Label(
    main,
    text="未指定の場合は、最初のPDFと同じ場所の result フォルダへ出力します。",
    foreground="gray",
).grid(row=row, column=1, columnspan=2, sticky="w", padx=8)

row += 1
ttk.Button(
    main,
    text="PDFから1シート形式で抽出",
    command=run_camelot,
).grid(
    row=row,
    column=0,
    columnspan=3,
    pady=20,
    ipadx=25,
    ipady=8,
)

root.mainloop()
