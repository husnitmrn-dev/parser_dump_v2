"""Parse Huawei MML text dumps into an Excel workbook."""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Sequence

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


COMMAND_RE = re.compile(
    r"^\s*(?:%%/\*[^*]+\*/)?\s*([A-Z][A-Z0-9_ ]*:\s*;)(?:%%)?\s*$"
)
SEPARATOR_RE = re.compile(r"^\s*-{10,}\s*$")
RESULT_RE = re.compile(r"^\s*\(Number of results\s*=")
TABLE_END_RE = re.compile(r"^\s*(?:To be continued|---+\s*END\b)")


@dataclass
class MMLTable:
    command: str
    ne_name: str
    table_type: str
    headers: list[str]
    rows: list[list[str]]


def _split_columns(line: str) -> list[str]:
    return [value.strip() for value in re.split(r"\s{2,}", line.strip())]


def parse_mml_text(text: str) -> list[MMLTable]:
    """Extract tables from an MML dump, including paginated report blocks."""
    lines = text.splitlines()
    tables: list[MMLTable] = []
    index = 0

    while index < len(lines):
        command_match = COMMAND_RE.match(lines[index])
        if command_match is None:
            index += 1
            continue

        command = re.sub(r":\s*;$", "", command_match.group(1)).strip()

        ne_name = ""
        for search_index in range(index + 1, min(index + 10, len(lines))):
            current = lines[search_index].strip()
            if "#MC" in current and not current.startswith("+++"):
                ne_name = current
                break
        if not ne_name:
            for search_index in range(index - 1, max(index - 10, -1), -1):
                current = lines[search_index].strip()
                if "#MC" in current and not current.startswith("+++"):
                    ne_name = current
                    break

        separator_index = -1
        for search_index in range(index + 1, len(lines)):
            current = lines[search_index].strip()
            if COMMAND_RE.match(current):
                break
            if SEPARATOR_RE.match(lines[search_index]):
                separator_index = search_index
                break

        if separator_index == -1 or separator_index + 1 >= len(lines):
            index += 1
            continue

        title_index = separator_index - 1
        while title_index > index and not lines[title_index].strip():
            title_index -= 1
        table_type = lines[title_index].strip()

        headers = _split_columns(lines[separator_index + 1])
        if not headers:
            index = separator_index + 1
            continue

        rows: list[list[str]] = []
        row_index = separator_index + 2
        while row_index < len(lines):
            raw_row = lines[row_index]
            row_text = raw_row.strip()
            if not row_text:
                row_index += 1
                continue
            if (
                RESULT_RE.match(row_text)
                or TABLE_END_RE.match(row_text)
                or raw_row.lstrip().startswith("+++")
                or COMMAND_RE.match(row_text)
            ):
                break

            if SEPARATOR_RE.match(raw_row):
                break

            values = _split_columns(raw_row)
            if values:
                rows.append(values)
            row_index += 1

        tables.append(
            MMLTable(
                command=command,
                ne_name=ne_name,
                table_type=table_type,
                headers=headers,
                rows=rows,
            )
        )
        index = max(row_index + 1, separator_index + 2)

    return tables


def _safe_sheet_name(command: str, used_names: set[str]) -> str:
    base = re.sub(r"[:;\s]+", "", command).strip()
    base = re.sub(r"[\\/?*\[\]]", "_", base)[:31] or "MML"
    candidate = base
    suffix = 2
    while candidate in used_names:
        ending = f"_{suffix}"
        candidate = f"{base[:31 - len(ending)]}{ending}"
        suffix += 1
    used_names.add(candidate)
    return candidate


def write_workbook(tables: Sequence[MMLTable], output_path: Path) -> None:
    """Write parsed MML tables using the legacy workbook layout."""
    workbook = Workbook()
    summary = workbook.active
    if summary is None:
        raise RuntimeError("Workbook did not create its default worksheet.")
    summary.title = "SUMMARY"
    sheet_dict = {}
    summary_count: dict[str, int] = {}
    used_sheet_names: set[str] = set()

    for table in tables:
        if table.command not in sheet_dict:
            sheet_name = _safe_sheet_name(table.command, used_sheet_names)
            worksheet = workbook.create_sheet(sheet_name)
            sheet_dict[table.command] = worksheet
            headers = ["NEName", "NEInfo", "TableType"] + table.headers
            for column_number, header in enumerate(headers, 1):
                cell = worksheet.cell(row=1, column=column_number, value=header)
                cell.font = Font(bold=True)

        worksheet = sheet_dict[table.command]
        row_number = worksheet.max_row + 1
        row_count = 0
        for values in table.rows:
            output_values = [
                table.ne_name,
                f"{table.command}:;",
                table.table_type,
            ] + values
            for column_number, value in enumerate(output_values, 1):
                worksheet.cell(row=row_number, column=column_number, value=value)
            row_number += 1
            row_count += 1

        sheet_name = worksheet.title
        summary_count[sheet_name] = summary_count.get(sheet_name, 0) + row_count

    summary.append(["Sheet Name", "Rows"])
    for cell in summary[1]:
        cell.font = Font(bold=True)

    for sheet_name, row_count in sorted(summary_count.items()):
        row_number = summary.max_row + 1
        safe_sheet_name = sheet_name.replace("'", "''")
        name_cell = summary.cell(row=row_number, column=1, value=sheet_name)
        name_cell.hyperlink = f"#'{safe_sheet_name}'!A1"
        name_cell.style = "Hyperlink"
        summary.cell(row=row_number, column=2, value=row_count)

    for worksheet in workbook.worksheets:
        worksheet.freeze_panes = "A2"
        for column_cells in worksheet.columns:
            column_index = column_cells[0].column
            max_length = max(
                (len(str(cell.value)) for cell in column_cells[:200] if cell.value),
                default=0,
            )
            worksheet.column_dimensions[get_column_letter(column_index)].width = min(
                max_length + 2, 50
            )

    workbook.save(output_path)


def _select_input_file() -> Path | None:
    from tkinter import Tk
    from tkinter.filedialog import askopenfilename

    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        selected = askopenfilename(
            title="Pilih File MML TXT",
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")],
        )
    finally:
        root.destroy()
    return Path(selected) if selected else None


def _read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8-sig")
    except UnicodeDecodeError:
        return path.read_text(encoding="cp1252")


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Parse a Huawei MML dump and export its tables to Excel."
    )
    parser.add_argument("input", nargs="?", type=Path, help="MML text file")
    parser.add_argument(
        "-o", "--output", type=Path, help="Output .xlsx path (default: beside input)"
    )
    args = parser.parse_args(argv)

    input_path = args.input or _select_input_file()
    if input_path is None:
        print("Tidak ada file yang dipilih.", file=sys.stderr)
        return 1
    if not input_path.is_file():
        print(f"File input tidak ditemukan: {input_path}", file=sys.stderr)
        return 1

    output_path = args.output or input_path.with_suffix(".xlsx")
    try:
        tables = parse_mml_text(_read_text(input_path))
        if not tables:
            print("Tidak ada tabel MML yang ditemukan.", file=sys.stderr)
            return 1
        write_workbook(tables, output_path)
    except OSError as error:
        print(f"Gagal membaca atau menyimpan file: {error}", file=sys.stderr)
        return 1

    total_rows = sum(len(table.rows) for table in tables)
    print(f"Input  : {input_path}")
    print(f"Output : {output_path}")
    print(f"Tabel  : {len(tables)}")
    print(f"Baris  : {total_rows}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
