"""Streamlit interface for parsing Huawei MML text dumps."""

from __future__ import annotations

from io import BytesIO
from pathlib import Path
import re

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
import streamlit as st


COMMAND_RE = re.compile(
    r"^\s*(?:%%/\*[^*]+\*/)?\s*([A-Z][A-Z0-9_ ]*:\s*;)(?:%%)?\s*$"
)
SEPARATOR_RE = re.compile(r"^\s*-{10,}\s*$")
RESULT_RE = re.compile(r"^\s*\(Number of results\s*=")
TABLE_END_RE = re.compile(r"^\s*(?:To be continued|---+\s*END\b)")


def split_columns(line: str) -> list[str]:
    return [value.strip() for value in re.split(r"\s{2,}", line.strip())]


def parse_mml_text(text: str) -> list[dict[str, object]]:
    lines = text.splitlines()
    tables: list[dict[str, object]] = []
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
            if COMMAND_RE.match(lines[search_index].strip()):
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
        headers = split_columns(lines[separator_index + 1])
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
                or SEPARATOR_RE.match(raw_row)
            ):
                break
            values = split_columns(raw_row)
            if values:
                rows.append(values)
            row_index += 1

        tables.append(
            {
                "command": command,
                "ne_name": ne_name,
                "table_type": lines[title_index].strip(),
                "headers": headers,
                "rows": rows,
            }
        )
        index = max(row_index + 1, separator_index + 2)

    return tables


def safe_sheet_name(command: str, used_names: set[str]) -> str:
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


def create_workbook(tables: list[dict[str, object]]) -> bytes:
    workbook = Workbook()
    summary = workbook.active
    if summary is None:
        raise RuntimeError("Workbook did not create its default worksheet.")
    summary.title = "SUMMARY"

    worksheets = {}
    summary_counts: dict[str, int] = {}
    used_names: set[str] = set()
    for table in tables:
        command = str(table["command"])
        if command not in worksheets:
            worksheet = workbook.create_sheet(safe_sheet_name(command, used_names))
            worksheets[command] = worksheet
            headers = ["NEName", "NEInfo", "TableType"] + table["headers"]
            for column, header in enumerate(headers, 1):
                worksheet.cell(row=1, column=column, value=header).font = Font(
                    bold=True
                )

        worksheet = worksheets[command]
        row_number = worksheet.max_row + 1
        rows = table["rows"]
        for values in rows:
            output_values = [
                table["ne_name"],
                f"{command}:;",
                table["table_type"],
            ] + values
            for column, value in enumerate(output_values, 1):
                worksheet.cell(row=row_number, column=column, value=value)
            row_number += 1
        summary_counts[worksheet.title] = (
            summary_counts.get(worksheet.title, 0) + len(rows)
        )

    summary.append(["Sheet Name", "Rows"])
    for cell in summary[1]:
        cell.font = Font(bold=True)
    for sheet_name, row_count in sorted(summary_counts.items()):
        row_number = summary.max_row + 1
        cell = summary.cell(row=row_number, column=1, value=sheet_name)
        safe_name = sheet_name.replace("'", "''")
        cell.hyperlink = f"#'{safe_name}'!A1"
        cell.style = "Hyperlink"
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

    output = BytesIO()
    workbook.save(output)
    return output.getvalue()


st.set_page_config(page_title="MML Dump Parser", page_icon="📄")
st.title("Huawei MML Dump Parser")
st.write("Unggah file MML TXT untuk mengubah tabelnya menjadi Excel.")

uploaded_file = st.file_uploader(
    "Pilih file MML",
    type=["txt"],
    accept_multiple_files=False,
)

if uploaded_file is not None:
    raw_content = uploaded_file.getvalue()
    try:
        content = raw_content.decode("utf-8-sig")
    except UnicodeDecodeError:
        content = raw_content.decode("cp1252")

    tables = parse_mml_text(content)
    if not tables:
        st.error("Tidak ada tabel MML yang ditemukan dalam file ini.")
    else:
        output = create_workbook(tables)
        total_rows = sum(len(table["rows"]) for table in tables)
        output_name = f"{Path(uploaded_file.name).stem}.xlsx"

        st.success(
            f"Berhasil memproses {len(tables)} blok laporan "
            f"dengan total {total_rows} baris."
        )
        st.download_button(
            "Unduh hasil Excel",
            data=output,
            file_name=output_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
