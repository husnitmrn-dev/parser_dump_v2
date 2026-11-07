"""Streamlit interface for parsing Huawei MML text dumps."""

from __future__ import annotations

from io import BytesIO
from pathlib import Path

import streamlit as st

from ubah_MML_dump_v2 import parse_mml_text, write_workbook


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
        output = BytesIO()
        write_workbook(tables, output)
        total_rows = sum(len(table.rows) for table in tables)
        output_name = f"{Path(uploaded_file.name).stem}.xlsx"

        st.success(
            f"Berhasil memproses {len(tables)} blok laporan "
            f"dengan total {total_rows} baris."
        )
        st.download_button(
            "Unduh hasil Excel",
            data=output.getvalue(),
            file_name=output_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
