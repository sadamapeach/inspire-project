import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
import time
import re
import io
from io import BytesIO

def round_half_up(series):
    return np.floor(series * 100 + 0.5) / 100

def format_rupiah(x):
    if pd.isna(x):
        return ""
    # pastikan bisa diubah ke float
    try:
        x = float(x)
    except:
        return x  # biarin apa adanya kalau bukan angka

    # kalau tidak punya desimal (misal 7000.0), tampilkan tanpa ,00
    if x.is_integer():
        formatted = f"{int(x):,}".replace(",", ".")
    else:
        formatted = f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        # hapus ,00 kalau desimalnya 0 semua (misal 7000,00 → 7000)
        if formatted.endswith(",00"):
            formatted = formatted[:-3]
    return formatted

def page():
    # Header Title
    st.markdown(
        """
        <div style="font-size:2.25rem; font-weight:700; margin-bottom:9px">
            ✂️ Table Extraction
        </div>
        """,
        unsafe_allow_html=True
    )
    # st.header("✂️ Table Extraction")
    st.markdown(
        ":red-badge[Indosat] :orange-badge[Ooredoo] :green-badge[Hutchison]"
    )
    st.caption("Upload your file — the tool will identify and split multiple tables within your sheets automatically ✨")

    # Divider custom
    st.markdown(
        """
        <hr style="margin-top:-5px; margin-bottom:10px; border: none; height: 2px; background-color: #ddd;">
        """,
        unsafe_allow_html=True
    )

    # File Uploader
    st.markdown("##### 📂 Upload File")
    upload_file = st.file_uploader("Upload your file here!", type=["xlsx", "xls"])

    if upload_file is not None:
        st.session_state["uploaded_file_table_extraction"] = upload_file

        # --- Animasi proses upload ---
        msg = st.toast("📂 Uploading file...")
        time.sleep(1.2)

        msg.toast("🔍 Reading sheets...")
        time.sleep(1.2)

        msg.toast("✅ File uploaded successfully!")
        time.sleep(0.5)

        # Baca semua sheet sekaligus
        all_df = pd.read_excel(upload_file, sheet_name=None)
        st.session_state["all_df_table_extraction_raw"] = all_df  # simpan versi mentah

    elif "all_df_table_extraction_raw" in st.session_state:
        all_df = st.session_state["all_df_table_extraction_raw"]
    else:
        return

    st.write("")
    # ====== TABLE EXTRACTION ======
    st.markdown("##### 📑 Result")
    all_sheets_tables = {}

    sheet_tabs = st.tabs(list(all_df.keys()))

    for tab, (sheet_name, df_raw) in zip(sheet_tabs, all_df.items()):
        with tab:
            df = df_raw.copy()

            # Logic 1: Split berdasarkan ROW kosong
            tables = []
            current_table = []

            # Split berdasarkan row kosong
            for idx, row in df.iterrows():
                if row.isna().all():
                    if current_table:
                        tables.append(pd.DataFrame(current_table))
                        current_table = []
                else:
                    current_table.append(row)

            if current_table:
                tables.append(pd.DataFrame(current_table))

            # Logic 2: Horizontal split berdasarkan NaN col
            final_tables = []

            for t in tables:
                # Cari kolom tidak kosong
                non_empty_cols = [c for c in t.columns if not t[c].isna().all()]
                if not non_empty_cols:
                    continue

                blocks = []
                current_block = [non_empty_cols[0]]

                # Horizontal block detection
                for col in non_empty_cols[1:]:
                    prev = t.columns.get_loc(current_block[-1])
                    curr = t.columns.get_loc(col)

                    if curr == prev + 1:
                        current_block.append(col)
                    else:
                        blocks.append(current_block)
                        current_block = [col]

                blocks.append(current_block)

                # Extract per-block
                for block_cols in blocks:
                    df_block = t[block_cols].copy()
                    final_tables.append(df_block)

            # Logic 3: Cleaning NaN row -> already clean the NaN col on the 2nd step
            clean_tables = []

            for i, t in enumerate(final_tables, start=1):
                df_clean = t.copy()
                df_clean = df_clean.dropna(axis=0, how='all')

                # Gunakan baris pertama sebagai header (hanya jika kolom belum ada nama atau Unnamed)
                if any("Unnamed" in str(c) for c in df_clean.columns):
                    df_clean.columns = df_clean.iloc[0]
                    df_clean = df_clean[1:].reset_index(drop=True)

                # Konversi tipe data otomatis ke pandas dtypes yang lebih fleksibel
                df_clean = df_clean.convert_dtypes()

                # Bersihkan tipe numpy di kolom, index, dan isi
                def safe_convert(x):
                    if isinstance(x, (np.generic, np.number)):
                        return x.item()
                    return x

                # Terapkan ke seluruh dataframe
                df_clean = df_clean.map(safe_convert)
                df_clean.columns = [safe_convert(c) for c in df_clean.columns]
                df_clean.index = [safe_convert(i) for i in df_clean.index]

                # Paksa semua header & index ke string agar JSON safe untuk Streamlit
                df_clean.columns = df_clean.columns.map(str)
                df_clean.index = df_clean.index.map(str)

                # Pembulatan
                num_cols = df_clean.select_dtypes(include=["number"]).columns
                df_clean[num_cols] = df_clean[num_cols].apply(round_half_up)

                # Format Rupiah
                df_clean_styled = df_clean.style.format({col: format_rupiah for col in num_cols})

                st.markdown(
                    f"""
                    <div style='display: flex; justify-content: space-between; 
                                align-items: center; margin-bottom: 8px;'>
                        <span style='font-size:14px;'>✨ {sheet_name} - Table {i}</span>
                        <span style='font-size:12px; color:#808080;'>
                            Total rows: <b>{len(df_clean):,}</b>
                        </span>
                    </div>
                    """,
                    unsafe_allow_html=True
                )
                st.dataframe(df_clean_styled, hide_index=True)

                clean_tables.append(df_clean)

            # Simpan hasil clean tables per sheet
            all_sheets_tables[sheet_name] = clean_tables

    st.divider()

    # SUPERR BOTTONN
    st.markdown("##### 🧑‍💻 Super Download — Export Selected Sheets")

    # Fungsi untuk convert list of DataFrame per sheet ke Excel
    def to_excel(dfs_dict):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book

            # Format rupiah
            format_rupiah_excel = workbook.add_format({'num_format': '#,##0.00'})

            for sheet_name, dfs in dfs_dict.items():
                for idx, df in enumerate(dfs, start=1):
                    df = df.copy()

                    numeric_cols = df.select_dtypes(include=['float', 'int']).columns
                    for col in numeric_cols:
                        df[col] = round_half_up(df[col])

                    sheet_tab_name = f"{sheet_name}_Table{idx}"

                    df.to_excel(writer, sheet_name=sheet_tab_name, index=False)

                    worksheet = writer.sheets[sheet_tab_name]

                    for col_idx, col_name in enumerate(df.columns):
                        if col_name in numeric_cols:
                            # apply number format to the entire column
                            worksheet.set_column(col_idx, col_idx, 18, format_rupiah_excel)

        return output.getvalue()
    
    # Buat multiselect untuk user pilih sheet
    selected_sheets = st.multiselect(
        "Select sheets to download in a single Excel file:",
        options=list(all_sheets_tables.keys()),
        default=list(all_sheets_tables.keys())
    )

    # --- FRAGMENT UNTUK BALLOONS ---
    @st.fragment
    def release_the_balloons():
        st.balloons()

    # ---- DOWNLOAD BUTTON ----
    if selected_sheets:
        # Filter dict sesuai pilihan user
        dfs_to_export = {k: v for k, v in all_sheets_tables.items() if k in selected_sheets}
        
        # Generate Excel
        excel_data = to_excel(dfs_to_export)

        # Tombol download
        st.download_button(
            label="Download",
            data=excel_data,
            file_name="Table Extraction.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            on_click=release_the_balloons,
            type="primary",
            use_container_width=True,        
        )
    else:
        st.info("Select at least one sheet to download.")