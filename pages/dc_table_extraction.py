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
    st.header("✂️ Table Extraction")
    st.markdown(
        ":red-badge[Indosat] :orange-badge[Ooredoo] :green-badge[Hutchison]"
    )
    st.caption("Upload your pricing template — the tool will generate your analytics summary automatically ✨")

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
                df_clean = df_clean.applymap(safe_convert)
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
