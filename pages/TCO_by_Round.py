import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
import time
import os
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

def highlight_total_row(row):
    # Cek apakah ada kolom yang berisi "TOTAL" (case-insensitive)
    if any(str(x).strip().upper() == "TOTAL" for x in row):
        return ["font-weight: bold;"] * len(row)
    else:
        return [""] * len(row)
    
def highlight_total_row_v2(row):
    # Cek apakah ada kolom yang berisi "TOTAL" (case-insensitive)
    if any(str(x).strip().upper() == "TOTAL" for x in row):
        return ["font-weight: bold; background-color: #D9EAD3; color: #1A5E20;"] * len(row)
    else:
        return [""] * len(row)
    
def highlight_1st_2nd_vendor(row, columns):
    styles = [""] * len(columns)
    first_vendor = row.get("1st Vendor")
    second_vendor = row.get("2nd Vendor")

    for i, col in enumerate(columns):
        if col == first_vendor:
            # styles[i] = "background-color: #f8c8dc; color: #7a1f47;"
            styles[i] = "background-color: #C6EFCE; color: #006100;"
        elif col == second_vendor:
            # styles[i] = "background-color: #d7c6f3; color: #402e72;"
            styles[i] = "background-color: #FFEB9C; color: #9C6500;"
    return styles
    
# Download button to Excel
@st.cache_data
def get_excel_download(df, sheet_name="Your_file_name"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        # --- Format untuk baris TOTAL ---
        bold_format = workbook.add_format({'bold': True})

        # Cari baris dengan label 'TOTAL' di kolom pertama
        total_rows = df.index[df.iloc[:, 0].astype(str).str.upper() == "TOTAL"].tolist()

        # Terapkan bold ke seluruh baris yang mengandung "TOTAL"
        for row in total_rows:
            worksheet.set_row(row + 1, None, bold_format)  # +1 karena header Excel mulai dari baris 1

        # (Opsional) Autofit kolom agar rapih
        for i, col in enumerate(df.columns):
            max_len = max(
                df[col].astype(str).map(len).max(),
                len(str(col))
            ) + 2
            worksheet.set_column(i, i, max_len)

    output.seek(0)
    return output.getvalue()

# Download highlight total
@st.cache_data
def get_excel_download_highlight_total(df, sheet_name="Sheet1"):
    output = BytesIO()

    # Buat file Excel dengan XlsxWriter
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        # Tentukan format
        highlight_format = workbook.add_format({
            "bold": True,
            "bg_color": "#D9EAD3",  # hijau lembut
            "font_color": "#1A5E20"  # hijau tua
        })

        # Jumlah kolom data
        num_cols = len(df.columns)

        # Iterasi baris (mulai dari baris 1 karena header di baris 0)
        for row_num, row_data in enumerate(df.itertuples(index=False), start=1):
            if any(str(x).strip().upper() == "TOTAL" for x in row_data if pd.notna(x)):
                # Highlight hanya sel di kolom yang berisi data
                for col_num in range(num_cols):
                    worksheet.write(row_num, col_num, row_data[col_num], highlight_format)

    return output.getvalue()

# Download Highlight 1st & 2nd Vendors
def get_excel_download_highlight_1st_2nd_lowest(df, sheet_name="Your_file_name"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        # --- Format umum ---
        format_first = workbook.add_format({'bg_color': '#C6EFCE'})  # hijau Excel-style
        format_second = workbook.add_format({'bg_color': '#FFEB9C'}) # kuning Excel-style

        # --- Loop baris dan kolom ---
        for row_idx, (_, row) in enumerate(df.iterrows(), start=1):
            first_vendor = row.get("1st Vendor")
            second_vendor = row.get("2nd Vendor")

            for col_idx, col in enumerate(df.columns):
                value = row[col]
                fmt = None

                # Tentukan warna highlight
                if col == first_vendor:
                    fmt = format_first
                elif col == second_vendor:
                    fmt = format_second

                # Handle semua jenis data NaN, inf, dan None
                if pd.isna(value) or (isinstance(value, (int, float)) and np.isinf(value)):
                    value = ""

                worksheet.write(row_idx, col_idx, value, fmt)

    return output.getvalue()

def page():
    # Header Title
    st.header("4️⃣ TCO Comparison Round by Round")
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

    # FILE UPLOADERR
    st.markdown("##### 📂 Upload Files")
    upload_files = st.file_uploader(
        "Upload multiple Excel files (e.g., L2R1.xlsx, L2R2.xlsx)", 
        type=["xlsx", "xls"],
        accept_multiple_files=True
    )

    # Pre-processing
    def clean_dataframe(df):
        df_clean = df.replace(r'^\s*$', None, regex=True)

        # Hapus baris dan kolom kosong
        df_clean = df_clean.dropna(how="all", axis=0).dropna(how="all", axis=1)

        # Jika header "Unnamed" -> set row 0 as header
        if any("Unnamed" in str(c) for c in df_clean.columns):
            df_clean.columns = df_clean.iloc[0]
            df_clean = df_clean[1:].reset_index(drop=True)

        # Konversi tipe data
        df_clean = df_clean.convert_dtypes()

        # Safe conversion from numpy types
        def safe_convert(x):
            if isinstance(x, (np.generic, np.number)):
                return x.item()
            return x
        
        df_clean = df_clean.map(safe_convert)
        df_clean.columns = [safe_convert(c) for c in df_clean.columns]
        df_clean.index = [safe_convert(i) for i in df_clean.index]

        # Paksa nama kolom dan index menjadi string
        df_clean.columns = df_clean.columns.map(str)
        df_clean.index = df_clean.index.map(str)

        return df_clean

    if upload_files:
        st.session_state["upload_multi_file_tco_by_round"] = upload_files
        files_to_process = upload_files
    
    elif "upload_multi_file_tco_by_round" in st.session_state:
        files_to_process = st.session_state["upload_multi_file_tco_by_round"]

    else:
        st.stop()
    
    all_rounds = []
    if "already_processed_tco_by_round" not in st.session_state:
        # --- Animasi proses upload ---
        msg = st.toast("📂 Uploading file...")
        time.sleep(1.2)

        msg.toast("🔍 Reading sheets...")
        time.sleep(1.2)

        msg.toast("✅ File uploaded successfully!")
        time.sleep(0.5)

    for file in files_to_process:
        # STEP 1: ambil nama file sebagai ROUND
        filename = os.path.splitext(file.name)[0]

        # STEP 2: baca sheet dalam file
        xls = pd.ExcelFile(file)
        df_raw = pd.read_excel(xls)
        df_clean = clean_dataframe(df_raw)  # cleaning

        # STEP 3: merge semua sheet dalam 1 file
        df_clean.insert(0, "ROUND", filename)

        # Tambahkan "TOTAL" jika belum ada
        scope_col = df_clean.columns[1]
        # Cek apakah sudah ada baris TOTAL
        has_total = df_clean.apply(
            lambda row: row.astype(str).str.upper().str.contains("TOTAL").any(),
            axis=1
        ).any()

        if not has_total:
            # Buat baris TOTAL baru
            total_row = {col: "" for col in df_clean.columns}

            # Isi nilai row TOTAL
            total_row["ROUND"] = filename
            total_row[scope_col] = "TOTAL"

            # Hitung jumlah vendor per kolom (skip NaN)
            vendor_cols = df_clean.select_dtypes(include=["int", "float"]).columns
            for v in vendor_cols:
                total_row[v] = df_clean[v].sum()

            # Tambahkan ke df
            df_clean = pd.concat([df_clean, pd.DataFrame([total_row])], ignore_index=True)
        
        # Masukkan ke list
        all_rounds.append(df_clean)

    # STEP 4: MERGE SEMUA FILEE
    df_final = pd.concat(all_rounds, ignore_index=True)

    # Simpan ke session
    st.session_state["merge_tco_by_round"] = df_final
    st.session_state["already_processed_tco_by_round"] = True

    st.divider()
   
    # MERGEE DATA
    st.markdown("##### 🗃️ Merge Data")
    st.caption(f"Successfully consolidated data from **{len(files_to_process)} files**.")

    # Pembulatan
    num_cols = df_final.select_dtypes(include=["number"]).columns
    df_final[num_cols] = df_final[num_cols].apply(round_half_up)

    # Format rupiah
    df_styled = (
        df_final.style
        .format({col: format_rupiah for col in num_cols})
        .apply(highlight_total_row_v2, axis=1)
    )
    st.dataframe(df_styled, hide_index=True)

    # Download
    excel_data = get_excel_download_highlight_total(df_final)
    col1, col2, col3 = st.columns([2.3,2,1])
    with col3:
        st.download_button(
            label="Download",
            data=excel_data,
            file_name="Merge TCO Round.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            icon=":material/download:",
        )

    st.divider()

    # COSTT SUMMARY
    st.markdown("##### 📑 Cost Summary")

    # Ambil semua kolom kecuali "ROUND"
    non_round_cols = [c for c in df_final.columns if c != "ROUND"]

    # Identifikasi kolom
    scope_cols = df_final[non_round_cols].select_dtypes(exclude=["number"]).columns.tolist()
    vendor_cols = df_final[non_round_cols].select_dtypes(include=['number']).columns.tolist()

    # Melt (Unpivot)
    df_summary = df_final.melt(
        id_vars=["ROUND"] + scope_cols,
        value_vars=vendor_cols,
        var_name="VENDOR",
        value_name='PRICE'
    )

    # Reorder
    final_cols = ["ROUND", "VENDOR"] + scope_cols + ["PRICE"]
    df_summary = df_summary[final_cols]

    # Sort
    df_summary = df_summary.sort_values(["ROUND", "VENDOR"] + scope_cols).reset_index(drop=True)

    # Slicer
    all_round  = sorted(df_summary["ROUND"].dropna().unique())
    all_vendor = sorted(df_summary["VENDOR"].dropna().unique())

    col_sel_1, col_sel_2 = st.columns(2)
    with col_sel_1:
        selected_round = st.multiselect(
            "Filter: Round",
            options=all_round,
            default=None,
            placeholder="Choose one or more rounds"
        )
    with col_sel_2:
        selected_vendor = st.multiselect(
            "Filter: Vendor",
            options=all_vendor,
            default=None,
            placeholder="Choose one or more vendors"
        )

    if selected_round and selected_vendor:
        df_filtered = df_summary[
            df_summary["ROUND"].isin(selected_round) &
            df_summary['VENDOR'].isin(selected_vendor)
        ]
    elif selected_round:
        df_filtered = df_summary[df_summary["ROUND"].isin(selected_round)]
    elif selected_vendor:
        df_filtered = df_summary[df_summary["VENDOR"].isin(selected_vendor)]
    else:
        df_filtered = df_summary.copy()

    # Format
    num_cols = df_filtered.select_dtypes(include=["number"]).columns
    df_filtered[num_cols] = df_filtered[num_cols].apply(round_half_up)

    df_summary_styled = (
        df_filtered.style
        .format({col: format_rupiah for col in num_cols})
        .apply(highlight_total_row_v2, axis=1)
    )

    st.caption(f"✨ Total number of data entries: **{len(df_filtered)}**")
    st.dataframe(df_summary_styled, hide_index=True)

    # Download
    excel_data = get_excel_download_highlight_total(df_filtered)
    col1, col2, col3 = st.columns([2.3,2,1])
    with col3:
        st.download_button(
            label="Download",
            data=excel_data,
            file_name="Cost Summary TCO Round.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            icon=":material/download:",
        )