import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
import time
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

def highlight_rank_summary(row, num_cols):
    styles = [""] * len(row)

    # Ambil nilai numeric vendor
    numeric_vals = row[num_cols]

    # EXCLUDE nilai 0 (vendor tidak ikut tender)
    numeric_vals = numeric_vals[numeric_vals != 0]

    # Skip jika kosong / NaN semua
    if numeric_vals.dropna().empty:
        return styles

    # Sort numeric values
    sorted_vals = numeric_vals.sort_values()

    # Determine 1st & 2nd rank
    first_vendor = sorted_vals.index[0]
    second_vendor = sorted_vals.index[1] if len(sorted_vals) > 1 else None

    # Apply styles
    for i, col in enumerate(row.index):
        if col == first_vendor:
            styles[i] = "background-color: #C6EFCE; color: #006100;"
        elif second_vendor and col == second_vendor:
            styles[i] = "background-color: #FFEB9C; color: #9C6500;"

    return styles

# Download button to Excel
@st.cache_data
def get_excel_download(df, sheet_name="Sheet1"):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)

        workbook  = writer.book
        worksheet = writer.sheets[sheet_name]

        # ================= FORMAT =================
        fmt_rupiah = workbook.add_format({"num_format": "#,##0"})
        fmt_pct = workbook.add_format({"num_format": '#,##0.0"%"'})

        # ================= COLUMN GROUP =================
        num_cols = df.select_dtypes(include=["number"]).columns.tolist()
        pct_cols = [c for c in df.columns if "%" in c]

        # ================= REWRITE CELLS =================
        for row_idx, row in enumerate(df.itertuples(index=False), start=1):
            for col_idx, col_name in enumerate(df.columns):
                val = row[col_idx]

                # Safety NaN / inf
                if pd.isna(val) or (isinstance(val, (int, float)) and np.isinf(val)):
                    worksheet.write(row_idx, col_idx, "")
                    continue

                # ===== PERCENT COLUMN =====
                if col_name in pct_cols:
                    worksheet.write_number(
                        row_idx,
                        col_idx,
                        val,
                        fmt_pct
                    )

                # ===== NUMERIC COLUMN =====
                elif col_name in num_cols:
                    worksheet.write_number(
                        row_idx,
                        col_idx,
                        val,
                        fmt_rupiah
                    )

                # ===== TEXT COLUMN =====
                else:
                    worksheet.write(row_idx, col_idx, val)

        # ================= AUTOFIT =================
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

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        workbook  = writer.book
        worksheet = writer.sheets[sheet_name]

        # ================= FORMAT =================
        format_rupiah = workbook.add_format({"num_format": "#,##0"})

        highlight_format = workbook.add_format({
            "bold": True,
            "bg_color": "#D9EAD3",
            "font_color": "#1A5E20",
            "num_format": "#,##0"
        })

        # ================= NUMERIC COLUMNS =================
        num_cols = df.select_dtypes(include=["number"]).columns

        # ================= LOOP DATA =================
        for row_idx, row in enumerate(df.itertuples(index=False), start=1):

            is_total = any(
                str(x).strip().upper() == "TOTAL"
                for x in row if pd.notna(x)
            )

            for col_idx, col_name in enumerate(df.columns):
                value = row[col_idx]

                # NaN / inf safety
                if pd.isna(value) or (isinstance(value, (int, float)) and np.isinf(value)):
                    value = ""

                # ===== TOTAL ROW =====
                if is_total:
                    if col_name in num_cols and value != "":
                        worksheet.write_number(
                            row_idx, col_idx, value, highlight_format
                        )
                    else:
                        worksheet.write(
                            row_idx, col_idx, value, highlight_format
                        )

                # ===== NORMAL ROW =====
                else:
                    if col_name in num_cols and value != "":
                        worksheet.write_number(
                            row_idx, col_idx, value, format_rupiah
                        )
                    else:
                        worksheet.write(
                            row_idx, col_idx, value
                        )

        # ================= AUTOFIT =================
        for i, col in enumerate(df.columns):
            max_len = max(
                df[col].astype(str).map(len).max(),
                len(str(col))
            ) + 2
            worksheet.set_column(i, i, max_len)

    output.seek(0)
    return output.getvalue()

# Download Highlight Summary
def get_excel_download_highlight_summary(df, sheet_name="Sheet1"):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)

        workbook  = writer.book
        worksheet = writer.sheets[sheet_name]

        # ================= FORMAT =================
        format_rupiah = workbook.add_format({"num_format": "#,##0"})
        format_bold = workbook.add_format({"bold": True})

        format_bold_rupiah = workbook.add_format({
            "bold": True,
            "num_format": "#,##0"
        })

        format_first = workbook.add_format({
            "bg_color": "#C6EFCE",
            "font_color": "#006100",
            "num_format": "#,##0"
        })

        format_second = workbook.add_format({
            "bg_color": "#FFEB9C",
            "font_color": "#9C6500",
            "num_format": "#,##0"
        })

        format_first_bold = workbook.add_format({
            "bg_color": "#C6EFCE",
            "font_color": "#006100",
            "bold": True,
            "num_format": "#,##0"
        })

        format_second_bold = workbook.add_format({
            "bg_color": "#FFEB9C",
            "font_color": "#9C6500",
            "bold": True,
            "num_format": "#,##0"
        })

        # ================= NUMERIC COLUMNS =================
        num_cols = df.select_dtypes(include=["number"]).columns

        # ================= LOOP DATA =================
        for row_idx, (_, row) in enumerate(df.iterrows(), start=1):

            is_total = any(str(x).strip().upper() == "TOTAL" for x in row)

            # ===== RANKING (EXCLUDE 0) =====
            numeric_vals = row[num_cols].dropna()
            numeric_vals = numeric_vals[numeric_vals != 0]

            first_vendor, second_vendor = None, None
            if not numeric_vals.empty:
                sorted_vals = numeric_vals.sort_values()
                first_vendor = sorted_vals.index[0]
                if len(sorted_vals) > 1:
                    second_vendor = sorted_vals.index[1]

            # ===== WRITE CELL =====
            for col_idx, col in enumerate(df.columns):
                value = row[col]

                if pd.isna(value) or (isinstance(value, (int, float)) and np.isinf(value)):
                    value = ""

                fmt = None

                # ----- NO HIGHLIGHT FOR ZERO -----
                if value == 0:
                    fmt = None

                else:
                    if col == first_vendor:
                        fmt = format_first_bold if is_total else format_first
                    elif col == second_vendor:
                        fmt = format_second_bold if is_total else format_second

                # ----- WRITE -----
                if col in num_cols and value != "":
                    worksheet.write_number(
                        row_idx,
                        col_idx,
                        value,
                        fmt or (format_bold_rupiah if is_total else format_rupiah)
                    )
                else:
                    worksheet.write(
                        row_idx,
                        col_idx,
                        value,
                        format_bold if is_total else None
                    )

        # ================= AUTOFIT =================
        for i, col in enumerate(df.columns):
            worksheet.set_column(
                i, i,
                max(len(str(col)), df[col].astype(str).map(len).max()) + 2
            )

    output.seek(0)
    return output.getvalue()

# Download Highlight Excel
def get_excel_download_highlight_1st_2nd_lowest(df, sheet_name="Sheet1"):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        workbook  = writer.book
        worksheet = writer.sheets[sheet_name]

        # ========== FORMAT ==========
        format_rupiah_xls = workbook.add_format({'num_format': '#,##0'})
        fmt_pct_rupiah = workbook.add_format({'num_format': '#,##0.0"%"'})

        format_first = workbook.add_format({
            'bg_color': '#C6EFCE',
            'font_color': '#006100',
            'num_format': '#,##0'
        })

        format_second = workbook.add_format({
            'bg_color': '#FFEB9C',
            'font_color': '#9C6500',
            'num_format': '#,##0'
        })

        # ========== COLUMN GROUP ==========
        num_cols = df.select_dtypes(include=["number"]).columns
        pct_cols = [c for c in df.columns if "%" in c]

        # ========== LOOP DATA ==========
        for row_idx, (_, row) in enumerate(df.iterrows(), start=1):
            first_vendor  = row.get("1st Vendor")
            second_vendor = row.get("2nd Vendor")

            for col_idx, col in enumerate(df.columns):
                value = row[col]

                # NaN / inf safety
                if pd.isna(value) or (isinstance(value, (int, float)) and np.isinf(value)):
                    worksheet.write(row_idx, col_idx, "")
                    continue

                # ===== PICK FORMAT =====
                fmt = None
                if col == first_vendor:
                    fmt = format_first
                elif col == second_vendor:
                    fmt = format_second

                # ===== WRITE CELL =====
                if col in pct_cols:
                    worksheet.write_number(
                        row_idx,
                        col_idx,
                        value,
                        fmt or fmt_pct_rupiah
                    )

                elif col in num_cols:
                    worksheet.write_number(
                        row_idx,
                        col_idx,
                        value,
                        fmt or format_rupiah_xls
                    )

                else:
                    worksheet.write(row_idx, col_idx, value)

        # ================= AUTOFIT =================
        for i, col in enumerate(df.columns):
            max_len = max(
                df[col].astype(str).map(len).max(),
                len(str(col))
            ) + 2
            worksheet.set_column(i, i, max_len)

    output.seek(0)
    return output.getvalue()

def page():
    # Header Title
    st.markdown(
        """
        <div style="font-size:2.25rem; font-weight:700; margin-bottom:9px">
            2️⃣ TCO Comparison by Region
        </div>
        """,
        unsafe_allow_html=True
    )
    # st.header("2️⃣ TCO Comparison by Region")
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
        st.session_state["uploaded_file_tco_by_region"] = upload_file

        # --- Animasi proses upload ---
        msg = st.toast("📂 Uploading file...")
        time.sleep(1.2)

        msg.toast("🔍 Reading sheets...")
        time.sleep(1.2)

        msg.toast("✅ File uploaded successfully!")
        time.sleep(0.5)

        # Baca semua sheet sekaligus
        all_df = pd.read_excel(upload_file, sheet_name=None)
        st.session_state["all_df_tco_by_region_raw"] = all_df  # simpan versi mentah

    elif "all_df_tco_by_region_raw" in st.session_state:
        all_df = st.session_state["all_df_tco_by_region_raw"]
    else:
        return
    
    st.write("")

    # OVERVIEW
    # st.markdown("##### 🔍 Overview")
    # Total bidders
    total_sheets = len(all_df)
    # st.caption(f"You're analyzing offers from **{total_sheets} participating bidders** in this session 🧐")

    result = {}

    for name, df in all_df.items():
        rows_before, cols_before = df.shape

        # Data cleaning
        df_clean = df.replace(r'^\s*$', np.nan, regex=True)
        df_clean = df_clean.dropna(how="all", axis=0).dropna(how="all", axis=1)

        # Header handling logic
        cols = df_clean.columns.astype(str)

        # Case 1: Semua kolom 'Unnamed'
        if len(cols) > 0 and all(col.startswith("Unnamed") for col in cols):
            # Gunakan baris pertama sebagai header
            df_clean.columns = df_clean.iloc[0]
            df_clean = df_clean[1:].reset_index(drop=True)

        # Case 2: Hanya beberapa kolom yang 'Unnamed'
        else:
            df_clean = df_clean.loc[:, ~cols.str.startswith("Unnamed")]

        rows_after, cols_after = df_clean.shape

        # Konversi tipe data otomatis
        df_clean = df_clean.convert_dtypes()

        # Bersihkan tipe numpy di kolom, index, dan isi
        def safe_convert(x):
            if isinstance(x, (np.generic, np.number)):
                return x.item()
            return x

        df_clean = df_clean.map(safe_convert)
        df_clean.columns = [safe_convert(c) for c in df_clean.columns]
        df_clean.index = [safe_convert(i) for i in df_clean.index]

        # Paksa semua header & index ke string agar JSON safe
        df_clean.columns = df_clean.columns.map(str)
        df_clean.index = df_clean.index.map(str)

        # Tambah kolom total
        if "TOTAL" not in df_clean.columns:
            df_clean["TOTAL"] = df_clean.sum(axis=1, numeric_only=True)

        # Pembulatan
        num_cols = df_clean.select_dtypes(include=["number"]).columns
        df_clean[num_cols] = df_clean[num_cols].apply(round_half_up)

        # Format Rupiah
        df_styled = df_clean.style.format({col: format_rupiah for col in num_cols})

        # # --- Styling: buat kolom "Total" jadi bold ---
        # df_styled = df_styled.set_properties(subset=["Total"], **{"font-weight": "bold"})

        # st.markdown(
        #     f"""
        #     <div style='display: flex; justify-content: space-between; 
        #                 align-items: center; margin-bottom: 8px;'>
        #         <span style='font-size:15px;'>✨ {name}</span>
        #         <span style='font-size:12px; color:#808080;'>
        #             Total rows: <b>{len(df_clean):,}</b>
        #         </span>
        #     </div>
        #     """,
        #     unsafe_allow_html=True
        # )

        result[name] = df_clean
        # st.dataframe(df_styled, hide_index=True)

        # # --- NOTIFIKASI KHUSUS ---
        # if (rows_after < rows_before) or (cols_after < cols_before):
        #     st.markdown(
        #         "<p style='font-size:12px; color:#808080; margin-top:-15px; margin-bottom:0;'>"
        #         "Preprocessing completed! Hidden rows and columns removed ✅</p>",
        #         unsafe_allow_html=True
        #     )

    st.session_state["result_tco_by_region"] = result
    # st.divider()

    tab1, tab2 = st.tabs(["ORIGINAL DATA", "TRANSPOSE DATA"])

    # MERGE
    tab1.markdown("##### 🗃️ Merge Data")

    # --- Merge all vendor data into one DataFrame ---
    merged_df = []

    for vendor_name, df_vendor in result.items():
        df_temp = df_vendor.copy()

        # Identifikasi kolom
        num_cols = df_temp.select_dtypes(include=["number"]).columns.tolist()

        # Buat baris TOTAL
        total_row = {col: "" for col in df_temp.columns}

        for col in df_temp.columns:
            if col == df_temp.columns[0]:
                total_row[col] = "TOTAL"
            elif col in num_cols:
                total_row[col] = df_temp[col].sum()
            else:
                total_row[col] = "" 
        
        # Tambahkan baris total
        df_temp = pd.concat([df_temp, pd.DataFrame([total_row])], ignore_index=True)

        # Tambahkan kolom vendor paling depan
        df_temp.insert(0, "VENDOR", vendor_name)

        merged_df.append(df_temp)

    # Gabungkan jadi satu dataframe
    df_merge = pd.concat(merged_df, ignore_index=True)

    # Simpan ke session_state jika perlu digunakan di halaman lain
    st.session_state["merged_all_data_tco_by_region"] = df_merge

    # Format Rupiah
    num_cols = df_merge.select_dtypes(include=["number"]).columns
    df_merge_styled = (
        df_merge.style
        .format({col: format_rupiah for col in num_cols})
        .apply(highlight_total_row_v2, axis=1)
    )

    tab1.caption(f"Data from **{total_sheets} vendors** have been successfully consolidated, analyzing a total of **{len(df_merge):,} combined records**.")
    tab1.dataframe(df_merge_styled, hide_index=True)

    # Download button to Excel
    excel_data = get_excel_download_highlight_total(df_merge)

    with tab1:
        # Layout tombol (rata kanan)
        col1, col2, col3 = st.columns([2.3,2,1])
        with col3:
            st.download_button(
                label="Download",
                data=excel_data,
                file_name="Merge Data - TCO by Region.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                icon=":material/download:",
            )

    tab1.divider()    

    # TCO PER SCOPEE
    tab1.markdown("##### 💸 TCO Summary — Scope")

    df = df_merge.drop(columns=["TOTAL"], errors="ignore").copy()

    # Identifikasi kolom
    vendor_col = "VENDOR"

    non_num_cols = df.select_dtypes(exclude="number").columns.tolist()
    non_num_cols = [c for c in non_num_cols if c != vendor_col] # selain vendor
    num_cols = df.select_dtypes(include="number").columns.tolist()

    # Kolom non-numeric pertama untuk naroh "TOTAL"
    first_non_num = non_num_cols[0]

    # hapus baris TOTAL
    df = df[df[first_non_num] != "TOTAL"]

    # Akumulasi seluruh region (semua kolom numerik)
    df["SUM_REGION"] = df[num_cols].sum(axis=1)

    # Pivot: non-num cols + ven. A + ven. B
    df_tco = (
        df.pivot_table(
            index=non_num_cols,
            columns=vendor_col,
            values="SUM_REGION",
            aggfunc="sum",
            fill_value=0
        )
        .reset_index()
    )

    # Tambahkan baris TOTAL
    vendor_cols = df_tco.columns[len(non_num_cols):]

    total_row = {col: "" for col in df_tco.columns}
    total_row[first_non_num] = "TOTAL"

    for v in vendor_cols:
        total_row[v] = df_tco[v].sum()

    df_tco = pd.concat([df_tco, pd.DataFrame([total_row])], ignore_index=True)

    # Fomat Rupiah & fungsi untuk styling baris TOTAL
    num_cols = df_tco.select_dtypes(include=["number"]).columns

    df_tco_styled = (
        df_tco.style
        .format({col: format_rupiah for col in num_cols})
        .apply(highlight_total_row, axis=1)
        .apply(lambda row: highlight_rank_summary(row, num_cols), axis=1)
    )

    tab1.markdown(
        """
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px;">
            <div style="font-size:0.88rem; color:gray; font-weight:400;">
                The following table presents the TCO summary across all scopes.
            </div>
            <div style="text-align:right;">
                <span style="background:#C6EFCE; padding:2px 8px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">1st Lowest</span>
                &nbsp;
                <span style="background:#FFEB9C; padding:2px 8px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">2nd Lowest</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.session_state["df_tco_by_region"] = df_tco
    tab1.dataframe(df_tco_styled, hide_index=True)

    # Download button to Excel
    excel_data = get_excel_download_highlight_summary(df_tco)

    with tab1:
        # Layout tombol (rata kanan)
        col1, col2, col3 = st.columns([2.3,2,1])
        with col3:
            st.download_button(
                label="Download",
                data=excel_data,
                file_name="TCO Summary - TCO by Region.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                icon=":material/download:",
            )

    tab1.divider()
    
    # GABUNGG SEMUA REGION JADI SATUU
    tab1.markdown("##### 🧠 Bid & Price Summary Analysis — Region")

    # Hapus kolom "TOTAL" 
    df_raw_analysis = df_merge.drop(columns=["TOTAL"], errors="ignore").copy()

    # Identifikasi kolom
    vendor_col = "VENDOR"
    
    non_num_cols = df_raw_analysis.select_dtypes(exclude=["number"]).columns.tolist()
    non_num_cols = [c for c in non_num_cols if c != vendor_col]

    first_non_num = non_num_cols[0] # kolom Scope

    num_cols = df_raw_analysis.select_dtypes(include=["number"]).columns.tolist()

    # Hapus row TOTAL
    df_raw_analysis = df_raw_analysis[
        ~df_raw_analysis.apply(
            lambda row: row.astype(str).str.upper().eq("TOTAL").any(),
            axis=1
        )
    ].copy()

    # Unpivot
    df_long = df_raw_analysis.melt(
        id_vars=[vendor_col] + non_num_cols,
        value_vars=num_cols,
        var_name="REGION",
        value_name="VALUE"
    )

    # Drop rows tanpa nilai
    df_long = df_long.dropna(subset=["VALUE"]).copy()

    # Pivot
    df_analysis = (
        df_long.pivot_table(
            index=["REGION"] + non_num_cols,
            columns=vendor_col,
            values="VALUE",
            aggfunc="sum",
            fill_value=0
        )
        .reset_index()
    )

    # Kolom vendor dinamis
    vendor_cols = df_analysis.select_dtypes(include=["number"]).columns.tolist()

    # Penanganan untuk 0 value
    vendor_values = df_analysis[vendor_cols].copy()
    vendor_values = vendor_values.replace(0, pd.NA)
    vendor_values = vendor_values.apply(pd.to_numeric, errors="coerce")

    # Hitung 1st dan 2nd lowest
    df_analysis["1st Lowest"] = vendor_values.min(axis=1)
    # df_analysis["1st Vendor"] = vendor_values.idxmin(axis=1)

    # Penanganan khusus 1st Vendor
    mask_all_nan1 = vendor_values.isna().all(axis=1)    # cek row yang all-NaN
    df_analysis["1st Vendor"] = None                    # inisialisasi
    valid_idx1 = (~mask_all_nan1)                       # ambil baris yang TIDAK all-NaN
    df_analysis.loc[valid_idx1, "1st Vendor"] = vendor_values.loc[valid_idx1].idxmin(axis=1) 

    # Hitung 2nd Lowest
    # Hilangkan dulu nilai 1st Lowest dari kandidat (agar kita dapat 2nd Lowest yang benar)
    temp = vendor_values.mask(vendor_values.eq(df_analysis["1st Lowest"], axis=0))

    df_analysis["2nd Lowest"] = temp.min(axis=1)
    # df_analysis["2nd Vendor"] = temp.idxmin(axis=1)

    # Penanganan khusus 2nd Vendor
    mask_all_nan2 = temp.isna().all(axis=1)     # cek row yang all-NaN
    df_analysis["2nd Vendor"] = None            # inisialisasi
    valid_idx2 = (~mask_all_nan2)               # ambil baris yang TIDAK all-NaN
    df_analysis.loc[valid_idx2, "2nd Vendor"] = temp.loc[valid_idx2].idxmin(axis=1) 

    # Hitung gap antara 1st dan 2nd lowest (%)
    df_analysis["Gap 1 to 2 (%)"] = ((df_analysis["2nd Lowest"] - df_analysis["1st Lowest"]) / df_analysis["1st Lowest"] * 100).round(2)

    # Hitung median price
    df_analysis["Median Price"] = vendor_values.median(axis=1)
    df_analysis["Median Price"] = pd.to_numeric(df_analysis["Median Price"], errors="coerce")

    # Hitung selisih tiap vendor dengan median (%)
    for v in vendor_cols:
        df_analysis[f"{v} to Median (%)"] = ((df_analysis[v] - df_analysis["Median Price"]) / df_analysis["Median Price"] * 100).round(2)

    # Simpan ke session state
    st.session_state["bid_and_price_analysis_tco_by_region"] = df_analysis

    # --- 🎯 Tambahkan slicer
    all_region = sorted(df_analysis["REGION"].dropna().unique())
    all_scope = sorted(df_analysis[first_non_num].dropna().unique())
    all_1st = sorted(df_analysis["1st Vendor"].dropna().unique())
    all_2nd = sorted(df_analysis["2nd Vendor"].dropna().unique())

    with tab1:
        col_sel_1, col_sel_2, col_sel_3, col_sel_4 = st.columns(4)
        with col_sel_1:
            selected_region = st.multiselect(
                "Filter: Region",
                options=all_region,
                default=[],
                placeholder="Choose regions",
                key="filter_region_2"
            )
        with col_sel_2:
            selected_scope = st.multiselect(
                "Filter: Scope",
                options=all_scope,
                default=[],
                placeholder="Choose scopes",
                key="filter_scope_2"
            )
        with col_sel_3:
            selected_1st = st.multiselect(
                "Filter: 1st vendor",
                options=all_1st,
                default=[],
                placeholder="Choose vendors",
                key="filter_1st_region"
            )
        with col_sel_4:
            selected_2nd = st.multiselect(
                "Filter: 2nd vendor",
                options=all_2nd,
                default=[],
                placeholder="Choose vendors",
                key="filter_2nd_region"
            )

        # --- Terapkan filter AND secara dinamis
        df_filtered_analysis = df_analysis.copy()

        if selected_region:
            df_filtered_analysis = df_filtered_analysis[df_filtered_analysis["REGION"].isin(selected_region)]

        if selected_scope:
            df_filtered_analysis = df_filtered_analysis[df_filtered_analysis[first_non_num].isin(selected_scope)]

        if selected_1st:
            df_filtered_analysis = df_filtered_analysis[df_filtered_analysis["1st Vendor"].isin(selected_1st)]

        if selected_2nd:
            df_filtered_analysis = df_filtered_analysis[df_filtered_analysis["2nd Vendor"].isin(selected_2nd)]

    # --- Styling (Rupiah & Persen) ---
    num_cols = df_filtered_analysis.select_dtypes(include=["number"]).columns
    format_dict = {col: format_rupiah for col in num_cols}
    format_dict.update({"Gap 1 to 2 (%)": "{:.1f}%"})
    for vendor in vendor_cols:
        format_dict[f"{vendor} to Median (%)"] = "{:+.1f}%"

    df_filtered_analysis_styled = (
        df_filtered_analysis.style
        .format(format_dict)
        .apply(lambda row: highlight_1st_2nd_vendor(row, df_filtered_analysis.columns), axis=1)
    )

    tab1.markdown(
        f"""
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px;">
            <div style="font-size:0.88rem; color:gray;">
                ✨ Total number of data entries: <b>{len(df_filtered_analysis)}</b>
            </div>
            <div style="text-align:right;">
                <span style="background:#C6EFCE; padding:2px 8px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">1st Lowest</span>
                &nbsp;
                <span style="background:#FFEB9C; padding:2px 8px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">2nd Lowest</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    # --- Tampilkan di Streamlit ---
    tab1.dataframe(df_filtered_analysis_styled, hide_index=True)

    excel_data = get_excel_download_highlight_1st_2nd_lowest(df_filtered_analysis)
    with tab1:
        # --- Optional: tombol download ---
        col1, col2, col3 = st.columns([2.3,2,1])
        with col3:
            st.download_button(
                label="Download",
                data=excel_data,
                file_name="Bid & Price Analysis - TCO by Region.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                icon=":material/download:",
            )

    # VISUALIZATIONN
    tab1.markdown("##### 📊 Visualization")

    with tab1:
        subtab1, subtab2 = st.tabs(["Win Rate Trend", "Average Gap Trend"])

        with subtab1:
            # --- WIN RATE VISUALIZATION ---
            # --- Hitung total kemenangan (1st & 2nd Vendor)
            win1_counts = df_analysis["1st Vendor"].value_counts(dropna=True).reset_index()
            win1_counts.columns = ["Vendor", "Wins1"]

            win2_counts = df_analysis["2nd Vendor"].value_counts(dropna=True).reset_index()
            win2_counts.columns = ["Vendor", "Wins2"]

            # --- Hitung total partisipasi vendor ---
            vendor_counts = (
                (df_analysis[vendor_cols].fillna(0) > 0)   # hanya True jika nilai > 0
                .sum()         # Hitung True per kolom
                .reset_index()
            )
            vendor_counts.columns = ["Vendor", "Total Participations"]

            # --- Gabungkan semua ---
            win_rate = (
                vendor_counts
                .merge(win1_counts, on="Vendor", how="left")
                .merge(win2_counts, on="Vendor", how="left")
                .fillna(0)
            )

            # --- Hitung Win Rate (%)
            win_rate["1st Win Rate (%)"] = np.where(
                win_rate["Total Participations"] > 0,
                (win_rate["Wins1"] / win_rate["Total Participations"] * 100).round(1),
                0
            )
            win_rate["2nd Win Rate (%)"] = np.where(
                win_rate["Total Participations"] > 0,
                (win_rate["Wins2"] / win_rate["Total Participations"] * 100).round(1),
                0
            )

            # --- Siapkan data long-format untuk visualisasi ---
            win_rate_long = win_rate.melt(
                id_vars=["Vendor"],
                value_vars=["1st Win Rate (%)", "2nd Win Rate (%)"],
                var_name="Metric",
                value_name="Percentage"
            )

            # --- Urutkan vendor berdasarkan 1st Win Rate tertinggi ---
            vendor_order = (
                win_rate.sort_values("1st Win Rate (%)", ascending=False)["Vendor"].tolist()
            )

            # --- Tentukan warna untuk kedua metrik ---
            metric_colors = {
                "1st Win Rate (%)": "#1f77b4",
                "2nd Win Rate (%)": "#ff7f0e"     # biru
            }

            # --- Highlight interaktif ---
            highlight = alt.selection_point(on='mouseover', fields=['Metric'], nearest=True)

            # --- Batas atas dan bawah sumbu Y ---
            y_max = win_rate_long["Percentage"].max()
            if not np.isfinite(y_max) or y_max <= 0:
                y_max = 1

            # Pastikan data diurutkan sesuai vendor_order
            win_rate_long["Vendor"] = pd.Categorical(win_rate_long["Vendor"], categories=vendor_order, ordered=True)
            win_rate_long = win_rate_long.sort_values(["Metric", "Vendor"])

            # --- Chart utama ---
            base = (
                alt.Chart(win_rate_long)
                .encode(
                    x=alt.X("Vendor:N", sort=vendor_order, title=None),
                    y=alt.Y(
                        "Percentage:Q",
                        title="Win Rate (%)",
                        scale=alt.Scale(domain=[0, y_max * 1.2])
                    ),
                    color=alt.Color(
                        "Metric:N",
                        scale=alt.Scale(
                            domain=list(metric_colors.keys()),
                            range=list(metric_colors.values())
                        ),
                        title="Rank"
                    ),
                    tooltip=[
                        alt.Tooltip("Vendor:N", title="Vendor"),
                        alt.Tooltip("Metric:N", title="Position"),
                        alt.Tooltip("Percentage:Q", title="Win Rate (%)", format=".1f")
                    ]
                )
            )

            # --- Garis dengan titik ---
            lines = base.mark_line(point=alt.OverlayMarkDef(size=70, filled=True), strokeWidth=3)

            # --- Label persentase di atas titik ---
            labels = base.mark_text(
                align='center',
                baseline='bottom',
                dy=-7,
                fontWeight='bold',
                color='gray'
            ).encode(
                text=alt.Text("Percentage:Q", format=".1f")
            ).transform_calculate(
                label="format(datum.Percentage, '.1f') + '%'"
            ).encode(
                text="label:N"
            )

            # --- Gabungkan semua elemen ---
            chart = (
                lines + labels
            ).properties(
                height=400,
                padding={"right": 15},
                title="Vendor Win Rate Comparison (1st vs 2nd Place)"
            ).configure_title(
                anchor="middle",
                offset=12,
                fontSize=14
            ).configure_axis(
                labelFontSize=12,
                titleFontSize=13
            ).configure_view(
                stroke='gray',
                strokeWidth=1
            ).configure_legend(
                titleFontSize=12,
                titleFontWeight="bold",
                labelFontSize=12,
                labelLimit=300,
                orient="bottom"
            )

            st.write("")
            # --- Tampilkan chart di Streamlit
            st.altair_chart(chart)

            # Kolom yang mau ditaruh di depan
            cols_front = ["Wins1", "Wins2"]

            # Sisanya (Vendor + kolom lain yang tidak ada di cols_front)
            cols_rest = [c for c in win_rate.columns if c not in cols_front]

            # Gabungkan urutannya
            win_rate = win_rate[cols_rest[:1] + cols_front + cols_rest[1:]]

            # --- Ganti nama kolom biar lebih konsisten & enak dibaca
            df_summary = win_rate.rename(columns={
                "Wins1": "1st Rank",
                "Wins2": "2nd Rank"
            })

            with st.expander("See explanation"):
                st.write('''
                    The visualization above compares the win rate of each vendor
                    based on how often they achieved 1st or 2nd place in all
                    tender evaluations.  
                            
                    **💡 How to interpret the chart**  
                            
                    - High 1st Win Rate (%)  
                        Vendor is highly competitive and often offers the best commercial terms.  
                    - High 2nd Win Rate (%)  
                        Vendor consistently performs well, often just slightly less competitive than the winner.  
                    - Large Gap Between 1st & 2nd Win Rate  
                        Shows clear market dominance by certain vendors.
                ''')

                num_cols = df_summary.select_dtypes(include=["number"]).columns
                format_dict = {col: format_rupiah for col in num_cols}
                format_dict.update({
                    "1st Win Rate (%)": "{:.1f}%",
                    "2nd Win Rate (%)": "{:.1f}%"
                })

                df_summary_styled = df_summary.style.format(format_dict)

                st.dataframe(df_summary_styled, hide_index=True)

                # Simpan hasil ke variabel
                excel_data = get_excel_download(df_summary)

                # Layout tombol (rata kanan)
                col1, col2, col3 = st.columns([3,1,1])
                with col3:
                    st.download_button(
                        label="Download",
                        data=excel_data,
                        file_name="Win Rate Trend Summary.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        icon=":material/download:",
                        key="download_Win_Rate_Trend_Tab1"
                    )

        with subtab2:
            # --- AVERAGE GAP VISUALIZATION ---
            # --- Hitung rata-rata Gap 1 to 2 (%) per Vendor (hanya untuk 1st Vendor)
            df_gap = df_analysis.copy()

            # Ubah kolom 'Gap 1 to 2 (%)' ke numerik (hapus simbol %)
            df_gap["Gap 1 to 2 (%)"] = (
                df_gap["Gap 1 to 2 (%)"]
                .replace("%", "", regex=True)
                .astype(float)
            )

            # Hitung rata-rata gap per vendor (hanya vendor yang jadi 1st Lowest)
            avg_gap = (
                df_gap.groupby("1st Vendor", dropna=True)["Gap 1 to 2 (%)"]
                .mean()
                .reset_index()
                .rename(columns={"1st Vendor": "Vendor", "Gap 1 to 2 (%)": "Average Gap (%)"})
                .sort_values("Average Gap (%)", ascending=False)
            )

            # st.dataframe(avg_gap)

            # Warna per vendor (biar konsisten kalau kamu sudah punya color mapping)
            colors_list = ["#F94144", "#F3722C", "#F8961E", "#F9C74F", "#90BE6D",
                        "#43AA8B", "#577590", "#E54787", "#BF219A", "#8E0F9C", "#4B1D91"]
            vendor_colors = {v: c for v, c in zip(avg_gap["Vendor"].unique(), colors_list)}

            # Interaksi hover
            highlight = alt.selection_point(on='mouseover', fields=['Vendor'], nearest=True)

            # --- Chart utama ---
            bars = (
                alt.Chart(avg_gap)
                .mark_bar()
                .encode(
                    x=alt.X("Vendor:N", sort='-y', title=None),
                    y=alt.Y("Average Gap (%):Q", title="Average Gap (%)", scale=alt.Scale(domain=[0, avg_gap["Average Gap (%)"].max() * 1.2])),
                    color=alt.Color("Vendor:N",
                                    scale=alt.Scale(domain=list(vendor_colors.keys()), range=list(vendor_colors.values())),
                                    legend=None),
                    tooltip=[
                        alt.Tooltip("Vendor:N", title="Vendor"),
                        alt.Tooltip("Average Gap (%):Q", title="Average Gap (%)", format=".1f")
                    ]
                )
                .add_params(highlight)
            )

            # Label teks di atas bar
            labels = (
                alt.Chart(avg_gap)
                .mark_text(dy=-7, fontWeight='bold', color='gray')
                .encode(
                    x="Vendor:N",
                    y="Average Gap (%):Q",
                    text=alt.Text("Average Gap (%):Q", format=".1f")  # Format angka
                )
                .transform_calculate(  # Tambahkan simbol %
                    label_text="format(datum['Average Gap (%)'], '.1f') + '%'"
                )
                .encode(
                    text="label_text:N"
                )
            )

            # Frame luar untuk gaya rapi
            frame = alt.Chart().mark_rect(stroke='gray', strokeWidth=1, fillOpacity=0)

            avg_line = alt.Chart(avg_gap).mark_rule(color='gray', strokeDash=[4,2], size=1.75).encode(
                y='mean(Average Gap (%)):Q'
            )

            # Gabungkan semua elemen
            chart = (bars + labels + frame + avg_line).properties(
                title="Average Gap (%) per 1st Vendor",
                height=400
            ).configure_title(
                anchor='middle',
                offset=12,
                fontSize=14
            ).configure_axis(
                labelFontSize=12,
                titleFontSize=13
            ).configure_view(
                stroke='gray',
                strokeWidth=1
            )

            st.write("")
            # --- Tampilkan di Streamlit ---
            st.altair_chart(chart)

            avg_value = avg_gap["Average Gap (%)"].mean()

            # with st:
            with st.expander("See explanation"):
                st.write(f'''
                    The chart above shows the average price difference between 
                    the lowest and second-lowest bids for each vendor when they 
                    rank 1st, indicating their pricing dominance or competitiveness.
                            
                    **💡 How to interpret the chart**  
                            
                    - High Gap  
                        High gap indicates strong vendor dominance (much lower prices).  
                    - Low Gap  
                        Low gap indicates intense competition with similar pricing among vendors.  
                    
                    The dashed line represents the average gap across all vendors, serving as a benchmark ({avg_value:.1f}%).
                ''')

    tab1.divider()

    # SUPERRR BUTTOONNN 1
    tab1.markdown("##### 🧑‍💻 Super Download — Export Selected Sheets")
    dataframes = {
        "Merge Data": df_merge,
        "TCO Summary": df_tco,
        "Bid & Price Analysis": df_filtered_analysis,
    }

    # Tampilkan multiselect
    selected_sheets = tab1.multiselect(
        "Select sheets to download in a single Excel file:",
        options=list(dataframes.keys()),
        default=list(dataframes.keys())  # default semua dipilih
    )

    # Fungsi "Super Button" & Formatting
    def generate_multi_sheet_excel(selected_sheets, df_dict):
        output = BytesIO()

        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            for sheet in selected_sheets:
                df = df_dict[sheet]
                df.to_excel(writer, index=False, sheet_name=sheet)

                workbook  = writer.book
                worksheet = writer.sheets[sheet]

                # ===== FORMAT =====
                fmt_rp   = workbook.add_format({'num_format': '#,##0'})
                fmt_pct  = workbook.add_format({'num_format': '#,##0.0"%"'})
                fmt_bold = workbook.add_format({'bold': True, 'num_format': '#,##0'})

                fmt_total = workbook.add_format({
                    'bold': True,
                    'bg_color': '#D9EAD3',
                    'font_color': '#1A5E20',
                    'num_format': '#,##0'
                })

                fmt_1 = workbook.add_format({
                    'bg_color': '#C6EFCE',
                    'font_color': '#006100',
                    'num_format': '#,##0'
                })

                fmt_2 = workbook.add_format({
                    'bg_color': '#FFEB9C',
                    'font_color': '#9C6500',
                    'num_format': '#,##0'
                })

                fmt_1b = workbook.add_format({
                    'bg_color': '#C6EFCE',
                    'font_color': '#006100',
                    'bold': True,
                    'num_format': '#,##0'
                })

                fmt_2b = workbook.add_format({
                    'bg_color': '#FFEB9C',
                    'font_color': '#9C6500',
                    'bold': True,
                    'num_format': '#,##0'
                })

                num_cols = df.select_dtypes(include=["number"]).columns.tolist()
                pct_cols = [c for c in df.columns if "%" in c]

                # ===== LOOP DATA =====
                for r, (_, row) in enumerate(df.iterrows(), start=1):
                    is_total = any(str(x).strip().upper() == "TOTAL" for x in row)

                    first = second = None

                    # ===== SUMMARY RANKING (EXCLUDE ZERO) =====
                    if sheet == "TCO Summary":
                        numeric_vals = row[num_cols]
                        numeric_vals = numeric_vals[
                            (numeric_vals.notna()) & (numeric_vals != 0)
                        ]

                        if not numeric_vals.empty:
                            sorted_vals = numeric_vals.sort_values()
                            first = sorted_vals.index[0]
                            if len(sorted_vals) > 1:
                                second = sorted_vals.index[1]

                    # ===== BID & PRICE =====
                    if sheet == "Bid & Price Analysis":
                        first = row.get("1st Vendor")
                        second = row.get("2nd Vendor")

                    for c, col in enumerate(df.columns):
                        val = row[col]

                        # ===== SAFETY =====
                        if pd.isna(val) or (isinstance(val, (int, float)) and np.isinf(val)):
                            worksheet.write(r, c, "")
                            continue

                        fmt = None

                        is_zero = isinstance(val, (int, float)) and val == 0
                        # ===== NO HIGHLIGHT FOR ZERO (EXCEPT MERGE DATA) =====
                        if is_zero and sheet != "Merge Data":
                            fmt = None

                        # ===== PICK FORMAT =====
                        elif col == first:
                            fmt = fmt_1b if is_total else fmt_1
                        elif col == second:
                            fmt = fmt_2b if is_total else fmt_2
                        elif is_total and sheet == "Merge Data":
                            fmt = fmt_total
                        elif is_total:
                            fmt = fmt_bold

                        # ===== WRITE CELL =====
                        if col in pct_cols:
                            worksheet.write_number(r, c, val, fmt or fmt_pct)
                        elif col in num_cols:
                            worksheet.write_number(r, c, val, fmt or fmt_rp)
                        else:
                            worksheet.write(r, c, val, fmt)

                    # ===== AUTOFIT =====
                    for i, col in enumerate(df.columns):
                        worksheet.set_column(
                            i, i,
                            max(len(str(col)), df[col].astype(str).map(len).max()) + 2
                        )

        output.seek(0)
        return output.getvalue()

    # --- FRAGMENT UNTUK BALLOONS ---
    @st.fragment
    def release_the_balloons():
        tab1.balloons()

    # ---- DOWNLOAD BUTTON ----
    if selected_sheets:
        excel_bytes = generate_multi_sheet_excel(selected_sheets, dataframes)

        tab1.download_button(
            label="Download",
            data=excel_bytes,
            file_name="TCO Comparison by Region - Original Data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            on_click=release_the_balloons,
            type="primary",
            use_container_width=True,
        )

    # TRANSPOSE
    # tab2.markdown("##### 🛸 Transpose Data")
    # Total bidders
    total_sheets = len(all_df)
    # tab2.caption(f"Data has been transposed — each record now represents a vendor’s offer across all regions!")

    result = {}

    for name, df in all_df.items():
        df_clean_transposed = df.copy()

        # Preprocessing
        df_clean_transposed = df_clean_transposed.replace(r'^\s*$', None, regex=True)
        df_clean_transposed = df_clean_transposed.dropna(how="all", axis=0).dropna(how="all", axis=1)

        # Gunakan baris pertama sebagai header (hanya jika kolom belum ada nama atau Unnamed)
        if any("Unnamed" in str(c) for c in df_clean_transposed.columns):
            df_clean_transposed.columns = df_clean_transposed.iloc[0]
            df_clean_transposed = df_clean_transposed[1:].reset_index(drop=True)

        # Konversi tipe data otomatis
        df_clean_transposed = df_clean_transposed.convert_dtypes()

        # Bersihkan tipe numpy di kolom, index, dan isi
        def safe_convert(x):
            if isinstance(x, (np.generic, np.number)):
                return x.item()
            return x

        df_clean_transposed = df_clean_transposed.map(safe_convert)
        df_clean_transposed.columns = [safe_convert(c) for c in df_clean_transposed.columns]
        df_clean_transposed.index = [safe_convert(i) for i in df_clean_transposed.index]

        # Paksa semua header & index ke string agar JSON safe
        df_clean_transposed.columns = df_clean_transposed.columns.map(str)
        df_clean_transposed.index = df_clean_transposed.index.map(str)

        # Identifikasi kolom
        non_num_cols = df_clean_transposed.select_dtypes(exclude=["number"]).columns.tolist()
        num_cols = df_clean_transposed.select_dtypes(include=["number"]).columns.tolist()

        # Kolom pivot = non-num[0] (misal "scope")
        pivot_col = non_num_cols[0]

        # Kolom info tambahan selain pivot (misal desc, uom)
        info_cols = non_num_cols[1:]

        # Unpivot region -> menjadi baris
        df_long = df_clean_transposed.melt(
            id_vars=[pivot_col] + info_cols,
            value_vars=num_cols,
            var_name="REGION",
            value_name="VALUE"
        ).dropna(subset=["VALUE"])

        # Pivot balik -> region jadi baris, scope jadi kolom
        df_transposed = (
            df_long.pivot_table(
                index=["REGION"] + info_cols,
                columns=pivot_col,
                values="VALUE",
                aggfunc="sum",
                fill_value=0
            )
            .reset_index()
        )

        # Tambah TOTAL
        if "TOTAL" not in df_transposed.columns:
            df_transposed["TOTAL"] = df_transposed.select_dtypes(include=["number"]).sum(axis=1)

        # --- Formatting rupiah ---
        num_cols_transposed = df_transposed.select_dtypes(include=["number"]).columns
        df_styled_transposed = df_transposed.style.format({col: format_rupiah for col in num_cols_transposed})

        # tab2.markdown(
        #     f"""
        #     <div style='display: flex; justify-content: space-between; 
        #                 align-items: center; margin-bottom: 8px;'>
        #         <span style='font-size:15px;'>✨ {name}</span>
        #         <span style='font-size:12px; color:#808080;'>
        #             Total rows: <b>{len(df_transposed):,}</b>
        #         </span>
        #     </div>
        #     """,
        #     unsafe_allow_html=True
        # )

        # tab2.dataframe(df_styled_transposed, hide_index=True)

        result[name] = df_transposed

        # excel_data = get_excel_download(df_transposed)
        # col1, col2, col3 = tab2.columns([2.3,2,1])
        # with col3:
        #     tab2.download_button(
        #         label="Download",
        #         data=excel_data,
        #         file_name=f"Transposed_{name}.xlsx",
        #         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        #         icon=":material/download:",
        #         key=f"download_transposed_{name}"
        #     )

    # tab2.divider()

    # MERGE TRANSPOSED
    tab2.markdown("##### 🗃️ Merge Transposed")

    # --- Merge all vendor data into one DataFrame ---
    df_merge_transposed = []

    for vendor_name, df_vendor in result.items():
        df_temp = df_vendor.copy()

        # Identifikasi kolom
        num_cols = df_temp.select_dtypes(include=["number"]).columns.tolist()

        # Buat baris TOTAL
        total_row = {col: "" for col in df_temp.columns}

        for col in df_temp.columns:
            if col == df_temp.columns[0]:
                total_row[col] = "TOTAL"
            elif col in num_cols:
                total_row[col] = df_temp[col].sum()
            else:
                total_row[col] = "" 
        
        # Tambahkan baris total
        df_temp = pd.concat([df_temp, pd.DataFrame([total_row])], ignore_index=True)

        # Tambahkan kolom vendor paling depan
        df_temp.insert(0, "VENDOR", vendor_name)

        df_merge_transposed.append(df_temp)

    # Gabungkan jadi satu dataframe
    df_merge_transposed = pd.concat(df_merge_transposed, ignore_index=True)

    # Simpan ke session_state jika perlu digunakan di halaman lain
    st.session_state["merged_all_data_transposed_tco_by_region"] = df_merge_transposed

    # Format Rupiah
    num_cols = df_merge_transposed.select_dtypes(include=["number"]).columns
    df_merge_transposed_styled = (
        df_merge_transposed.style
        .format({col: format_rupiah for col in num_cols})
        .apply(highlight_total_row_v2, axis=1)
    )

    tab2.caption(f"Data from **{total_sheets} vendors** have been successfully consolidated, analyzing a total of **{len(df_merge_transposed):,} combined records**.")
    tab2.dataframe(df_merge_transposed_styled, hide_index=True)

    # Download button to Excel
    excel_data = get_excel_download_highlight_total(df_merge_transposed)

    with tab2:
        # Layout tombol (rata kanan)
        col1, col2, col3 = st.columns([2.3,2,1])
        with col3:
            st.download_button(
                label="Download",
                data=excel_data,
                file_name="Merge Transposed - TCO by Region.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                icon=":material/download:",
            )

    tab2.divider()

    # TCO per Regionn
    tab2.markdown("##### 💸 TCO Summary — Region")
    
    # Hapus kolom & row "TOTAL"
    df = df_merge_transposed.drop(columns=["TOTAL"], errors="ignore").copy()
    df = df[df["REGION"].astype(str).str.upper() != "TOTAL"].copy()

    # Identifikasi kolom
    vendor_col = "VENDOR"
    region_col = "REGION"

    # Kolom non numeric selain vendor & region
    non_num_cols = df.select_dtypes(exclude="number").columns.tolist()
    non_num_cols = [c for c in non_num_cols if c not in [vendor_col, region_col]]

    # Kolom scope numerik
    scope_cols = df.select_dtypes(include="number").columns.tolist()

    # Hitung total per-vendor per-region
    df["SUM_SCOPE"] = df[scope_cols].sum(axis=1)

    # Pivot
    df_tco_transposed = (
        df.pivot_table(
            index=[region_col] + non_num_cols,
            columns=vendor_col,
            values="SUM_SCOPE",
            aggfunc="sum",
            fill_value=0
        )
        .reset_index()
    )

    # Ambil daftar vendor
    vendor_list = df_tco_transposed.columns[len([region_col] + non_num_cols):]

    # Tambah row "TOTAL"
    total_row = {col: "" for col in df_tco_transposed.columns}
    total_row[region_col] = "TOTAL"

    for v in vendor_list:
        total_row[v] = df_tco_transposed[v].sum()

    df_tco_transposed = pd.concat([df_tco_transposed, pd.DataFrame([total_row])], ignore_index=True)

    # Format Rupiah & highlight baris TOTAL
    num_cols = df_tco_transposed.select_dtypes(include=["number"]).columns

    df_tco_transposed_styled = (
        df_tco_transposed.style
        .format({col: format_rupiah for col in num_cols})
        .apply(highlight_total_row, axis=1)
        .apply(lambda row: highlight_rank_summary(row, num_cols), axis=1)
    )

    tab2.markdown(
        """
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px;">
            <div style="font-size:0.88rem; color:gray; font-weight:400;">
                The following table presents the TCO summary across all regions.
            </div>
            <div style="text-align:right;">
                <span style="background:#C6EFCE; padding:2px 8px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">1st Lowest</span>
                &nbsp;
                <span style="background:#FFEB9C; padding:2px 8px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">2nd Lowest</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.session_state["df_tco_transposed_by_region"] = df_tco_transposed
    tab2.dataframe(df_tco_transposed_styled, hide_index=True)

    # Download button to Excel
    excel_data = get_excel_download_highlight_summary(df_tco_transposed)

    with tab2:
        # Layout tombol (rata kanan)
        col1, col2, col3 = st.columns([2.3,2,1])
        with col3:
            st.download_button(
                label="Download",
                data=excel_data,
                file_name="TCO Transposed - TCO by Region.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                icon=":material/download:",
            )

    tab2.divider()

    # GABUNGG SEMUA SCOPE JADI SATU
    tab2.markdown("##### 🧠 Bid & Price Summary Analysis — Scope")

    # Hapus kolom "TOTAL" 
    df_raw_analysis_transpose = df_merge_transposed.drop(columns=["TOTAL"], errors="ignore").copy()

    # Identifikasi kolom
    vendor_col = "VENDOR"
    
    non_num_cols = df_raw_analysis_transpose.select_dtypes(exclude=["number"]).columns.tolist()
    non_num_cols = [c for c in non_num_cols if c != vendor_col]

    first_non_num = non_num_cols[0] # kolom Region

    num_cols = df_raw_analysis_transpose.select_dtypes(include=["number"]).columns.tolist()

    # Hapus row TOTAL
    df_raw_analysis_transpose = df_raw_analysis_transpose[
        ~df_raw_analysis_transpose.apply(
            lambda row: row.astype(str).str.upper().eq("TOTAL").any(),
            axis=1
        )
    ].copy()

    # Unpivot
    df_long = df_raw_analysis_transpose.melt(
        id_vars=[vendor_col] + non_num_cols,
        value_vars=num_cols,
        var_name="SCOPE",
        value_name="VALUE"
    )

    # Drop rows tanpa nilai
    df_long = df_long.dropna(subset=["VALUE"]).copy()

    # Pivot
    df_analysis_transposed = (
        df_long.pivot_table(
            index=["SCOPE"] + non_num_cols,
            columns=vendor_col,
            values="VALUE",
            aggfunc="sum",
            fill_value=0
        )
        .reset_index()
    )

    # Kolom vendor dinamis
    vendor_cols = df_analysis_transposed.select_dtypes(include=["number"]).columns.tolist()

    # Penanganan untuk 0 value
    vendor_values = df_analysis_transposed[vendor_cols].copy()
    vendor_values = vendor_values.replace(0, pd.NA)
    vendor_values = vendor_values.apply(pd.to_numeric, errors="coerce")

    # Hitung 1st dan 2nd lowest
    df_analysis_transposed["1st Lowest"] = vendor_values.min(axis=1)
    # df_analysis_transposed["1st Vendor"] = vendor_values.idxmin(axis=1)

    # Penanganan khusus 1st Vendor
    mask_all_nan3 = vendor_values.isna().all(axis=1)    # cek row yang all-NaN
    df_analysis_transposed["1st Vendor"] = None         # inisialisasi
    valid_idx3 = (~mask_all_nan3)                       # ambil baris yang TIDAK all-NaN
    df_analysis_transposed.loc[valid_idx3, "1st Vendor"] = vendor_values.loc[valid_idx3].idxmin(axis=1) 

    # Hitung 2nd Lowest
    # Hilangkan dulu nilai 1st Lowest dari kandidat (agar kita dapat 2nd Lowest yang benar)
    temp = vendor_values.mask(vendor_values.eq(df_analysis_transposed["1st Lowest"], axis=0))

    df_analysis_transposed["2nd Lowest"] = temp.min(axis=1)
    # df_analysis_transposed["2nd Vendor"] = temp.idxmin(axis=1)

    # Penanganan khusus 2nd Vendor
    mask_all_nan4 = temp.isna().all(axis=1)     # cek row yang all-NaN
    df_analysis_transposed["2nd Vendor"] = None # inisialisasi
    valid_idx4 = (~mask_all_nan4)               # ambil baris yang TIDAK all-NaN
    df_analysis_transposed.loc[valid_idx4, "2nd Vendor"] = temp.loc[valid_idx4].idxmin(axis=1) 

    # Hitung gap antara 1st dan 2nd lowest (%)
    df_analysis_transposed["Gap 1 to 2 (%)"] = ((df_analysis_transposed["2nd Lowest"] - df_analysis_transposed["1st Lowest"]) / df_analysis_transposed["1st Lowest"] * 100).round(2)

    # Hitung median price
    df_analysis_transposed["Median Price"] = vendor_values.median(axis=1)
    df_analysis_transposed["Median Price"] = pd.to_numeric(df_analysis_transposed["Median Price"], errors="coerce")

    # Hitung selisih tiap vendor dengan median (%)
    for v in vendor_cols:
        df_analysis_transposed[f"{v} to Median (%)"] = ((df_analysis_transposed[v] - df_analysis_transposed["Median Price"]) / df_analysis_transposed["Median Price"] * 100).round(2)

    # Simpan ke session state
    st.session_state["bid_and_price_analysis_transposed_tco_by_region"] = df_analysis_transposed

    # --- 🎯 Tambahkan slicer
    all_scope = sorted(df_analysis_transposed["SCOPE"].dropna().unique())
    all_region = sorted(df_analysis_transposed[first_non_num].dropna().unique())
    all_1st = sorted(df_analysis_transposed["1st Vendor"].dropna().unique())
    all_2nd = sorted(df_analysis_transposed["2nd Vendor"].dropna().unique())

    with tab2:
        col_sel_1, col_sel_2, col_sel_3, col_sel_4 = st.columns(4)
        with col_sel_1:
            selected_scope = st.multiselect(
                "Filter: Scope",
                options=all_scope,
                default=[],
                placeholder="Choose scopes",
                key="filter_scope_1"
            )
        with col_sel_2:
            selected_region = st.multiselect(
                "Filter: Region",
                options=all_region,
                default=[],
                placeholder="Choose regions",
                key="filter_region_1"
            )
        with col_sel_3:
            selected_1st = st.multiselect(
                "Filter: 1st vendor",
                options=all_1st,
                default=[],
                placeholder="Choose vendors",
                key="filter_1st_scope"
            )
        with col_sel_4:
            selected_2nd = st.multiselect(
                "Filter: 2nd vendor",
                options=all_2nd,
                default=[],
                placeholder="Choose vendors",
                key="filter_2nd_scope"
            )

        # --- Terapkan filter AND secara dinamis
        df_filtered_transposed = df_analysis_transposed.copy()

        if selected_scope:
            df_filtered_transposed = df_filtered_transposed[df_filtered_transposed["SCOPE"].isin(selected_scope)]

        if selected_region:
            df_filtered_transposed = df_filtered_transposed[df_filtered_transposed[first_non_num].isin(selected_region)]

        if selected_1st:
            df_filtered_transposed = df_filtered_transposed[df_filtered_transposed["1st Vendor"].isin(selected_1st)]

        if selected_2nd:
            df_filtered_transposed = df_filtered_transposed[df_filtered_transposed["2nd Vendor"].isin(selected_2nd)]


    # --- Styling (Rupiah & Persen) ---
    num_cols = df_filtered_transposed.select_dtypes(include=["number"]).columns
    format_dict = {col: format_rupiah for col in num_cols}
    format_dict.update({"Gap 1 to 2 (%)": "{:.1f}%"})
    for vendor in vendor_cols:
        format_dict[f"{vendor} to Median (%)"] = "{:+.1f}%"

    df_filtered_transposed_styled = (
        df_filtered_transposed.style
        .format(format_dict)
        .apply(lambda row: highlight_1st_2nd_vendor(row, df_filtered_transposed.columns), axis=1)
    )

    tab2.markdown(
        f"""
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
            <div style="font-size:0.9rem; color:gray;">
                ✨ Total number of data entries: <b>{len(df_filtered_transposed)}</b>
            </div>
            <div style="text-align:right;">
                <span style="background:#C6EFCE; padding:2px 8px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">1st Lowest</span>
                &nbsp;
                <span style="background:#FFEB9C; padding:2px 8px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">2nd Lowest</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    tab2.dataframe(df_filtered_transposed_styled, hide_index=True)

    excel_data = get_excel_download_highlight_1st_2nd_lowest(df_filtered_transposed)
    with tab2:
        col1, col2, col3 = st.columns([2.3,2,1])
        with col3:
            st.download_button(
                label="Download",
                data=excel_data,
                file_name="Bid & Price Analysis Transposed - TCO by Region.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                icon=":material/download:",
            )

    # VISUALIZATIONN
    tab2.markdown("##### 📊 Visualization")

    with tab2:
        subtab1, subtab2 = st.tabs(["Win Rate Trend", "Average Gap Trend"])

        with subtab1:
            # --- WIN RATE VISUALIZATION ---
            # --- Hitung total kemenangan (1st & 2nd Vendor)
            win1_counts = df_analysis_transposed["1st Vendor"].value_counts(dropna=True).reset_index()
            win1_counts.columns = ["Vendor", "Wins1"]

            win2_counts = df_analysis_transposed["2nd Vendor"].value_counts(dropna=True).reset_index()
            win2_counts.columns = ["Vendor", "Wins2"]

            # --- Hitung total partisipasi vendor ---
            vendor_counts = (
                (df_analysis[vendor_cols].fillna(0) > 0)   # hanya True jika nilai > 0
                .sum()         # Hitung True per kolom
                .reset_index()
            )
            vendor_counts.columns = ["Vendor", "Total Participations"]

            # --- Gabungkan semua ---
            win_rate = (
                vendor_counts
                .merge(win1_counts, on="Vendor", how="left")
                .merge(win2_counts, on="Vendor", how="left")
                .fillna(0)
            )

            # --- Hitung Win Rate (%)
            win_rate["1st Win Rate (%)"] = np.where(
                win_rate["Total Participations"] > 0,
                (win_rate["Wins1"] / win_rate["Total Participations"] * 100).round(1),
                0
            )
            win_rate["2nd Win Rate (%)"] = np.where(
                win_rate["Total Participations"] > 0,
                (win_rate["Wins2"] / win_rate["Total Participations"] * 100).round(1),
                0
            )

            # --- Siapkan data long-format untuk visualisasi ---
            win_rate_long = win_rate.melt(
                id_vars=["Vendor"],
                value_vars=["1st Win Rate (%)", "2nd Win Rate (%)"],
                var_name="Metric",
                value_name="Percentage"
            )

            # --- Urutkan vendor berdasarkan 1st Win Rate tertinggi ---
            vendor_order = (
                win_rate.sort_values("1st Win Rate (%)", ascending=False)["Vendor"].tolist()
            )

            # --- Tentukan warna untuk kedua metrik ---
            metric_colors = {
                "1st Win Rate (%)": "#1f77b4",
                "2nd Win Rate (%)": "#ff7f0e"     # biru
            }

            # --- Highlight interaktif ---
            highlight = alt.selection_point(on='mouseover', fields=['Metric'], nearest=True)

            # --- Batas atas dan bawah sumbu Y ---
            y_max = win_rate_long["Percentage"].max()
            if not np.isfinite(y_max) or y_max <= 0:
                y_max = 1

            # Pastikan data diurutkan sesuai vendor_order
            win_rate_long["Vendor"] = pd.Categorical(win_rate_long["Vendor"], categories=vendor_order, ordered=True)
            win_rate_long = win_rate_long.sort_values(["Metric", "Vendor"])

            # --- Chart utama ---
            base = (
                alt.Chart(win_rate_long)
                .encode(
                    x=alt.X("Vendor:N", sort=vendor_order, title=None),
                    y=alt.Y(
                        "Percentage:Q",
                        title="Win Rate (%)",
                        scale=alt.Scale(domain=[0, y_max * 1.2])
                    ),
                    color=alt.Color(
                        "Metric:N",
                        scale=alt.Scale(
                            domain=list(metric_colors.keys()),
                            range=list(metric_colors.values())
                        ),
                        title="Rank"
                    ),
                    tooltip=[
                        alt.Tooltip("Vendor:N", title="Vendor"),
                        alt.Tooltip("Metric:N", title="Position"),
                        alt.Tooltip("Percentage:Q", title="Win Rate (%)", format=".1f")
                    ]
                )
            )

            # --- Garis dengan titik ---
            lines = base.mark_line(point=alt.OverlayMarkDef(size=70, filled=True), strokeWidth=3)

            # --- Label persentase di atas titik ---
            labels = base.mark_text(
                align='center',
                baseline='bottom',
                dy=-7,
                fontWeight='bold',
                color='gray'
            ).encode(
                text=alt.Text("Percentage:Q", format=".1f")
            ).transform_calculate(
                label="format(datum.Percentage, '.1f') + '%'"
            ).encode(
                text="label:N"
            )

            # --- Gabungkan semua elemen ---
            chart = (
                lines + labels
            ).properties(
                height=400,
                padding={"right": 15},
                title="Vendor Win Rate Comparison (1st vs 2nd Place)"
            ).configure_title(
                anchor="middle",
                offset=12,
                fontSize=14
            ).configure_axis(
                labelFontSize=12,
                titleFontSize=13
            ).configure_view(
                stroke='gray',
                strokeWidth=1
            ).configure_legend(
                titleFontSize=12,
                titleFontWeight="bold",
                labelFontSize=12,
                labelLimit=300,
                orient="bottom"
            )

            st.write("")
            # --- Tampilkan chart di Streamlit
            st.altair_chart(chart)

            # Kolom yang mau ditaruh di depan
            cols_front = ["Wins1", "Wins2"]

            # Sisanya (Vendor + kolom lain yang tidak ada di cols_front)
            cols_rest = [c for c in win_rate.columns if c not in cols_front]

            # Gabungkan urutannya
            win_rate = win_rate[cols_rest[:1] + cols_front + cols_rest[1:]]

            # --- Ganti nama kolom biar lebih konsisten & enak dibaca
            df_summary = win_rate.rename(columns={
                "Wins1": "1st Rank",
                "Wins2": "2nd Rank"
            })

            with st.expander("See explanation"):
                st.write('''
                    The visualization above compares the win rate of each vendor
                    based on how often they achieved 1st or 2nd place in all
                    tender evaluations.  
                            
                    **💡 How to interpret the chart**  
                            
                    - High 1st Win Rate (%)  
                        Vendor is highly competitive and often offers the best commercial terms.  
                    - High 2nd Win Rate (%)  
                        Vendor consistently performs well, often just slightly less competitive than the winner.  
                    - Large Gap Between 1st & 2nd Win Rate  
                        Shows clear market dominance by certain vendors.
                ''')

                num_cols = df_summary.select_dtypes(include=["number"]).columns
                format_dict = {col: format_rupiah for col in num_cols}
                format_dict.update({
                    "1st Win Rate (%)": "{:.1f}%",
                    "2nd Win Rate (%)": "{:.1f}%"
                })

                df_summary_styled = df_summary.style.format(format_dict)

                st.dataframe(df_summary_styled, hide_index=True)

                # Simpan hasil ke variabel
                excel_data = get_excel_download(df_summary)

                # Layout tombol (rata kanan)
                col1, col2, col3 = st.columns([3,1,1])
                with col3:
                    st.download_button(
                        label="Download",
                        data=excel_data,
                        file_name="Win Rate Trend Summary.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        icon=":material/download:",
                        key="download_Win_Rate_Trend_Tab2"
                    )

        with subtab2:
            # --- AVERAGE GAP VISUALIZATION ---
            # --- Hitung rata-rata Gap 1 to 2 (%) per Vendor (hanya untuk 1st Vendor)
            df_gap = df_analysis_transposed.copy()

            # Ubah kolom 'Gap 1 to 2 (%)' ke numerik (hapus simbol %)
            df_gap["Gap 1 to 2 (%)"] = (
                df_gap["Gap 1 to 2 (%)"]
                .replace("%", "", regex=True)
                .astype(float)
            )

            # Hitung rata-rata gap per vendor (hanya vendor yang jadi 1st Lowest)
            avg_gap = (
                df_gap.groupby("1st Vendor", dropna=True)["Gap 1 to 2 (%)"]
                .mean()
                .reset_index()
                .rename(columns={"1st Vendor": "Vendor", "Gap 1 to 2 (%)": "Average Gap (%)"})
                .sort_values("Average Gap (%)", ascending=False)
            )

            # Warna per vendor (biar konsisten kalau kamu sudah punya color mapping)
            colors_list = ["#F94144", "#F3722C", "#F8961E", "#F9C74F", "#90BE6D",
                        "#43AA8B", "#577590", "#E54787", "#BF219A", "#8E0F9C", "#4B1D91"]
            vendor_colors = {v: c for v, c in zip(avg_gap["Vendor"].unique(), colors_list)}

            # Interaksi hover
            highlight = alt.selection_point(on='mouseover', fields=['Vendor'], nearest=True)

            # Pastikan tidak ada inf/NaN pada max
            y_max = avg_gap["Average Gap (%)"].replace([np.inf, -np.inf], np.nan).max()
            if pd.isna(y_max):
                y_max = 0

            # --- Chart utama ---
            bars = (
                alt.Chart(avg_gap)
                .mark_bar()
                .encode(
                    x=alt.X("Vendor:N", sort='-y', title=None),
                    # y=alt.Y("Average Gap (%):Q", title="Average Gap (%)", scale=alt.Scale(domain=[0, y_max * 1.2])),
                    y=alt.Y("Average Gap (%):Q", title="Average Gap (%)", scale=alt.Scale(domain=[0, avg_gap["Average Gap (%)"].max() * 1.2])),
                    color=alt.Color("Vendor:N",
                                    scale=alt.Scale(domain=list(vendor_colors.keys()), range=list(vendor_colors.values())),
                                    legend=None),
                    tooltip=[
                        alt.Tooltip("Vendor:N", title="Vendor"),
                        alt.Tooltip("Average Gap (%):Q", title="Average Gap (%)", format=".1f")
                    ]
                )
                .add_params(highlight)
            )

            # Label teks di atas bar
            labels = (
                alt.Chart(avg_gap)
                .mark_text(dy=-7, fontWeight='bold', color='gray')
                .encode(
                    x="Vendor:N",
                    y="Average Gap (%):Q",
                    text=alt.Text("Average Gap (%):Q", format=".1f")  # Format angka
                )
                .transform_calculate(  # Tambahkan simbol %
                    label_text="format(datum['Average Gap (%)'], '.1f') + '%'"
                )
                .encode(
                    text="label_text:N"
                )
            )

            # Frame luar untuk gaya rapi
            frame = alt.Chart().mark_rect(stroke='gray', strokeWidth=1, fillOpacity=0)

            avg_line = alt.Chart(avg_gap).mark_rule(color='gray', strokeDash=[4,2], size=1.75).encode(
                y='mean(Average Gap (%)):Q'
            )

            # Gabungkan semua elemen
            chart = (bars + labels + frame + avg_line).properties(
                title="Average Gap (%) per 1st Vendor",
                height=400
            ).configure_title(
                anchor='middle',
                offset=12,
                fontSize=14
            ).configure_axis(
                labelFontSize=12,
                titleFontSize=13
            ).configure_view(
                stroke='gray',
                strokeWidth=1
            )

            st.write("")
            # --- Tampilkan di Streamlit ---
            st.altair_chart(chart)

            avg_value = avg_gap["Average Gap (%)"].mean()

            # with st:
            with st.expander("See explanation"):
                st.write(f'''
                    The chart above shows the average price difference between 
                    the lowest and second-lowest bids for each vendor when they 
                    rank 1st, indicating their pricing dominance or competitiveness.
                            
                    **💡 How to interpret the chart**  
                            
                    - High Gap  
                        High gap indicates strong vendor dominance (much lower prices).  
                    - Low Gap  
                        Low gap indicates intense competition with similar pricing among vendors.  
                    
                    The dashed line represents the average gap across all vendors, serving as a benchmark ({avg_value:.1f}%).
                ''')

    tab2.divider()

    # SUPERRR BUTTOONNN 2
    tab2.markdown("##### 🧑‍💻 Super Download — Export Selected Sheets")
    dataframes = {
        "Merge Transposed": df_merge_transposed,
        "TCO Summary Transposed": df_tco_transposed,
        "Bid & Price Analysis Transposed": df_filtered_transposed,
    }

    # Tampilkan multiselect
    selected_sheets = tab2.multiselect(
        "Select sheets to download in a single Excel file:",
        options=list(dataframes.keys()),
        default=list(dataframes.keys())  # default semua dipilih
    )

    # Fungsi "Super Button" & Formatting
    def generate_multi_sheet_excel_transposed(selected_sheets, df_dict):
        output = BytesIO()

        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            for sheet in selected_sheets:
                df = df_dict[sheet]
                df.to_excel(writer, index=False, sheet_name=sheet)

                workbook  = writer.book
                worksheet = writer.sheets[sheet]

                # ===== FORMAT =====
                fmt_rp   = workbook.add_format({'num_format': '#,##0'})
                fmt_pct  = workbook.add_format({'num_format': '#,##0.0"%"'})
                fmt_bold = workbook.add_format({'bold': True, 'num_format': '#,##0'})

                fmt_total = workbook.add_format({
                    'bold': True,
                    'bg_color': '#D9EAD3',
                    'font_color': '#1A5E20',
                    'num_format': '#,##0'
                })

                fmt_1 = workbook.add_format({
                    'bg_color': '#C6EFCE',
                    'font_color': '#006100',
                    'num_format': '#,##0'
                })

                fmt_2 = workbook.add_format({
                    'bg_color': '#FFEB9C',
                    'font_color': '#9C6500',
                    'num_format': '#,##0'
                })

                fmt_1b = workbook.add_format({
                    'bg_color': '#C6EFCE',
                    'font_color': '#006100',
                    'bold': True,
                    'num_format': '#,##0'
                })

                fmt_2b = workbook.add_format({
                    'bg_color': '#FFEB9C',
                    'font_color': '#9C6500',
                    'bold': True,
                    'num_format': '#,##0'
                })

                num_cols = df.select_dtypes(include=["number"]).columns.tolist()
                pct_cols = [c for c in df.columns if "%" in c]

                # ===== LOOP DATA =====
                for r, (_, row) in enumerate(df.iterrows(), start=1):
                    is_total = any(str(x).strip().upper() == "TOTAL" for x in row)

                    first = second = None

                    # ===== SUMMARY RANKING (EXCLUDE ZERO) =====
                    if sheet == "TCO Summary Transposed":
                        numeric_vals = row[num_cols]
                        numeric_vals = numeric_vals[
                            (numeric_vals.notna()) & (numeric_vals != 0)
                        ]

                        if not numeric_vals.empty:
                            sorted_vals = numeric_vals.sort_values()
                            first = sorted_vals.index[0]
                            if len(sorted_vals) > 1:
                                second = sorted_vals.index[1]

                    # ===== BID & PRICE =====
                    if sheet == "Bid & Price Analysis Transposed":
                        first = row.get("1st Vendor")
                        second = row.get("2nd Vendor")

                    for c, col in enumerate(df.columns):
                        val = row[col]

                        # ===== SAFETY =====
                        if pd.isna(val) or (isinstance(val, (int, float)) and np.isinf(val)):
                            worksheet.write(r, c, "")
                            continue

                        fmt = None

                        is_zero = isinstance(val, (int, float)) and val == 0
                        # ===== NO HIGHLIGHT FOR ZERO (EXCEPT MERGE DATA) =====
                        if is_zero and sheet != "Merge Transposed":
                            fmt = None

                        # ===== PICK FORMAT =====
                        elif col == first:
                            fmt = fmt_1b if is_total else fmt_1
                        elif col == second:
                            fmt = fmt_2b if is_total else fmt_2
                        elif is_total and sheet == "Merge Transposed":
                            fmt = fmt_total
                        elif is_total:
                            fmt = fmt_bold

                        # ===== WRITE CELL =====
                        if col in pct_cols:
                            worksheet.write_number(r, c, val, fmt or fmt_pct)
                        elif col in num_cols:
                            worksheet.write_number(r, c, val, fmt or fmt_rp)
                        else:
                            worksheet.write(r, c, val, fmt)

                    # ===== AUTOFIT =====
                    for i, col in enumerate(df.columns):
                        worksheet.set_column(
                            i, i,
                            max(len(str(col)), df[col].astype(str).map(len).max()) + 2
                        )

        output.seek(0)
        return output.getvalue()

    # --- FRAGMENT UNTUK BALLOONS ---
    @st.fragment
    def release_the_balloons():
        tab2.balloons()

    # ---- DOWNLOAD BUTTON ----
    if selected_sheets:
        excel_bytes = generate_multi_sheet_excel_transposed(selected_sheets, dataframes)

        tab2.download_button(
            label="Download",
            data=excel_bytes,
            file_name="TCO Comparison by Region - Transposed Data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            on_click=release_the_balloons,
            type="primary",
            use_container_width=True,
        )
