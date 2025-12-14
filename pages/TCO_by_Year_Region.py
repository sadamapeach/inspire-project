import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
import time
import re
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
    if str(row.iloc[0]).strip().upper() == "TOTAL":
        return ["font-weight: bold;"] * len(row)
    else:
        return [""] * len(row)
    
def highlight_total_row_v2(row):
    # Cek apakah ada kolom yang berisi "TOTAL" (case-insensitive)
    if any(str(x).strip().upper() == "TOTAL" for x in row):
        return ["font-weight: bold; background-color: #D9EAD3; color: #1A5E20;"] * len(row)
    else:
        return [""] * len(row)

# Highlight total per year
def highlight_total_per_year(row):
    if str(row["SCOPE"]).strip().upper() == "TOTAL" and pd.notna(row["YEAR"]):
        return ["font-weight: bold; background-color: #FFEB9C; color: #9C6500;"] * len(row)
    else:
        return [""] * len(row)

# Highlight vendor total
def highlight_vendor_total(row):
    if str(row["YEAR"]).strip().upper() == "TOTAL":
        return ["font-weight: bold; background-color: #C6EFCE; color: #006100;"] * len(row)
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

# Download highlight total (kuning + hijau) -> df_merge
@st.cache_data
def get_excel_download_with_highlight(df, sheet_name="Sheet1"):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)

        workbook  = writer.book
        worksheet = writer.sheets[sheet_name]

        # ================= FORMAT =================
        fmt_rupiah = workbook.add_format({"num_format": "#,##0"})

        fmt_total_year = workbook.add_format({
            "bold": True,
            "bg_color": "#FFEB9C",   # kuning
            "font_color": "#9C6500",
            "num_format": "#,##0"
        })

        fmt_vendor_total = workbook.add_format({
            "bold": True,
            "bg_color": "#C6EFCE",   # hijau
            "font_color": "#006100",
            "num_format": "#,##0"
        })

        fmt_text_bold = workbook.add_format({"bold": True})

        # ================= COLUMN GROUP =================
        num_cols = df.select_dtypes(include=["number"]).columns.tolist()

        # ================= REWRITE CELLS =================
        for row_idx, row in enumerate(df.itertuples(index=False), start=1):

            scope_val = str(row[2]).strip().upper()  # kolom Scope
            year_val  = str(row[1]).strip().upper()  # kolom Year

            # Tentukan format baris
            row_fmt = None
            if scope_val == "TOTAL" and year_val != "TOTAL":
                row_fmt = fmt_total_year
            elif year_val == "TOTAL":
                row_fmt = fmt_vendor_total

            for col_idx, col_name in enumerate(df.columns):
                val = row[col_idx]

                # ===== SAFETY =====
                if pd.isna(val) or (isinstance(val, (int, float)) and np.isinf(val)):
                    worksheet.write(row_idx, col_idx, "")
                    continue

                # ===== NUMERIC COLUMN =====
                if col_name in num_cols:
                    worksheet.write_number(
                        row_idx,
                        col_idx,
                        val,
                        row_fmt or fmt_rupiah
                    )

                # ===== TEXT COLUMN =====
                else:
                    worksheet.write(
                        row_idx,
                        col_idx,
                        val,
                        row_fmt if row_fmt else None
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

# Download highlight total ver2 (kuning + hijau) -> df_cost_summary
@st.cache_data
def get_excel_download_with_highlight_v2(df, sheet_name="Sheet1"):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)

        workbook  = writer.book
        worksheet = writer.sheets[sheet_name]

        # ================= FORMAT =================
        fmt_rupiah = workbook.add_format({"num_format": "#,##0"})

        fmt_total_year = workbook.add_format({
            "bold": True,
            "bg_color": "#FFEB9C",   # kuning
            "font_color": "#9C6500",
            "num_format": "#,##0"
        })

        fmt_vendor_total = workbook.add_format({
            "bold": True,
            "bg_color": "#C6EFCE",   # hijau
            "font_color": "#006100",
            "num_format": "#,##0"
        })

        # ================= COLUMN GROUP =================
        num_cols = df.select_dtypes(include=["number"]).columns.tolist()

        # Deteksi kolom dinamis
        year_col = next((c for c in df.columns if "YEAR" in c.upper()), None)
        scope_col = next((c for c in df.columns if "SCOPE" in c.upper()), None)

        year_idx = df.columns.get_loc(year_col) if year_col else None
        scope_idx = df.columns.get_loc(scope_col) if scope_col else None

        # ================= REWRITE CELLS =================
        for row_idx, row in enumerate(df.itertuples(index=False), start=1):

            year_val = (
                str(row[year_idx]).strip().upper()
                if year_idx is not None else ""
            )
            scope_val = (
                str(row[scope_idx]).strip().upper()
                if scope_idx is not None else ""
            )

            # Tentukan format baris
            row_fmt = None
            if scope_val == "TOTAL" and year_val != "TOTAL":
                row_fmt = fmt_total_year
            elif year_val == "TOTAL":
                row_fmt = fmt_vendor_total

            for col_idx, col_name in enumerate(df.columns):
                val = row[col_idx]

                # ===== SAFETY =====
                if pd.isna(val) or (isinstance(val, (int, float)) and np.isinf(val)):
                    worksheet.write(row_idx, col_idx, "")
                    continue

                # ===== NUMERIC COLUMN =====
                if col_name in num_cols:
                    worksheet.write_number(
                        row_idx,
                        col_idx,
                        val,
                        row_fmt or fmt_rupiah
                    )

                # ===== TEXT COLUMN =====
                else:
                    worksheet.write(
                        row_idx,
                        col_idx,
                        val,
                        row_fmt
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
            3️⃣ TCO Comparison by Year + Region
        </div>
        """,
        unsafe_allow_html=True
    )
    # st.markdown("##3️⃣ TCO Comparison by Year + Region")
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
        st.session_state["uploaded_file_tco_by_year_region"] = upload_file

        # --- Animasi proses upload ---
        msg = st.toast("📂 Uploading file...")
        time.sleep(1.2)

        msg.toast("🔍 Reading sheets...")
        time.sleep(1.2)

        msg.toast("✅ File uploaded successfully!")
        time.sleep(0.5)

        # Baca semua sheet sekaligus
        all_df = pd.read_excel(upload_file, sheet_name=None)
        st.session_state["all_df_tco_by_year_region_raw"] = all_df  # simpan versi mentah

    elif "all_df_tco_by_year_region_raw" in st.session_state:
        all_df = st.session_state["all_df_tco_by_year_region_raw"]
    else:
        return
    
    st.divider()

    # OVERVIEW
    # st.markdown("##### 🔍 Overview")
    # Total bidders
    total_sheets = len(all_df)
    # st.caption(f"You're analyzing offers from **{total_sheets} participating bidders** in this session 🧐")

    result = {}

    for name, df in all_df.items():
        rows_before, cols_before = df.shape
        # Data cleaning
        df_clean = df.replace(r'^\s*$', None, regex=True)
        df_clean = df_clean.dropna(how="all", axis=0).dropna(how="all", axis=1)

        rows_after, cols_after = df_clean.shape

        # Gunakan baris pertama sebagai header (hanya jika kolom belum ada nama atau Unnamed)
        if any("Unnamed" in str(c) for c in df_clean.columns):
            df_clean.columns = df_clean.iloc[0]
            df_clean = df_clean[1:].reset_index(drop=True)

        # --- Pastikan kolom pertama (Year) tetap sebagai teks ---
        first_col = df_clean.columns[0]  # Kolom Year
        df_clean[first_col] = df_clean[first_col].astype(str)

        # --- Konversi tipe data otomatis untuk kolom lainnya ---
        df_clean.iloc[:, 1:] = df_clean.iloc[:, 1:].convert_dtypes()

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

        # --- Tambah kolom total (kecuali Year) ---
        if "TOTAL" not in df_clean.columns:
            df_clean["TOTAL"] = df_clean.iloc[:, 1:].sum(axis=1, numeric_only=True)

        # Pembulatan
        num_cols = df_clean.select_dtypes(include=["number"]).columns
        df_clean[num_cols] = df_clean[num_cols].apply(round_half_up)

        # Format Rupiah
        df_styled = df_clean.style.format({col: format_rupiah for col in num_cols})

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

    st.session_state["result_tco_by_year_region"] = result
    # st.divider()

    # MERGEE
    st.markdown("##### 🗃️ Merge Data")

    merged_list = []

    for vendor_name, df_vendor in result.items():
        df_temp = df_vendor.copy()

        # Identifikasi kolom
        non_num_cols = df_temp.select_dtypes(exclude=["number"]).columns.tolist()
        num_cols     = df_temp.select_dtypes(include=["number"]).columns.tolist()

        # Ambil dua pertama untuk grouping utama (year & scope)
        year_col  = non_num_cols[0]
        scope_col = non_num_cols[1]

        # Tambahkan kolom VENDOR
        df_temp.insert(0, "VENDOR", vendor_name)

        vendor_result = []

        # --- Loop tiap YEAR (dynamic)
        for year, group_year in df_temp.groupby(year_col, dropna=False):

            # --- Row TOTAL per year
            total_row = {
                "VENDOR": vendor_name,
                year_col: year,
                scope_col: "TOTAL"
            }

            # Untuk kolom non-numeric lainnya (selain year & scope), isi kosong
            for col in non_num_cols[2:]:
                total_row[col] = ""

            # Numeric sum
            for col in num_cols:
                total_row[col] = group_year[col].sum(numeric_only=True)

            tco_year_with_total = pd.concat(
                [group_year, pd.DataFrame([total_row])],
                ignore_index=True
            )

            vendor_result.append(tco_year_with_total)

        # --- Gabungkan semua year untuk vendor ini
        df_vendor_with_year_total = pd.concat(vendor_result, ignore_index=True)

        # --- TOTAL BESAR per vendor → hanya dari row TOTAL tiap year
        total_rows = df_vendor_with_year_total[df_vendor_with_year_total[scope_col] == "TOTAL"]

        vendor_total = {"VENDOR": vendor_name}

        vendor_total[year_col] = "TOTAL"
        vendor_total[scope_col] = ""

        # Kosongkan semua non_num selain year & scope
        for col in non_num_cols[2:]:
            vendor_total[col] = ""

        # Sum semua numeric
        for col in num_cols:
            vendor_total[col] = total_rows[col].sum(numeric_only=True)

        # Gabungkan ke dataframe vendor
        df_vendor_final = pd.concat(
            [df_vendor_with_year_total, pd.DataFrame([vendor_total])],
            ignore_index=True
        )

        merged_list.append(df_vendor_final)

    # --- Gabungkan semua vendor ---
    df_merge = pd.concat(merged_list, ignore_index=True)

    # --- Urutkan kolom supaya rapi: VENDOR + non-num + num-col ---
    final_cols = ["VENDOR"] + non_num_cols + num_cols
    df_merge = df_merge[final_cols]

    # --- Styling (opsional) ---
    num_cols = df_merge.select_dtypes(include=["number"]).columns
    df_merge_styled = (
        df_merge.style
        .format({col: format_rupiah for col in num_cols})
        .apply(highlight_total_per_year, axis=1)
        .apply(highlight_vendor_total, axis=1)
    )

    st.caption(f"Data from **{total_sheets} vendors** have been successfully consolidated, analyzing a total of **{len(df_merge):,} combined records**.")
    st.dataframe(df_merge_styled, hide_index=True)

    # Download button to Excel
    excel_data = get_excel_download_with_highlight(df_merge)
    col1, col2, col3 = st.columns([2.3,2,1])
    with col3:
        st.download_button(
            label="Download",
            data=excel_data,
            file_name="Merge Data - TCO by Year Region.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            icon=":material/download:",
        )
    
    st.divider()

    # COST SUMMARY
    st.markdown("##### 📑 Cost Summary")

    # --- Identifikasi kolom non-numeric dan numeric ---
    non_num_cols = df_merge.select_dtypes(exclude=["number"]).columns.tolist()
    numeric_cols = df_merge.select_dtypes(include=["number"]).columns.tolist()

    vendor_col = non_num_cols[0]
    year_col   = non_num_cols[1]
    scope_col  = non_num_cols[2]

    # --- Sisa non-number (dinamis) ---
    other_non_num = [c for c in non_num_cols if c not in [vendor_col, year_col, scope_col]]

    # --- Region columns = semua numeric kecuali kolom 'TOTAL' ---
    region_cols = [c for c in numeric_cols if c.upper() != "TOTAL"]

    # --- Transform to long format (Vendor, Year, Region, Scope, Price)
    df_cost_summary = df_merge.melt(
        id_vars=[vendor_col, year_col, scope_col] + other_non_num,
        value_vars=region_cols,
        var_name="REGION",
        value_name="[PRICE]"
    )

    # --- Rapikan urutan kolom ---
    final_cols = (
        [vendor_col, year_col, "REGION", scope_col] 
        + other_non_num 
        + ["[PRICE]"]
    )

    df_cost_summary = df_cost_summary[final_cols]

    # Simpan ke session_state jika perlu
    st.session_state["merged_long_format_total_price"] = df_cost_summary

    # Format Rupiah untuk kolom PRICE
    df_cost_summary_styled = (
        df_cost_summary.style
        .format({"[PRICE]": format_rupiah})
        .apply(highlight_total_per_year, axis=1)
        .apply(highlight_vendor_total, axis=1)
    )

    # Tampilkan
    st.caption(f"Consolidated cost summary containing **{len(df_cost_summary):,} records** across multiple vendors and regions.")
    st.dataframe(df_cost_summary_styled, hide_index=True)

    # Download button to Excel
    excel_data = get_excel_download_with_highlight_v2(df_cost_summary)
    col1, col2, col3 = st.columns([2.3,2,1])
    with col3:
        st.download_button(
            label="Download",
            data=excel_data,
            file_name="Cost Summary - TCO by Year Region.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            icon=":material/download:",
        )

    # TCOO
    st.markdown("##### 💸 TCO Summary")
    tab1, tab2, tab3 = st.tabs(["YEAR", "REGION", "SCOPE"])

    vendor_col = "VENDOR"
    year_col = df_cost_summary.columns[1]
    region_col = "REGION"
    scope_col = df_cost_summary.columns[3]
    price_col = "[PRICE]"

    # Tab1: YEAR
    # --- Hapus baris TOTAL agar tidak double count ---
    tco_year_clean = df_cost_summary[
        (df_cost_summary[year_col].astype(str).str.upper() != "TOTAL") &
        (df_cost_summary[scope_col].astype(str).str.upper() != "TOTAL")
    ]

    tco_year = tco_year_clean.pivot_table(
        index=year_col,
        columns="VENDOR",
        values="[PRICE]",
        aggfunc="sum",
        fill_value=0
    ).reset_index()

    total_row = pd.DataFrame({
        year_col: ["TOTAL"], 
        **{col: [tco_year[col].sum()] for col in tco_year.columns if col != year_col}
    })
    tco_year = pd.concat([tco_year, total_row], ignore_index=True)

    num_cols_year = tco_year.select_dtypes(include=["number"]).columns

    # Format rupiah, exclude kolom pertama (YEAR)
    format_dict = {col: format_rupiah for col in tco_year.columns[1:]}  # skip kolom pertama

    tco_year_styled = (
        tco_year.style
        .format(format_dict)  # hanya format kolom numeric
        .apply(highlight_total_row, axis=1)
        .apply(lambda row: highlight_rank_summary(row, num_cols_year), axis=1)
    )

    tab1.markdown(
        """
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px;">
            <div style="font-size:0.88rem; color:gray; font-weight:400;">
                Overview of total vendor costs per year.
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

    tab1.dataframe(tco_year_styled, hide_index=True)

    excel_data = get_excel_download_highlight_summary(tco_year)
    with tab1:
        col1, col2, col3 = st.columns([2.3,2,1])
        with col3:
            st.download_button(
                label="Download",
                data=excel_data,
                file_name="TCO Summary (Year) - TCO by Year Region.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                icon=":material/download:",
                key="download_year"
            )

    # Tab2: REGION
    # --- Hapus baris TOTAL agar tidak double count ---
    tco_region_clean = df_cost_summary[
        (df_cost_summary[year_col].astype(str).str.upper() != "TOTAL") &
        (df_cost_summary[scope_col].astype(str).str.upper() != "TOTAL")
    ]

    # --- Buat pivot ---
    tco_region = tco_region_clean.pivot_table(
        index=region_col,
        columns=vendor_col,
        values="[PRICE]",
        aggfunc="sum",
        fill_value=0
    ).reset_index()

    # --- Tambahkan baris TOTAL ---
    total_row = pd.DataFrame({
        region_col: ["TOTAL"],
        **{col: [tco_region[col].sum()] for col in tco_region.columns if col != region_col}
    })

    tco_region = pd.concat([tco_region, total_row], ignore_index=True)

    num_cols_region = tco_region.select_dtypes(include=["number"]).columns

    tco_region_styled = (
        tco_region.style
        .format(format_rupiah)
        .apply(highlight_total_row, axis=1)
        .apply(lambda row: highlight_rank_summary(row, num_cols_region), axis=1)
    )

    tab2.markdown(
        """
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px;">
            <div style="font-size:0.88rem; color:gray; font-weight:400;">
                Breakdown of vendor costs across regions.
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

    tab2.dataframe(tco_region_styled, hide_index=True)

    excel_data = get_excel_download_highlight_summary(tco_region)
    with tab2:
        col1, col2, col3 = st.columns([2.3,2,1])
        with col3:
            st.download_button(
                label="Download",
                data=excel_data,
                file_name="TCO Summary (Region) - TCO by Year Region.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                icon=":material/download:",
                key="download_region"
            )

    # Tab3: Scope
    tco_scope_clean = df_cost_summary[
        (df_cost_summary[year_col].astype(str).str.upper() != "TOTAL") &
        (df_cost_summary[scope_col].astype(str).str.upper() != "TOTAL")
    ]

    tco_scope = tco_scope_clean.pivot_table(
        index=scope_col,
        columns=vendor_col,
        values=price_col,
        aggfunc="sum",
        fill_value=0
    ).reset_index()

    total_row = pd.DataFrame({
        scope_col: ["TOTAL"],
        **{col: [tco_scope[col].sum()] for col in tco_scope.columns if col != scope_col}
    })

    tco_scope = pd.concat([tco_scope, total_row], ignore_index=True)

    num_cols_scope = tco_scope.select_dtypes(include=["number"]).columns

    tco_scope_styled = (
        tco_scope.style
        .format(format_rupiah)
        .apply(highlight_total_row, axis=1)
        .apply(lambda row: highlight_rank_summary(row, num_cols_scope), axis=1)
    )

    tab3.markdown(
        """
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px;">
            <div style="font-size:0.88rem; color:gray; font-weight:400;">
                Summary of vendor costs by project scope.
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

    tab3.dataframe(tco_scope_styled, hide_index=True)

    excel_data = get_excel_download_highlight_summary(tco_scope)
    with tab3:
        col1, col2, col3 = st.columns([2.3,2,1])
        with col3:
            st.download_button(
                label="Download",
                data=excel_data,
                file_name="TCO Summary (Scope) - TCO by Year Region.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                icon=":material/download:",
                key="download_scope"
            )

    st.divider()

    # --- ANALYTICAL COLUMNS ---
    st.markdown("##### 🧠 Bid & Price Analysis")

    # ---- COPY ----
    df_clean = df_merge.copy()

    # ---- IDENTIFIKASI KOLUMN ----
    non_num_cols = df_clean.select_dtypes(exclude=["number"]).columns.tolist()
    numeric_cols  = df_clean.select_dtypes(include=["number"]).columns.tolist()

    vendor_col = non_num_cols[0]          # contoh: VENDOR
    year_col   = non_num_cols[1]          # contoh: YEAR
    scope_col  = non_num_cols[2]          # contoh: SCOPE

    # kolom non-num tambahan seperti DESC, UOM, Category, dsb
    extra_non_num = non_num_cols[3:]      # boleh kosong

    # --- Ubah dari format wide (Region1, Region2, dst) ke long ---
    df_melted = df_clean.melt(
        id_vars=[vendor_col, year_col, scope_col] + extra_non_num,
        value_vars=[c for c in numeric_cols if c.upper() != "TOTAL"],
        var_name="REGION",
        value_name="[PRICE]"
    )

    df_melted["[PRICE]"] = pd.to_numeric(df_melted["[PRICE]"], errors="coerce").fillna(0)

    # --- Pivot untuk jadi format kolom per vendor ---
    df_pivot = df_melted.pivot_table(
        index=[year_col, "REGION", scope_col] + extra_non_num,
        columns=vendor_col,
        values="[PRICE]",
        aggfunc="sum",
        fill_value=0
    ).reset_index()

    # Identifikasi kolom
    vendor_cols = df_pivot.select_dtypes(include=["number"]).columns.tolist()

    # Hapus baris TOTAL untuk analisis per komponen
    df_no_total = df_pivot[
        ~df_pivot.apply(
            lambda row: row.astype(str).str.upper().eq("TOTAL").any(),
            axis=1
        )
    ].copy()

    # Penanganan untuk 0 value
    vendor_values = df_no_total[vendor_cols].copy()
    vendor_values = vendor_values.replace(0, pd.NA)
    vendor_values = vendor_values.apply(pd.to_numeric, errors="coerce")

    # Hitung nilai analisis per baris
    df_no_total["1st Lowest"] = vendor_values.min(axis=1)
    # df_no_total["1st Vendor"] = vendor_values.idxmin(axis=1)

    # Penanganan khusus 1st Vendor
    mask_all_nan1 = vendor_values.isna().all(axis=1)
    df_no_total["1st Vendor"] = None
    valid_idx1 = (~mask_all_nan1)
    df_no_total.loc[valid_idx1, "1st Vendor"] = vendor_values.loc[valid_idx1].idxmin(axis=1) 

    # Hitung 2nd Lowest
    # Hilangkan dulu nilai 1st Lowest dari kandidat (agar kita dapat 2nd Lowest yang benar)
    temp = vendor_values.mask(vendor_values.eq(df_no_total["1st Lowest"], axis=0))

    df_no_total["2nd Lowest"] = temp.min(axis=1)
    # df_no_total["2nd Vendor"] = temp.idxmin(axis=1)

    # Penanganan khusus 2nd Vendor
    mask_all_nan2 = temp.isna().all(axis=1)
    df_no_total["2nd Vendor"] = None
    valid_idx2 = (~mask_all_nan2)
    df_no_total.loc[valid_idx2, "2nd Vendor"] = temp.loc[valid_idx2].idxmin(axis=1)

    # Hitung Gap 1 to 2 (%)
    df_no_total["Gap 1 to 2 (%)"] = (
        (df_no_total["2nd Lowest"] - df_no_total["1st Lowest"])
        / df_no_total["1st Lowest"] * 100
    ).round(2)

    # Hitung median price
    df_no_total["Median Price"] = vendor_values.median(axis=1)
    df_no_total["Median Price"] = pd.to_numeric(df_no_total["Median Price"], errors="coerce")

    # Hitung deviasi tiap vendor terhadap median
    for v in vendor_cols:
        df_no_total[f"{v} to Median (%)"] = (
            (df_no_total[v] - df_no_total["Median Price"])
            / df_no_total["Median Price"] * 100
        ).round(2)

    # --- Urutkan kolom agar rapi ---
    summary_cols = (
        [year_col, "REGION", scope_col] +
        extra_non_num +
        vendor_cols +
        ["1st Lowest", "1st Vendor",
        "2nd Lowest", "2nd Vendor",
        "Gap 1 to 2 (%)", "Median Price"] +
        [f"{v} to Median (%)" for v in vendor_cols]
    )

    df_summary = df_no_total[summary_cols]

    # --- 🎯 Tambahkan slicer
    all_year = sorted(df_summary["YEAR"].dropna().unique())
    all_region = sorted(df_summary["REGION"].dropna().unique())
    all_1st = sorted(df_summary["1st Vendor"].dropna().unique())
    all_2nd = sorted(df_summary["2nd Vendor"].dropna().unique())

    col_sel_1, col_sel_2, col_sel_3, col_sel_4 = st.columns(4)
    with col_sel_1:
        selected_year = st.multiselect(
            "Filter: Year",
            options=all_year,
            default=None,
            placeholder="Choose years",
            key="filter_year"
        )
    with col_sel_2:
        selected_region = st.multiselect(
            "Filter: Region",
            options=all_region,
            default=None,
            placeholder="Choose region",
            key="filter_region"
        )
    with col_sel_3:
        selected_1st = st.multiselect(
            "Filter: 1st vendor",
            options=all_1st,
            default=None,
            placeholder="Choose vendors",
            key="filter_1st_vendor"
        )
    with col_sel_4:
        selected_2nd = st.multiselect(
            "Filter: 2nd vendor",
            options=all_2nd,
            default=None,
            placeholder="Choose vendors",
            key="filter_2nd_vendor"
        )

    # --- Terapkan filter AND secara dinamis
    df_filtered_analysis = df_summary.copy()

    if selected_year:
        df_filtered_analysis = df_filtered_analysis[df_filtered_analysis["YEAR"].isin(selected_year)]

    if selected_region:
        df_filtered_analysis = df_filtered_analysis[df_filtered_analysis["REGION"].isin(selected_region)]

    if selected_1st:
        df_filtered_analysis = df_filtered_analysis[df_filtered_analysis["1st Vendor"].isin(selected_1st)]

    if selected_2nd:
        df_filtered_analysis = df_filtered_analysis[df_filtered_analysis["2nd Vendor"].isin(selected_2nd)]

    # --- Format rupiah & persentase hanya untuk df_filtered_analysis
    num_cols = df_filtered_analysis.select_dtypes(include=["number"]).columns
    format_dict = {col: format_rupiah for col in num_cols}
    format_dict.update({"Gap 1 to 2 (%)": "{:.1f}%"})
    for v in vendor_cols:
        format_dict[f"{v} to Median (%)"] = "{:+.1f}%"

    df_summary_styled = (
        df_filtered_analysis.style
        .format(format_dict)
        .apply(lambda row: highlight_1st_2nd_vendor(row, df_filtered_analysis.columns), axis=1)
    )

    st.markdown(
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

    # --- Tampilkan hasil ---
    st.dataframe(df_summary_styled, hide_index=True)

    # Simpan hasil ke variabel
    excel_data = get_excel_download_highlight_1st_2nd_lowest(df_filtered_analysis)

    # Layout tombol (rata kanan)
    col1, col2, col3 = st.columns([2.3,2,1])
    with col3:
        st.download_button(
            label="Download",
            data=excel_data,
            file_name="Bid & Price Analysis - TCO by Year Region.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            icon=":material/download:"
        )

    # VISUALIZATION
    st.markdown("##### 📊 Visualization")
    tab1, tab2 = st.tabs(["Win Rate Trend", "Average Gap Trend"])

    # --- WIN RATE VISUALIZATION ---
    # --- Hitung total kemenangan (1st & 2nd Vendor)
    win1_counts = df_summary["1st Vendor"].value_counts(dropna=True).reset_index()
    win1_counts.columns = ["Vendor", "Wins1"]

    win2_counts = df_summary["2nd Vendor"].value_counts(dropna=True).reset_index()
    win2_counts.columns = ["Vendor", "Wins2"]

    # --- Hitung total partisipasi vendor ---
    vendor_counts = (
        (df_summary[vendor_cols].fillna(0) > 0)   # hanya True jika nilai > 0
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

    tab1.write("")
    # --- Tampilkan chart di Streamlit
    tab1.altair_chart(chart)

    # Kolom yang mau ditaruh di depan
    cols_front = ["Wins1", "Wins2"]

    # Sisanya (Vendor + kolom lain yang tidak ada di cols_front)
    cols_rest = [c for c in win_rate.columns if c not in cols_front]

    # Gabungkan urutannya
    win_rate = win_rate[cols_rest[:1] + cols_front + cols_rest[1:]]

    # --- Ganti nama kolom biar lebih konsisten & enak dibaca
    df_summary_chart = win_rate.rename(columns={
        "Wins1": "1st Rank",
        "Wins2": "2nd Rank"
    })

    with tab1:
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

            num_cols = df_summary_chart.select_dtypes(include=["number"]).columns
            format_dict = {col: format_rupiah for col in num_cols}
            format_dict.update({
                "1st Win Rate (%)": "{:.1f}%",
                "2nd Win Rate (%)": "{:.1f}%"
            })

            df_summary_chart_styled = df_summary_chart.style.format(format_dict)
            st.dataframe(df_summary_chart_styled, hide_index=True)

            # Simpan hasil ke variabel
            excel_data = get_excel_download(df_summary_chart)

            # Layout tombol (rata kanan)
            col1, col2, col3 = st.columns([3,1,1])
            with col3:
                st.download_button(
                    label="Download",
                    data=excel_data,
                    file_name="Win Rate Trend Summary.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    icon=":material/download:",
                    key=f"download_Win_Rate_Trend"
                )

    # --- AVERAGE GAP VISUALIZATION ---
    # --- Hitung rata-rata Gap 1 to 2 (%) per Vendor (hanya untuk 1st Vendor)
    df_gap = df_summary.copy()

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

    tab2.write("")
    # --- Tampilkan di Streamlit ---
    tab2.altair_chart(chart)

    avg_value = avg_gap["Average Gap (%)"].mean()

    with tab2:
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

    st.divider()

    # SUPERRR BUTTONN
    st.markdown("##### 🧑‍💻 Super Download — Export Selected Sheets")
    dataframes = {
        "Merge Data": df_merge,
        "Cost Summary": df_cost_summary,
        "TCO Summary (Year)": tco_year,
        "TCO Summary (Region)": tco_region,
        "TCO Summary (Scope)": tco_scope,
        "Bid & Price Analysis": df_filtered_analysis,
    }

    # Tampilkan multiselect
    selected_sheets = st.multiselect(
        "Select sheets to download in a single Excel file:",
        options=list(dataframes.keys()),
        default=list(dataframes.keys())  # default semua dipilih
    )

    # Fungsi "Super Button" & Formatting
    def generate_multi_sheet_excel(selected_sheets, df_dict):
        """
        Gabungkan sheet dengan highlight khusus per sheet:
        - Merge Data / Cost Summary: highlight TOTAL per year & TOTAL vendor
        - Bid & Price Analysis: highlight 1st & 2nd vendor + TOTAL
        - TCO sheets: highlight TOTAL
        Semua akses baris pakai index, jadi aman untuk kolom dengan spasi/simbol.
        """
        output = BytesIO()

        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            for sheet in selected_sheets:
                df = df_dict[sheet].copy()
                workbook  = writer.book
                worksheet = workbook.add_worksheet(sheet)
                writer.sheets[sheet] = worksheet

                # --- Format umum ---
                fmt_rupiah = workbook.add_format({'num_format': '#,##0'})
                fmt_pct    = workbook.add_format({'num_format': '#,##0.0"%"'})
                fmt_total  = workbook.add_format({'bold': True, 'bg_color': '#D9EAD3', 'font_color': '#1A5E20', 'num_format': '#,##0'})
                fmt_first  = workbook.add_format({'bg_color': '#C6EFCE', "num_format": "#,##0"})
                fmt_second = workbook.add_format({'bg_color': '#FFEB9C', "num_format": "#,##0"})
                fmt_total_year  = workbook.add_format({'bold': True, 'bg_color': '#FFEB9C', 'font_color': '#1A1A1A', 'num_format': '#,##0'})
                fmt_total_vendor  = workbook.add_format({'bold': True, 'bg_color': '#C6EFCE', 'font_color': '#1A5E20', 'num_format': '#,##0'})

                # --- Tulis header ---
                for col_idx, col_name in enumerate(df.columns):
                    worksheet.write(0, col_idx, col_name)

                numeric_cols = df.select_dtypes(include=["number"]).columns.tolist()
                vendor_cols  = [c for c in numeric_cols] if sheet == "Bid & Price Analysis" else []

                # Atur lebar kolom
                for col_idx, col_name in enumerate(df.columns):
                    if col_name in numeric_cols:
                        worksheet.set_column(col_idx, col_idx, 15, fmt_rupiah)
                    if "%" in col_name:
                        worksheet.set_column(col_idx, col_idx, 15, fmt_pct)

                # --- Tulis baris dengan highlight ---
                for row_idx, row in enumerate(df.itertuples(index=False), start=1):
                    fmt_row = [None]*len(df.columns)

                    # --- Merge Data highlight ---
                    if sheet == "Merge Data":
                        year_val  = str(row[1]).upper()   # kolom Year
                        scope_val = str(row[2]).upper()   # kolom Scope
                        if scope_val == "TOTAL" and year_val != "TOTAL":
                            fmt_row = [fmt_total_year]*len(df.columns)
                        elif year_val == "TOTAL":
                            fmt_row = [fmt_total_vendor]*len(df.columns)

                    # --- Cost Summary highlight ---
                    elif sheet == "Cost Summary":
                        year_val  = str(row[1]).upper()
                        scope_val = str(row[3]).upper()
                        if scope_val == "TOTAL" and year_val != "TOTAL":
                            fmt_row = [fmt_total_year]*len(df.columns)
                        elif year_val == "TOTAL":
                            fmt_row = [fmt_total_vendor]*len(df.columns)

                    # --- Bid & Price Analysis highlight ---
                    elif sheet == "Bid & Price Analysis":
                        first_vendor_name  = row[df.columns.get_loc("1st Vendor")]
                        second_vendor_name = row[df.columns.get_loc("2nd Vendor")]
                        for col_idx2, col_name2 in enumerate(df.columns):
                            if col_name2 == first_vendor_name:
                                fmt_row[col_idx2] = fmt_first
                            elif col_name2 == second_vendor_name:
                                fmt_row[col_idx2] = fmt_second
                            elif str(row[col_idx2]).upper() == "TOTAL":
                                fmt_row[col_idx2] = fmt_total

                    # --- TCO sheets highlight TOTAL ---
                    else:
                        if any(str(row[i]).upper() == "TOTAL" for i in range(len(df.columns)) if row[i] is not None):
                            fmt_row = [fmt_total]*len(df.columns)

                    # --- Tulis sel ---
                    for col_idx2 in range(len(df.columns)):
                        value = row[col_idx2]
                        if pd.isna(value) or (isinstance(value, (int,float)) and np.isinf(value)):
                            value = ""
                        worksheet.write(row_idx, col_idx2, value, fmt_row[col_idx2] if fmt_row[col_idx2] else None)

        output.seek(0)
        return output.getvalue()

    # --- FRAGMENT UNTUK BALLOONS ---
    @st.fragment
    def release_the_balloons():
        st.balloons()

    # ---- DOWNLOAD BUTTON ----
    if selected_sheets:
        excel_bytes = generate_multi_sheet_excel(selected_sheets, dataframes)

        st.download_button(
            label="Download",
            data=excel_bytes,
            file_name="TCO Comparison by Year + Region.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            on_click=release_the_balloons,
            type="primary",
            use_container_width=True,
        )
