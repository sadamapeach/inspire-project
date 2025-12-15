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

# Download Highlight 1st & 2nd Vendors
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
            1️⃣ TCO Comparison by Year
        </div>
        """,
        unsafe_allow_html=True
    )
    # st.header("1️⃣ TCO Comparison by Year")
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
        st.session_state["uploaded_file_tco_by_year"] = upload_file

        # --- Animasi proses upload ---
        msg = st.toast("📂 Uploading file...")
        time.sleep(1.2)

        msg.toast("🔍 Reading sheets...")
        time.sleep(1.2)

        msg.toast("✅ File uploaded successfully!")
        time.sleep(0.5)

        # Baca semua sheet sekaligus
        all_df = pd.read_excel(upload_file, sheet_name=None)
        st.session_state["all_df_tco_by_year_raw"] = all_df  # simpan versi mentah

    elif "all_df_tco_by_year_raw" in st.session_state:
        all_df = st.session_state["all_df_tco_by_year_raw"]
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

        # PENANGANAN KOLOM TOTALLL
        # Ambil semua kolom numerik
        num_cols = df_clean.select_dtypes(include=["number"]).columns.tolist()

        # Safety: kalau numeric col < 2 → tidak mungkin ada total
        if len(num_cols) < 2:
            df_clean["TOTAL"] = df_clean[num_cols].sum(axis=1)
        else:
            # Hitung apakah kolom numerik terakhir = penjumlahan sebelumnya
            user_has_total = (
                (df_clean[num_cols[:-1]].sum(axis=1) - df_clean[num_cols[-1]])
                .abs() < 1e-6
            ).all()  # semua baris harus match

            if user_has_total:
                # Anggap kolom terakhir sebagai TOTAL
                df_clean = df_clean.rename(columns={num_cols[-1]: "TOTAL"})
            else:
                # User belum hitung TOTAL → kita buat
                df_clean["TOTAL"] = df_clean[num_cols].sum(axis=1)

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

    st.session_state["result_tco_by_year"] = result
    # st.divider()

    # MERGE OVERVIEW
    st.markdown("##### 🗃️ Merge Data")
    st.caption(f"Successfully consolidated data from **{total_sheets} vendors**.")

    merged_list = []
    for vendor, df_clean in result.items():
        df_temp = df_clean.copy()

        # Identifikasi kolom numerik & non-numerik
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
        df_temp.insert(0, "VENDOR", vendor)

        merged_list.append(df_temp)

    # Gabungkan semua vendor jadi satu DataFrame
    df_merged = pd.concat(merged_list, ignore_index=True)

    # Pastikan kolom berurutan (vendor as index-0)
    cols = ["VENDOR"] + [c for c in df_merged.columns if c != "VENDOR"]
    df_merged = df_merged[cols]

    # Simpan session
    st.session_state["merge_overview_tco_by_year"] = df_merged

    # Format rupiah dan tampilkan
    num_cols = df_merged.select_dtypes(include=["number"]).columns
    df_merged_styled = (
        df_merged.style
        .format({col: format_rupiah for col in num_cols})
        .apply(highlight_total_row_v2, axis=1)
    )
    st.dataframe(df_merged_styled, hide_index=True)

    # Download
    excel_data = get_excel_download_highlight_total(df_merged)
    col1, col2, col3 = st.columns([2.3,2,1])
    with col3:
        st.download_button(
            label="Download",
            data=excel_data,
            file_name="Merge Data - TCO by Year.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            icon=":material/download:",
        )

    st.divider()

    # TCO SUMMARY
    st.markdown("##### 💸 TCO Summary")

    # Merge semua sheet berdasarkan TCO Component (indeks kolom pertama)
    merged = None
    ref_order = None  # simpan urutan referensi dari sheet pertama

    for i, (name, df_sub) in enumerate(result.items()):
        num_cols = df_sub.select_dtypes(include=["number"]).columns.tolist()
        non_num_cols = [c for c in df_sub.columns if c not in num_cols]
        total_col = "TOTAL"     # Total cost 5Y

        df_merge_tco = df_sub[non_num_cols + [total_col]].rename(columns={total_col: name})

        # Simpan urutan referensi dari sheet pertama
        if i == 0:
            ref_order = df_merge_tco[non_num_cols].astype(str)
            first_non_num_cols = non_num_cols.copy()
            merged = df_merge_tco
        else:
            merged = merged.merge(df_merge_tco, on=non_num_cols, how="outer")

    # Reorder baris sesuai urutan dari sheet pertama
    if ref_order is not None:
        for col in first_non_num_cols:
            merged[col] = merged[col].astype(str)
            categories = list(dict.fromkeys(ref_order[col].astype(str).tolist()))
            merged[col] = pd.Categorical(
                merged[col], 
                categories=categories,
                ordered=True
            )
        merged = merged.sort_values(first_non_num_cols).reset_index(drop=True)

    # Menambahkan baris total di akhir
    total_row = {col: "" for col in merged.columns}  # kosongkan dulu

    # isi label TOTAL di kolom non-numerik pertama
    if len(first_non_num_cols) > 0:
        total_row[first_non_num_cols[0]] = "TOTAL"

    # hitung total untuk numeric columns
    for col in merged.columns:
        if pd.api.types.is_numeric_dtype(merged[col]):
            total_row[col] = merged[col].sum()

    df_tco_summary = pd.concat([merged, pd.DataFrame([total_row])], ignore_index=True)

    # Fomat Rupiah & fungsi untuk styling baris TOTAL
    num_cols = df_tco_summary.select_dtypes(include=["number"]).columns

    df_tco_summary_styled = (
        df_tco_summary.style
        .format({col: format_rupiah for col in num_cols})
        .apply(highlight_total_row, axis=1)
        .apply(lambda row: highlight_rank_summary(row, num_cols), axis=1)
    )

    st.markdown(
        """
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px;">
            <div style="font-size:0.88rem; color:gray; font-weight:400;">
                The following table presents the summary of the analysis results.
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

    st.session_state["merged_tco_by_year"] = df_tco_summary
    st.dataframe(df_tco_summary_styled, hide_index=True)

    # Download button to Excel
    excel_data = get_excel_download_highlight_summary(df_tco_summary)

    # NEW FEATURE: CONVERGE
    # --- Fungsi reset ---
    def reset_fields():
        for key in ["amount_input", "currency_input", "tco_by_year_amount", "tco_by_year_currency", "converted_tco_by_year", "widget_key"]:
            if key in st.session_state:
                del st.session_state[key]
        # Tambah key unik biar widget benar-benar re-render kosong
        st.session_state["widget_key"] = str(time.time())
        # st.rerun()

        # Set flag untuk rerun
        st.session_state["should_rerun"] = True
    
    # Jalankan rerun di luar callback (AMAN)
    if st.session_state.get("should_rerun", False):
        st.session_state["should_rerun"] = False
        st.rerun()

    # # --- Dapatkan key unik untuk widget ---
    widget_key = st.session_state.get("widget_key", "default")

    # Ambil nilai default (kalau sebelumnya sudah ada)
    default_amount = st.session_state.get("tco_by_year_amount", "")
    default_currency = st.session_state.get("tco_by_year_currency", "")
    
    # --- Popover ---
    col1, col2, col3 = st.columns([2.3,2,1])
    with col1:
        with st.popover("Currency Converter"):
            col1, col2 = st.columns([2, 1])

            with col1:
                amount_input = st.text_input(
                    "Enter amount to convert",
                    placeholder="e.g. 15000, 0.67",
                    key=f"amount_input_{widget_key}",  # 🔑 pakai key unik
                    value=default_amount
                )

            with col2:
                currency_options = ["", "USD", "EUR", "GBP", "SGD", "JPY", "CNY", "INR", "AUD", "CHF", "IDR"]
                index = currency_options.index(default_currency) if default_currency in currency_options else 0
                currency_input = st.selectbox(
                    "Currency",
                    currency_options,
                    key=f"currency_input_{widget_key}",  # 🔑 pakai key unik
                    index=index
                )

            st.button("Reset", on_click=reset_fields, type="primary")

    with col3:
        st.download_button(
            label="Download",
            data=excel_data,
            file_name="TCO Summary - TCO by Year.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            icon=":material/download:",
        )

    # --- Simpan nilai setelah widget dirender ---
    if amount_input:
        st.session_state["tco_by_year_amount"] = amount_input
    if currency_input:
        st.session_state["tco_by_year_currency"] = currency_input

    # --- Ambil kembali nilai ---
    amount = st.session_state.get("tco_by_year_amount", "")
    currency = st.session_state.get("tco_by_year_currency", "")

    # Tampilkan tabel hasil konversi hanya jika input terisi
    if amount and currency:
        try:
            # Preprocessing dulu
            cleaned_amount = re.sub(r"[^\d,\.]", "", amount)
            if "," in cleaned_amount and "." not in cleaned_amount:
                cleaned_amount = cleaned_amount.replace(",", ".")

            # Hapus tanda pemisah ribuan (baik koma maupun titik)
            cleaned_amount = re.sub(r"(?<=\d)[.,](?=\d{3}(\D|$))", "", cleaned_amount)

            # Konversi nominal jadi float
            amount_value = float(cleaned_amount)

            # Salin merged untuk dikalikan
            df_converted = merged.copy(deep=True)

            # Identifikasi kolom numerik
            def is_convertible_numeric(series: pd.Series) -> bool:
                coerced = pd.to_numeric(series, errors="coerce")
                return coerced.notna().any()

            # jangan sentuh kolom pertama (biasanya TCO Component)
            cols_except_first = list(df_converted.columns[1:])

            # pilih kolom yang "bisa" menjadi numeric dari kolom-kolom tersebut
            numeric_cols_to_multiply = [
                c for c in cols_except_first if is_convertible_numeric(df_converted[c])
            ]

            # Konversi & kalikan hanya pada kolom terdeteksi
            if numeric_cols_to_multiply:
                df_converted.loc[:, numeric_cols_to_multiply] = (
                    df_converted.loc[:, numeric_cols_to_multiply]
                    .apply(pd.to_numeric, errors="coerce")
                    .multiply(amount_value)
                )

            # Simpan hasil ke session_state (biar tidak hilang)
            st.session_state["converted_tco_by_year"] = df_converted

        except ValueError:
            st.error("❌ Invalid number format. Please check your input.")

        # Download button for converter
        if "converted_tco_by_year" in st.session_state:
            df_converted = st.session_state["converted_tco_by_year"]
            currency = st.session_state.get("tco_by_year_currency", "")

            st.markdown("###### Converted Price")

            # Identifikasi numeric & non-numeric columns
            num_cols = df_converted.select_dtypes(include=["number"]).columns.tolist()
            non_num_cols = [c for c in df_converted.columns if c not in num_cols]

            # Buat baris total dinamis
            total_row = {col: "" for col in df_converted.columns}

            # Isi label "TOTAL" pada kolom non-numeric pertama
            if len(non_num_cols) > 0:
                total_row[non_num_cols[0]] = "TOTAL"

            # Hitung sum hanya untuk kolom numeric
            for col in num_cols:
                total_row[col] = df_converted[col].sum()

            # Gabungkan
            df_tco_converted = pd.concat([df_converted, pd.DataFrame([total_row])], ignore_index=True)

            # Fomat Rupiah & fungsi untuk styling baris TOTAL
            num_cols_after = df_tco_converted.select_dtypes(include=["number"]).columns

            converted_styled = (
                df_tco_converted.style
                .format({col: format_rupiah for col in num_cols_after})
                .apply(highlight_total_row, axis=1)
                .apply(lambda row: highlight_rank_summary(row, num_cols_after), axis=1)
            )

            st.markdown(
                f"""
                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px;">
                    <div style="font-size:0.88rem; color:gray; font-weight:400;">
                        Summary of total bidders after converting to <b>{currency}</b> at a rate of <b>{amount}</b>.
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

            st.dataframe(converted_styled, hide_index=True)

            # Download button to Excel
            excel_data_converted = get_excel_download_highlight_summary(df_tco_converted)

            # Layout tombol (rata kanan)
            col1, col2, col3 = st.columns([2.3,2,1])
            with col3:
                st.download_button(
                    label="Download",
                    data=excel_data_converted,
                    file_name="TCO Converted - TCO by Year.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    icon=":material/download:",
                )
    
    st.divider()

    # BID & PRICE ANALYSIS
    st.markdown("##### 🧠 Bid & Price Analysis")
    # st.caption("Comparative analysis across vendors including lowest price, gap percentage, and deviation from median.")

    df_analysis = df_tco_summary.copy()

    # Identifikasi kolom
    non_numeric_cols = df_analysis.select_dtypes(exclude=["number"]).columns.tolist()
    vendor_cols = df_analysis.select_dtypes(include=["number"]).columns.tolist()

    # Hapus baris TOTAL untuk analisis per komponen
    df_no_total = df_analysis[
        ~df_analysis.apply(
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

    # Urutkan kolom sesuai struktur yang diinginkan
    analysis_cols = (
        non_numeric_cols
        + vendor_cols
        + ["1st Lowest", "1st Vendor", "2nd Lowest", "2nd Vendor", "Gap 1 to 2 (%)", "Median Price"]
        + [f"{v} to Median (%)" for v in vendor_cols]
    )

    df_analysis_final = df_no_total[analysis_cols]

    # SLICERR FOR ANALYSIS
    all_1st = sorted(df_analysis_final["1st Vendor"].dropna().unique())
    all_2nd = sorted(df_analysis_final["2nd Vendor"].dropna().unique())

    col_sel_1, col_sel_2 = st.columns(2)
    with col_sel_1:
        selected_1st = st.multiselect(
            "Filter: 1st vendor",
            options=all_1st,
            default=None,
            placeholder="Choose one or more vendors",
            key=f"filter_1st_vendor"
        )
    with col_sel_2:
        selected_2nd = st.multiselect(
            "Filter: 2nd vendor",
            options=all_2nd,
            default=None,
            placeholder="Choose one or more vendors",
            key=f"filter_2nd_vendor"
        )

    # --- Terapkan filter dengan logika AND
    if selected_1st and selected_2nd:
        df_filtered_analysis = df_analysis_final[
            df_analysis_final["1st Vendor"].isin(selected_1st) &
            df_analysis_final["2nd Vendor"].isin(selected_2nd)
        ]
    elif selected_1st:
        df_filtered_analysis = df_analysis_final[df_analysis_final["1st Vendor"].isin(selected_1st)]
    elif selected_2nd:
        df_filtered_analysis = df_analysis_final[df_analysis_final["2nd Vendor"].isin(selected_2nd)]
    else:
        df_filtered_analysis = df_analysis_final.copy()

    # Format rupiah
    num_cols = df_filtered_analysis.select_dtypes(include=["number"]).columns
    format_dic = {col: format_rupiah for col in num_cols}
    format_dic.update({"Gap 1 to 2 (%)": "{:.1f}%"})
    for v in vendor_cols:
        format_dic[f"{v} to Median (%)"] = "{:+.1f}%"

    df_analysis_styled = (
        df_filtered_analysis.style
        .format(format_dic)
        .apply(lambda row: highlight_1st_2nd_vendor(row, df_analysis_final.columns), axis=1)
    )

    st.markdown(
        f"""
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px;">
            <div style="font-size:0.88rem; color:gray; font-weight:400;">
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
        
    st.dataframe(df_analysis_styled, hide_index=True)

    # Download button to Excel
    excel_data = get_excel_download_highlight_1st_2nd_lowest(df_filtered_analysis)

    # Layout tombol (rata kanan)
    col1, col2, col3 = st.columns([2.3,2,1])
    with col3:
        st.download_button(
            label="Download",
            data=excel_data,
            file_name="Bid & Price Analysis - TCO by Year.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            icon=":material/download:",
        )

    # RANK
    if merged is not None:
        st.markdown("##### 🏅 Rank Visualization")
        # st.caption("Rangking is generated based on each vendor's overall total cost.")
        # Tab
        tab1, tab2 = st.tabs(["💲Original Price", "💱 Converted Price"])

        # =============== TAB 1 =============== #
        # Logic sum
        sum_series = merged.iloc[:, 1:].sum(numeric_only=True)

        # Ranking vendor
        df_chart = (
            sum_series.reset_index()
            .rename(columns={"index": "Vendor", 0: "Total"})
            .sort_values("Total", ascending=True)
        )
        df_chart["Rank"] = range(1, len(df_chart) + 1)
        df_chart["Total_str"] = df_chart["Total"].apply(lambda x: f"{int(x):,}".replace(",", "."))
        df_chart["Legend"] = df_chart.apply(lambda x: f"Rank {x['Rank']} - {x['Total_str']}", axis=1)

        # Warna vendor
        vendor_colors = {
            v: c for v, c in zip(df_chart["Legend"], [
                "#32AAB5", "#E7D39A", "#F1AA60", "#F27B68",
                "#E54787", "#BF219A", "#8E0F9C", "#4B1D91"
            ])
        }

        # Format angka besar di sumbu Y → jadi singkat (1K, 1M)
        y_axis = alt.Axis(title=None, grid=False, format=".0s")

        highlight = alt.selection_point(on='mouseover', fields=['Vendor'], nearest=True)

        # Chart bar utama
        bars = (
            alt.Chart(df_chart)
            .mark_bar()
            .encode(
                x=alt.X("Vendor:N", axis=alt.Axis(title=None)),
                y=alt.Y("Total:Q", axis=y_axis, scale=alt.Scale(domain=[0, df_chart["Total"].max() * 1.1])),
                color=alt.Color(
                    "Legend:N",
                    title="Total Offer by Rank",
                    scale=alt.Scale(domain=list(vendor_colors.keys()), range=list(vendor_colors.values()))
                ),
                tooltip=[
                    alt.Tooltip("Vendor:N", title="Vendor"),
                    alt.Tooltip("Total_str:N", title="Total (USD)")
                ]
            ).add_params(highlight)
        )

        # Temukan nilai total terendah
        lowest_value = df_chart["Total"].min()

        # Garis horizontal dashed di posisi nilai terendah
        lowest_line = (
            alt.Chart(df_chart)
            .mark_rule(color="red", strokeDash=[5, 3], strokeWidth=1)
            .encode(y=alt.datum(lowest_value))
        )

        # Label Rank di atas bar
        rank_text = (
            alt.Chart(df_chart)
            .mark_text(dy=-7, color="gray", fontWeight="bold")
            .encode(x="Vendor:N", y="Total:Q", text="Rank:N")
        )

        # Frame border
        frame = (
            alt.Chart().mark_rect(stroke='gray', strokeWidth=1, fillOpacity=0)
        )

        tab1.write("")
        # Gabung semua layer + style
        chart = (
            (bars + lowest_line + rank_text + frame)
            .properties(title="✨ Original Price: Bidder Rangking by Total Offer ✨")
            .configure_title(anchor="middle", offset=12, fontSize=13)
            .configure_legend(
                titleFontSize=11,
                titleFontWeight="bold",
                labelFontSize=11,
                labelLimit=300,
                orient="right"
            )
        )

        tab1.altair_chart(chart)

        # =============== TAB 2 =============== #
        if amount.strip() != "" and currency.strip() != "":
            # Logic sum
            sum_series = df_converted.iloc[:, 1:].sum(numeric_only=True)

            # Ranking vendor
            df_chart = (
                sum_series.reset_index()
                .rename(columns={"index": "Vendor", 0: "Total"})
                .sort_values("Total", ascending=True)
            )
            df_chart["Rank"] = range(1, len(df_chart) + 1)
            df_chart["Total_str"] = df_chart["Total"].apply(lambda x: f"{int(x):,}".replace(",", "."))
            df_chart["Legend"] = df_chart.apply(lambda x: f"Rank {x['Rank']} - {x['Total_str']} {currency}", axis=1)

            # Warna vendor
            vendor_colors = {
                v: c for v, c in zip(df_chart["Legend"], [
                    "#32AAB5", "#E7D39A", "#F1AA60", "#F27B68",
                    "#E54787", "#BF219A", "#8E0F9C", "#4B1D91"
                ])
            }

            # Format angka besar di sumbu Y → jadi singkat (1K, 1M)
            y_axis = alt.Axis(title=None, grid=False, format=".0s")

            highlight = alt.selection_point(on='mouseover', fields=['Vendor'], nearest=True)

            # Chart bar utama
            bars = (
                alt.Chart(df_chart)
                .mark_bar()
                .encode(
                    x=alt.X("Vendor:N", axis=alt.Axis(title=None)),
                    y=alt.Y("Total:Q", axis=y_axis, scale=alt.Scale(domain=[0, df_chart["Total"].max() * 1.1])),
                    color=alt.Color(
                        "Legend:N",
                        title="Total Offer by Rank",
                        scale=alt.Scale(domain=list(vendor_colors.keys()), range=list(vendor_colors.values()))
                    ),
                    tooltip=[
                        alt.Tooltip("Vendor:N", title="Vendor"),
                        alt.Tooltip("Total_str:N", title="Total (USD)")
                    ]
                ).add_params(highlight)
            )

            # Temukan nilai total terendah
            lowest_value = df_chart["Total"].min()

            # Garis horizontal dashed di posisi nilai terendah
            lowest_line = (
                alt.Chart(df_chart)
                .mark_rule(color="red", strokeDash=[5, 3], strokeWidth=1)
                .encode(y=alt.datum(lowest_value))
            )

            # Label Rank di atas bar
            rank_text = (
                alt.Chart(df_chart)
                .mark_text(dy=-7, color="gray", fontWeight="bold")
                .encode(x="Vendor:N", y="Total:Q", text="Rank:N")
            )

            # Frame border
            frame = (
                alt.Chart().mark_rect(stroke='gray', strokeWidth=1, fillOpacity=0)
            )

            tab2.write("")
            # Gabung semua layer + style
            chart = (
                (bars + lowest_line + rank_text + frame)
                .properties(title="✨ Converted Price: Bidder Rangking by Total Offer ✨")
                .configure_title(anchor="middle", offset=12, fontSize=13)
                .configure_legend(
                    titleFontSize=11,
                    titleFontWeight="bold",
                    labelFontSize=11,
                    labelLimit=300,
                    orient="right"
                )
            )

            tab2.altair_chart(chart)

        else:
            tab2.markdown(
                """
                <div style='background-color:#ffe6f2; padding:8px 12px; border-radius:8px; margin-bottom:15px;'>
                    <p style='font-size:13px; color:#a8326d; margin:4px;'>
                        💡 No converted data found. Please use the <b>Currency Converter</b> first.
                    </p>
                </div>
                """,
                unsafe_allow_html=True
            )

        # COMPARISON
        st.markdown("##### 📊 Component Comparison")
        # st.caption("Each chart shows total offer comparison and ranking for every requirement.")
        if merged is not None and not merged.empty:
            tab3, tab4 = st.tabs(["💲Original Price", "💱 Converted Price"])

            # Identifikasi kolom
            non_numeric_cols = merged.select_dtypes(exclude=["number"]).columns.tolist()
            vendor_cols = merged.select_dtypes(include=["number"]).columns.tolist()

            # Requirements (pakai kolom non-numeric pertama, misal TCO Component)
            reqs = merged[non_numeric_cols[0]]
            vendors = vendor_cols

            tab3.caption("")

            n_cols = 2
            for i in range(0, len(reqs), n_cols):
                cols = tab3.columns(n_cols)

                for j in range(n_cols):
                    if i + j < len(reqs):
                        req = reqs[i + j]

                        # --- Data vendor + harga ---
                        prices = merged.loc[i + j, vendor_cols].reset_index()
                        prices.columns = ["Vendor", "Total"]

                        # pastikan numeric
                        prices["Total"] = pd.to_numeric(prices["Total"], errors="coerce")

                        # Urutkan harga + tambahkan rank
                        df_chart = prices.sort_values("Total", ascending=True).reset_index(drop=True)
                        df_chart["Rank"] = range(1, len(df_chart) + 1)

                        # Format angka ribuan
                        df_chart["Total_str"] = df_chart["Total"].apply(lambda x: f"{int(x):,}".replace(",", "."))
                        df_chart["Legend"] = df_chart.apply(lambda x: f"R{x['Rank']} - {x['Total_str']}", axis=1)

                        # --- Warna manual konsisten antar chart ---
                        vendor_colors = {
                            v: c for v, c in zip(df_chart["Legend"], [
                                "#32AAB5", "#E7D39A", "#F1AA60", "#F27B68",
                                "#E54787", "#BF219A", "#8E0F9C", "#4B1D91"
                            ])
                        }

                        # Format angka besar di sumbu Y → jadi singkat (1K, 1M)
                        y_axis = alt.Axis(title=None, grid=False, format=".0s", labelPadding=12)

                        highlight = alt.selection_point(on='mouseover', fields=['Vendor'], nearest=True)

                        # --- Dashed line pada nilai minimum ---
                        min_val = df_chart["Total"].min()
                        dashed_line = (
                            alt.Chart(pd.DataFrame({"y": [min_val]}))
                            .mark_rule(strokeDash=[6, 3], color="red", strokeWidth=1)
                            .encode(y="y:Q")
                        )

                        # --- Bar chart utama ---
                        bars = (
                            alt.Chart(df_chart)
                            .mark_bar()
                            .encode(
                                x=alt.X("Vendor:N", axis=alt.Axis(title=None)),
                                y=alt.Y("Total:Q", axis=y_axis,
                                        scale=alt.Scale(domain=[0, df_chart["Total"].max() * 1.1])),
                                color=alt.Color(
                                    "Legend:N",
                                    title=f"Total {req} Offer by Rank",
                                    scale=alt.Scale(domain=list(vendor_colors.keys()),
                                                    range=list(vendor_colors.values()))
                                ),
                                tooltip=[
                                    alt.Tooltip("Vendor:N", title="Vendor"),
                                    alt.Tooltip("Total_str:N", title="Total (USD)"),
                                    alt.Tooltip("Rank:N", title="Rank")
                                ]
                            ).add_params(highlight)
                        )

                        # --- Rank di tengah bar ---
                        rank_text = (
                            alt.Chart(df_chart)
                            .mark_text(dy=-7, color="gray", fontWeight="bold")
                            .encode(x="Vendor:N", y="Total:Q", text="Rank:N")
                        )

                        # --- Frame keliling ---
                        frame = (
                            alt.Chart(df_chart)
                            .mark_rect(stroke="gray", strokeWidth=1, fillOpacity=0)
                        )

                        # --- Gabungkan chart ---
                        chart = (
                            (bars + dashed_line + rank_text + frame)
                            .properties(
                                title=f"💲Comparison: {req}",
                            )
                            .configure_title(
                                anchor="middle",
                                fontSize=13,
                                offset=12,
                                fontWeight="bold"
                            )
                            .configure_legend(
                                orient="bottom",
                                direction="horizontal",
                                columns=2,
                                titleFontSize=11,
                                titleFontWeight="bold",
                                labelFontSize=10,
                                labelLimit=300,
                                symbolSize=70
                            )
                        )

                        # --- Tampilkan di kolom ---
                        cols[j].altair_chart(chart)

            # Tab 4
            # Cek apakah data hasil konversi sudah ada
            if amount.strip() != "" and currency.strip() != "":
                reqs = df_converted[df_converted.columns[0]]
                vendor_cols = df_converted.select_dtypes(include=["number"]).columns.tolist()
                vendors = vendor_cols

                tab4.caption("")

                n_cols = 2
                for i in range(0, len(reqs), n_cols):
                    cols = tab4.columns(n_cols)

                    for j in range(n_cols):
                        if i + j < len(reqs):
                            req = reqs[i + j]

                            # --- Data vendor + harga ---
                            prices = df_converted.loc[i + j, vendors].reset_index()
                            prices.columns = ["Vendor", "Total"]

                            # Urutkan harga + tambahkan rank
                            df_chart = prices.sort_values("Total", ascending=True).reset_index(drop=True)
                            df_chart["Rank"] = range(1, len(df_chart) + 1)

                            # Format angka ribuan
                            df_chart["Total_str"] = df_chart["Total"].apply(lambda x: f"{int(x):,}".replace(",", "."))
                            df_chart["Legend"] = df_chart.apply(lambda x: f"R{x['Rank']} - {x['Total_str']} {currency}", axis=1)

                            # --- Warna manual konsisten antar chart ---
                            vendor_colors = {
                                v: c for v, c in zip(df_chart["Legend"], [
                                    "#32AAB5", "#E7D39A", "#F1AA60", "#F27B68",
                                    "#E54787", "#BF219A", "#8E0F9C", "#4B1D91"
                                ])
                            }

                            # Format angka besar di sumbu Y → jadi singkat (1K, 1M)
                            y_axis = alt.Axis(title=None, grid=False, format=".0s", labelPadding=12)

                            highlight = alt.selection_point(on='mouseover', fields=['Vendor'], nearest=True)

                            # --- Dashed line pada nilai minimum ---
                            min_val = df_chart["Total"].min()
                            dashed_line = (
                                alt.Chart(pd.DataFrame({"y": [min_val]}))
                                .mark_rule(strokeDash=[6, 3], color="red", strokeWidth=1)
                                .encode(y="y:Q")
                            )

                            # --- Bar chart utama ---
                            bars = (
                                alt.Chart(df_chart)
                                .mark_bar()
                                .encode(
                                    x=alt.X("Vendor:N", axis=alt.Axis(title=None)),
                                    y=alt.Y("Total:Q", axis=y_axis,
                                            scale=alt.Scale(domain=[0, df_chart["Total"].max() * 1.1])),
                                    color=alt.Color(
                                        "Legend:N",
                                        title=f"Total {req} Offer by Rank",
                                        scale=alt.Scale(domain=list(vendor_colors.keys()),
                                                        range=list(vendor_colors.values()))
                                    ),
                                    tooltip=[
                                        alt.Tooltip("Vendor:N", title="Vendor"),
                                        alt.Tooltip("Total_str:N", title="Total (USD)"),
                                        alt.Tooltip("Rank:N", title="Rank")
                                    ]
                                ).add_params(highlight)
                            )

                            # --- Rank di tengah bar ---
                            rank_text = (
                                alt.Chart(df_chart)
                                .mark_text(dy=-7, color="gray", fontWeight="bold")
                                .encode(x="Vendor:N", y="Total:Q", text="Rank:N")
                            )

                            # --- Frame keliling ---
                            frame = (
                                alt.Chart(df_chart)
                                .mark_rect(stroke="gray", strokeWidth=1, fillOpacity=0)
                            )

                            # --- Gabungkan chart ---
                            chart = (
                                (bars + dashed_line + rank_text + frame)
                                .properties(
                                    title=f"💱 Comparison: {req}",
                                )
                                .configure_title(
                                    anchor="middle",
                                    fontSize=13,
                                    offset=12,
                                    fontWeight="bold"
                                )
                                .configure_legend(
                                    orient="bottom",
                                    direction="horizontal",
                                    columns=2,
                                    titleFontSize=11,
                                    titleFontWeight="bold",
                                    labelFontSize=10,
                                    labelLimit=300,
                                    symbolSize=70
                                )
                            )

                            # --- Tampilkan di kolom ---
                            cols[j].altair_chart(chart)
            else:
                # Jika belum ada data konversi
                tab4.markdown(
                    """
                    <div style='background-color:#ffe6f2; padding:8px 12px; border-radius:8px; margin-bottom:15px;'>
                        <p style='font-size:13px; color:#a8326d; margin:4px;'>
                            💡 No converted data found. Please use the <b>Currency Converter</b> first.
                        </p>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

    st.divider()

    # SUPERRR BUTTONN
    st.markdown("##### 🧑‍💻 Super Download — Export Selected Sheets")
    dataframes = {
        "Merge Data": df_merged,
        "TCO Summary": df_tco_summary,
        "Bid & Price Analysis": df_filtered_analysis,
    }

    if "converted_tco_by_year" in st.session_state:
        dataframes["TCO Converted"] = df_tco_converted

    # Tampilkan multiselect
    selected_sheets = st.multiselect(
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
                    if sheet in ["TCO Summary", "TCO Converted"]:
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
        st.balloons()

    # ---- DOWNLOAD BUTTON ----
    if selected_sheets:
        excel_bytes = generate_multi_sheet_excel(selected_sheets, dataframes)

        st.download_button(
            label="Download",
            data=excel_bytes,
            file_name="TCO Comparison by Year.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            on_click=release_the_balloons,
            type="primary",
            use_container_width=True,
        )