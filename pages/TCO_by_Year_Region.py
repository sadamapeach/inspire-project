import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
import time
import re
from io import BytesIO
from functools import reduce

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
def get_excel_download_with_highlight(df, sheet_name="Sheet1"):
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        
        workbook  = writer.book
        worksheet = writer.sheets[sheet_name]

        # --- Format untuk highlight ---
        total_year_format = workbook.add_format({
            'bold': True, 
            'bg_color': '#FFEB9C',  # kuning lembut
            'font_color': '#1A1A1A'  # kuning agak gelap
        })
        vendor_total_format = workbook.add_format({
            'bold': True, 
            'bg_color': '#C6EFCE',  # hijau lembut
            'font_color': '#1A5E20'  # hijau agak gelap
        })

        # --- Iterasi baris untuk highlight ---
        for row_num, row in df.iterrows():
            scope_val = str(row[df.columns[2]]).strip().upper()  # kolom Scope
            year_val = str(row[df.columns[1]]).strip().upper()   # kolom Year
            
            if scope_val == "TOTAL" and year_val != "TOTAL":
                fmt = total_year_format
            elif year_val == "TOTAL":
                fmt = vendor_total_format
            else:
                continue

            # Warnai hanya kolom yang berisi data
            for col_num, value in enumerate(row):
                worksheet.write(row_num + 1, col_num, value, fmt)

    return output.getvalue()

# Download highlight total ver2
@st.cache_data
def get_excel_download_with_highlight_v2(df, sheet_name="Sheet1"):
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        
        workbook  = writer.book
        worksheet = writer.sheets[sheet_name]

        # --- Format untuk highlight ---
        total_year_format = workbook.add_format({
            'bold': True, 
            'bg_color': '#FFEB9C',  # kuning lembut
            'font_color': '#1A1A1A'  # teks agak gelap
        })
        vendor_total_format = workbook.add_format({
            'bold': True, 
            'bg_color': '#C6EFCE',  # hijau lembut
            'font_color': '#1A5E20'  # teks hijau tua
        })

        # --- Deteksi nama kolom dinamis ---
        year_col = next((c for c in df.columns if "YEAR" in c.upper()), None)
        scope_col = next((c for c in df.columns if "SCOPE" in c.upper()), None)

        # --- Iterasi baris untuk highlight ---
        for row_num, row in df.iterrows():
            year_val = str(row[year_col]).strip().upper() if year_col else ""
            scope_val = str(row[scope_col]).strip().upper() if scope_col else ""
            
            if scope_val == "TOTAL" and year_val != "TOTAL":
                fmt = total_year_format
            elif year_val == "TOTAL":
                fmt = vendor_total_format
            else:
                continue

            # Warnai hanya kolom berisi data
            for col_num, value in enumerate(row):
                worksheet.write(row_num + 1, col_num, value, fmt)

    return output.getvalue()

# Download Highlight Excel
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
    st.title("3️⃣ TCO by Year + Region")
    st.markdown(
        ":red-badge[Indosat] :orange-badge[Ooredoo] :green-badge[Hutchison]"
    )
    st.caption("Hai Team! Drop in your annual pricing template and let this analytics system work its magic ✨")

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

    #     # --- NOTIFIKASI KHUSUS ---
    #     if (rows_after < rows_before) or (cols_after < cols_before):
    #         st.markdown(
    #             "<p style='font-size:12px; color:#808080; margin-top:-15px; margin-bottom:0;'>"
    #             "Preprocessing completed! Hidden rows and columns removed ✅</p>",
    #             unsafe_allow_html=True
    #         )

    st.session_state["result_tco_by_year_region"] = result
    # st.divider()

    # MERGEE
    st.markdown("##### 🗃️ Merge Data")

    merged_list = []

    for vendor_name, df_vendor in result.items():
        df_temp = df_vendor.copy()
        year_col = df_temp.columns[0]   # Year
        scope_col = df_temp.columns[1]  # Scope
        numeric_cols = df_temp.select_dtypes(include=["number"]).columns

        # Tambahkan kolom VENDOR
        df_temp.insert(0, "VENDOR", vendor_name)

        vendor_result = []

        # --- Loop tiap year ---
        for year, group_year in df_temp.groupby(year_col, dropna=False):
            # Row TOTAL per year
            total_per_year = {
                "VENDOR": vendor_name,
                year_col: year,
                scope_col: "TOTAL",
                **{col: group_year[col].sum(numeric_only=True) for col in numeric_cols}
            }

            # Gabungkan row scope + TOTAL per year
            df_year_with_total = pd.concat(
                [group_year, pd.DataFrame([total_per_year])],
                ignore_index=True
            )
            vendor_result.append(df_year_with_total)

        # --- Gabungkan semua year untuk vendor ini ---
        df_vendor_with_year_total = pd.concat(vendor_result, ignore_index=True)

        # Row TOTAL besar per vendor → hanya jumlahkan row yang scope == 'TOTAL'
        total_rows_only = df_vendor_with_year_total[df_vendor_with_year_total[scope_col] == "TOTAL"]
        vendor_total_row = {
            "VENDOR": vendor_name,
            year_col: "TOTAL",
            scope_col: "",
            **{col: total_rows_only[col].sum(numeric_only=True) for col in numeric_cols}
        }

        df_vendor_final = pd.concat(
            [df_vendor_with_year_total, pd.DataFrame([vendor_total_row])],
            ignore_index=True
        )

        merged_list.append(df_vendor_final)

    # --- Gabungkan semua vendor ---
    df_all_vendors = pd.concat(merged_list, ignore_index=True)

    # --- Urutkan kolom supaya rapi ---
    first_cols = ["VENDOR", year_col, scope_col]
    other_cols = [c for c in df_all_vendors.columns if c not in first_cols]
    df_all_vendors = df_all_vendors[first_cols + other_cols]

    # --- Styling (opsional) ---
    num_cols = df_all_vendors.select_dtypes(include=["number"]).columns
    df_styled = (
        df_all_vendors.style
        .format({col: format_rupiah for col in num_cols})
        .apply(highlight_total_per_year, axis=1)
        .apply(highlight_vendor_total, axis=1)
    )

    st.caption(f"Data from **{total_sheets} vendors** have been successfully consolidated, analyzing a total of **{len(df_all_vendors):,} combined records**.")
    st.dataframe(df_styled, hide_index=True)

    # Download button to Excel
    excel_data = get_excel_download_with_highlight(df_all_vendors, sheet_name="Merge TCO Year + Region")
    # Pastikan berada di tab atau st
    col1, col2, col3 = st.columns([2.3,2,1])
    with col3:
        st.download_button(
            label="Download",
            data=excel_data,
            file_name="Merge_TCO_Year_Region.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            icon=":material/download:",
        )
    
    st.divider()

    # COST SUMMARY
    st.markdown("##### 📑 Cost Summary")

    # --- Tentukan kolom utama ---
    vendor_col = df_all_vendors.columns[0]
    year_col = df_all_vendors.columns[1]
    scope_col = df_all_vendors.columns[2]

    # --- Tentukan kolom region (exclude kolom awal & TOTAL) ---
    region_cols = [
        c for c in df_all_vendors.columns
        if c not in [vendor_col, year_col, scope_col, "TOTAL"]
    ]

    # --- Transform to long format (Vendor, Year, Region, Scope, Price)
    df_total_price = df_all_vendors.melt(
        id_vars=[vendor_col, year_col, scope_col],
        value_vars=region_cols,
        var_name="REGION",
        value_name="PRICE"
    )

    # --- Rapikan urutan kolom ---
    df_total_price = df_total_price[
        [vendor_col, year_col, "REGION", scope_col, "PRICE"]
    ]

    # Simpan ke session_state jika perlu
    st.session_state["merged_long_format_total_price"] = df_total_price

    # Format Rupiah untuk kolom PRICE
    df_total_price_styled = (
        df_total_price.style
        .format({"PRICE": format_rupiah})
        .apply(highlight_total_per_year, axis=1)
        .apply(highlight_vendor_total, axis=1)
    )

    # Tampilkan
    st.caption(f"Consolidated cost summary containing **{len(df_total_price):,} records** across multiple vendors and regions.")
    st.dataframe(df_total_price_styled, hide_index=True)

    # Download button to Excel
    excel_data = get_excel_download_with_highlight_v2(df_total_price)

    # Layout tombol (rata kanan)
    col1, col2, col3 = st.columns([2.3,2,1])
    with col3:
        st.download_button(
            label="Download",
            data=excel_data,
            file_name="Total Price.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            icon=":material/download:",
        )

    # TCOO
    st.markdown("##### 💸 TCO Summary")
    tab1, tab2, tab3 = st.tabs(["YEAR", "REGION", "SCOPE"])

    vendor_col = "VENDOR"
    year_col = df_total_price.columns[1]
    region_col = "REGION"
    scope_col = df_total_price.columns[3]
    price_col = "PRICE"

    # Tab1: YEAR
    # --- Hapus baris TOTAL agar tidak double count ---
    df_year_clean = df_total_price[
        (df_total_price[year_col].astype(str).str.upper() != "TOTAL") &
        (df_total_price[scope_col].astype(str).str.upper() != "TOTAL")
    ]

    df_year = df_year_clean.pivot_table(
        index=year_col,
        columns="VENDOR",
        values="PRICE",
        aggfunc="sum",
        fill_value=0
    ).reset_index()

    total_row = pd.DataFrame({
        year_col: ["TOTAL"], 
        **{col: [df_year[col].sum()] for col in df_year.columns if col != year_col}
    })
    df_year = pd.concat([df_year, total_row], ignore_index=True)

    # Format rupiah, exclude kolom pertama (YEAR)
    format_dict = {col: format_rupiah for col in df_year.columns[1:]}  # skip kolom pertama

    df_year_styled = (
        df_year.style
        .format(format_dict)  # hanya format kolom numeric
        .apply(highlight_total_row, axis=1)
    )

    tab1.caption("Overview of total vendor costs per year, highlighting annual spending trends and competitiveness.")
    tab1.dataframe(df_year_styled, hide_index=True)

    excel_data = get_excel_download(df_year)
    with tab1:
        col1, col2, col3 = st.columns([2.3,2,1])
        with col3:
            st.download_button(
                label="Download",
                data=excel_data,
                file_name="TCO by Year.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                icon=":material/download:",
                key="download_year"
            )

    # Tab2: REGION
    # --- Hapus baris TOTAL agar tidak double count ---
    df_region_clean = df_total_price[
        (df_total_price[year_col].astype(str).str.upper() != "TOTAL") &
        (df_total_price[scope_col].astype(str).str.upper() != "TOTAL")
    ]

    # --- Buat pivot ---
    df_region = df_region_clean.pivot_table(
        index=region_col,
        columns=vendor_col,
        values="PRICE",
        aggfunc="sum",
        fill_value=0
    ).reset_index()

    # --- Tambahkan baris TOTAL ---
    total_row = pd.DataFrame({
        region_col: ["TOTAL"],
        **{col: [df_region[col].sum()] for col in df_region.columns if col != region_col}
    })

    df_region = pd.concat([df_region, total_row], ignore_index=True)

    df_region_styled = (
        df_region.style
        .format(format_rupiah)
        .apply(highlight_total_row, axis=1)
    )

    tab2.caption("Breakdown of vendor costs across regions, showing geographical spending distribution and variations.")
    tab2.dataframe(df_region_styled, hide_index=True)

    excel_data = get_excel_download(df_region)
    with tab2:
        col1, col2, col3 = st.columns([2.3,2,1])
        with col3:
            st.download_button(
                label="Download",
                data=excel_data,
                file_name="TCO by Region.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                icon=":material/download:",
                key="download_region"
            )

    # Tab3: Scope
    df_scope_clean = df_total_price[
        (df_total_price[year_col].astype(str).str.upper() != "TOTAL") &
        (df_total_price[scope_col].astype(str).str.upper() != "TOTAL")
    ]

    df_scope = df_scope_clean.pivot_table(
        index=scope_col,
        columns=vendor_col,
        values=price_col,
        aggfunc="sum",
        fill_value=0
    ).reset_index()

    total_row = pd.DataFrame({
        scope_col: ["TOTAL"],
        **{col: [df_scope[col].sum()] for col in df_scope.columns if col != scope_col}
    })

    df_scope = pd.concat([df_scope, total_row], ignore_index=True)

    df_scope_styled = (
        df_scope.style
        .format(format_rupiah)
        .apply(highlight_total_row, axis=1)
    )

    tab3.caption("Summary of vendor costs by project scope, providing insight into allocation across different work packages.")
    tab3.dataframe(df_scope_styled, hide_index=True)

    excel_data = get_excel_download(df_scope)
    with tab3:
        col1, col2, col3 = st.columns([2.3,2,1])
        with col3:
            st.download_button(
                label="Download",
                data=excel_data,
                file_name="TCO by Scope.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                icon=":material/download:",
                key="download_scope"
            )

    st.divider()

    # --- ANALYTICAL COLUMNS ---
    st.markdown("##### 🧠 Bid & Price Analysis")

    # --- DROP baris TOTAL ---
    df_clean = df_all_vendors[
        ~df_all_vendors["SCOPE"].astype(str).str.upper().eq("TOTAL")
    ].copy()

    # --- Ubah dari format wide (Region1, Region2, dst) ke long ---
    df_melted = df_clean.melt(
        id_vars=["VENDOR", year_col, scope_col],
        var_name="REGION",
        value_name="PRICE"
    )

    # --- Pastikan data numerik bersih ---
    df_melted["PRICE"] = pd.to_numeric(df_melted["PRICE"], errors="coerce").fillna(0)

    # --- Pivot untuk jadi format kolom per vendor ---
    df_pivot = df_melted.pivot_table(
        index=[year_col, "REGION", scope_col],
        columns="VENDOR",
        values="PRICE",
        aggfunc="sum",
        fill_value=0
    ).reset_index()

    # --- Hitung analisis kompetitif ---
    vendor_cols = [c for c in df_pivot.columns if c not in [year_col, "REGION", scope_col]]

    # Hitung lowest & second lowest
    df_pivot["1st Lowest"] = df_pivot[vendor_cols].min(axis=1)
    df_pivot["1st Vendor"] = df_pivot[vendor_cols].idxmin(axis=1)

    # Dapatkan nilai second lowest dengan replace nilai lowest sementara ke NaN
    def second_lowest(series):
        sorted_vals = series.sort_values().unique()
        return sorted_vals[1] if len(sorted_vals) > 1 else sorted_vals[0]

    def second_vendor(row):
        sorted_vendors = row[vendor_cols].sort_values()
        return sorted_vendors.index[1] if len(sorted_vendors) > 1 else sorted_vendors.index[0]

    df_pivot["2nd Lowest"] = df_pivot[vendor_cols].apply(second_lowest, axis=1)
    df_pivot["2nd Vendor"] = df_pivot[vendor_cols].apply(second_vendor, axis=1)

    # Hitung GAP antar lowest
    df_pivot["Gap 1 to 2 (%)"] = ((df_pivot["2nd Lowest"] - df_pivot["1st Lowest"]) / df_pivot["1st Lowest"] * 100).round(2)

    # Hitung median price
    df_pivot["Median Price"] = df_pivot[vendor_cols].median(axis=1)

    # Hitung selisih tiap vendor dengan median (%)
    for v in vendor_cols:
        df_pivot[f"{v} to Median (%)"] = ((df_pivot[v] - df_pivot["Median Price"]) / df_pivot["Median Price"] * 100).round(2)

    # --- Urutkan kolom agar rapi ---
    summary_cols = [
        year_col, "REGION", scope_col
    ] + vendor_cols + [
        "1st Lowest", "1st Vendor",
        "2nd Lowest", "2nd Vendor",
        "Gap 1 to 2 (%)", "Median Price"
    ] + [f"{v} to Median (%)" for v in vendor_cols]

    df_summary = df_pivot[summary_cols]

    # --- 🎯 Tambahkan dua slicer terpisah untuk 1st Vendor dan 2nd Vendor
    all_1st = sorted(df_summary["1st Vendor"].dropna().unique())
    all_2nd = sorted(df_summary["2nd Vendor"].dropna().unique())

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
        df_filtered = df_summary[
            df_summary["1st Vendor"].isin(selected_1st) &
            df_summary["2nd Vendor"].isin(selected_2nd)
        ]
    elif selected_1st:
        df_filtered = df_summary[df_summary["1st Vendor"].isin(selected_1st)]
    elif selected_2nd:
        df_filtered = df_summary[df_summary["2nd Vendor"].isin(selected_2nd)]
    else:
        df_filtered = df_summary.copy()

    # --- Format rupiah & persentase hanya untuk df_filtered
    num_cols = df_filtered.select_dtypes(include=["number"]).columns
    format_dict = {col: format_rupiah for col in num_cols}
    format_dict.update({"Gap 1 to 2 (%)": "{:.1f}%"})
    for v in vendor_cols:
        format_dict[f"{v} to Median (%)"] = "{:+.1f}%"

    df_summary_styled = (
        df_filtered.style
        .format(format_dict)
        .apply(lambda row: highlight_1st_2nd_vendor(row, df_filtered.columns), axis=1)
    )

    # --- Tampilkan hasil ---
    st.caption(f"✨ Total number of data entries: **{len(df_filtered)}**")
    st.dataframe(df_summary_styled, hide_index=True)

    # Simpan hasil ke variabel
    excel_data = get_excel_download_highlight_1st_2nd_lowest(df_filtered)

    # Layout tombol (rata kanan)
    col1, col2, col3 = st.columns([2.3,2,1])
    with col3:
        st.download_button(
            label="Download",
            data=excel_data,
            file_name=f"Bid & Price Analysis - Year Region.xlsx",
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
        df_summary[vendor_cols]
        .notna()       # True kalau vendor berpartisipasi (ada harga)
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
    df_summary = win_rate.rename(columns={
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
            st.dataframe(df_summary, hide_index=True)

            # Simpan hasil ke variabel
            excel_data = get_excel_download(df_summary, sheet_name="Win Rate Trend Summary")

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
    df_gap = df_pivot.copy()

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
