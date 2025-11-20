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

# Download button to Excel
@st.cache_data
def get_excel_download(df, sheet_name="Sheet1"):
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

# Download Highlight Excel
def get_excel_download_highlight_1st_2nd_lowest(df, sheet_name="Sheet1"):
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
    st.header("2️⃣ TCO Comparison by Region")
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
        df_clean = df.replace(r'^\s*$', None, regex=True)
        df_clean = df_clean.dropna(how="all", axis=0).dropna(how="all", axis=1)

        rows_after, cols_after = df_clean.shape

        # Gunakan baris pertama sebagai header (hanya jika kolom belum ada nama atau Unnamed)
        if any("Unnamed" in str(c) for c in df_clean.columns):
            df_clean.columns = df_clean.iloc[0]
            df_clean = df_clean[1:].reset_index(drop=True)

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

        # Tambahkan baris total di akhir
        total_row = {df_temp.columns[0]: "TOTAL"}
        for col in df_temp.columns[1:]:
            total_row[col] = df_temp[col].sum(numeric_only=True)
        df_temp = pd.concat([df_temp, pd.DataFrame([total_row])], ignore_index=True)

        df_temp.insert(0, "VENDOR", vendor_name)  # Tambahkan kolom vendor di depan
        merged_df.append(df_temp)

    df_all_vendors = pd.concat(merged_df, ignore_index=True)

    # Simpan ke session_state jika perlu digunakan di halaman lain
    st.session_state["merged_all_data_tco_by_region"] = df_all_vendors

    # Format Rupiah
    num_cols = df_all_vendors.select_dtypes(include=["number"]).columns
    df_all_vendors_styled = (
        df_all_vendors.style
        .format({col: format_rupiah for col in num_cols})
        .apply(highlight_total_row_v2, axis=1)
    )

    tab1.caption(f"Data from **{total_sheets} vendors** have been successfully consolidated, analyzing a total of **{len(df_all_vendors):,} combined records**.")
    tab1.dataframe(df_all_vendors_styled, hide_index=True)

    # Download button to Excel
    excel_data = get_excel_download_highlight_total(df_all_vendors, sheet_name="Merge Original Data")

    with tab1:
        # Layout tombol (rata kanan)
        col1, col2, col3 = st.columns([2.3,2,1])
        with col3:
            st.download_button(
                label="Download",
                data=excel_data,
                file_name="Merged_All_Dataset.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                icon=":material/download:",
            )

    tab1.divider()    

    # TCO PER SCOPEE
    tab1.markdown("##### 💸 TCO Summary — Scope")
    tab1.caption("The following table presents the TCO summary across all scopes.")

    # Merge semua sheet berdasarkan TCO Component (indeks kolom pertama)
    merged_tco = None
    ref_order = None  # simpan urutan referensi dari sheet pertama

    for i, (name, df_sub) in enumerate(result.items()):
        first_col = df_sub.columns[0]   # TCO Component
        total_col = "TOTAL"
        df_merge_tco = df_sub[[first_col, total_col]].rename(columns={total_col: name})

        # Simpan urutan referensi dari sheet pertama
        if i == 0:
            ref_order = df_merge_tco[first_col].tolist()
            merged_tco = df_merge_tco
        else:
            merged_tco = merged_tco.merge(df_merge_tco, on=first_col, how="outer")

    # Reorder baris sesuai urutan dari sheet pertama
    if ref_order is not None:
        merged_tco[first_col] = pd.Categorical(merged_tco[first_col], categories=ref_order, ordered=True)
        merged_tco = merged_tco.sort_values(first_col).reset_index(drop=True)

    # Menambahkan baris total di akhir
    total_row = {first_col: "TOTAL"}
    for col in merged_tco.columns[1:]:
        total_row[col] = merged_tco[col].sum(numeric_only=True)

    merged_tco_total = pd.concat([merged_tco, pd.DataFrame([total_row])], ignore_index=True)

    # Fomat Rupiah & fungsi untuk styling baris TOTAL
    num_cols = merged_tco_total.select_dtypes(include=["number"]).columns

    merged_tco_styled = (
        merged_tco_total.style
        .format({col: format_rupiah for col in num_cols})
        .apply(highlight_total_row, axis=1)
    )

    st.session_state["merged_tco_by_region"] = merged_tco_total
    tab1.dataframe(merged_tco_styled, hide_index=True)

    # Download button to Excel
    excel_data = get_excel_download(merged_tco_total)

    with tab1:
        # Layout tombol (rata kanan)
        col1, col2, col3 = st.columns([2.3,2,1])
        with col3:
            st.download_button(
                label="Download",
                data=excel_data,
                file_name="Merged_TCO_Region.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                icon=":material/download:",
            )

    # # REGIONN -> SCOPE COMPARISONN
    # tab1.markdown("##### 🧠 Bid & Price Analysis — Scope")

    # Ambil daftar region dari salah satu vendor (asumsi sama di semua)
    sample_vendor = next(iter(result.values()))
    region_cols = sample_vendor.columns[1:-1]  # exclude Scope & Total

    # # Loop per Region
    # tabs = tab1.tabs(region_cols.tolist())

    # for i, region in enumerate(region_cols):
    #     with tabs[i]:
    #         df_region = pd.DataFrame()
    #         ref_order = None  # simpan urutan scope vendor pertama

    #         # Ambil kolom region dari semua vendor
    #         for j, (vendor, df_vendor) in enumerate(result.items()):
    #             scope_col = df_vendor.columns[0]
    #             df_temp = df_vendor[[scope_col, region]].copy()
    #             df_temp.rename(columns={region: f"{vendor}"}, inplace=True)

    #             # Simpan urutan scope vendor pertama
    #             if j == 0:
    #                 ref_order = df_temp[scope_col].tolist()
    #                 df_region = df_temp
    #             else:
    #                 df_region = df_region.merge(df_temp, on=scope_col, how="outer")

    #         # Reorder kembali sesuai urutan vendor pertama
    #         if ref_order is not None:
    #             df_region[scope_col] = pd.Categorical(df_region[scope_col], categories=ref_order, ordered=True)
    #             df_region = df_region.sort_values(scope_col).reset_index(drop=True)

    #         # --- Hitung ranking harga per baris ---
    #         vendor_cols = [c for c in df_region.columns if c != scope_col]
    #         df_region["1st Lowest"] = df_region[vendor_cols].min(axis=1, numeric_only=True)
    #         df_region["1st Vendor"] = df_region[vendor_cols].idxmin(axis=1)
    #         df_region["2nd Lowest"] = df_region[vendor_cols].apply(
    #             lambda row: sorted([v for v in row if pd.notna(v)])[1]
    #             if len([v for v in row if pd.notna(v)]) > 1
    #             else None,
    #             axis=1,
    #         )
    #         df_region["2nd Vendor"] = df_region.apply(
    #             lambda row: next(
    #                 (col for col in vendor_cols if row[col] == row["2nd Lowest"]), None
    #             ),
    #             axis=1,
    #         )

    #         # Gap 1 to 2 (%)
    #         df_region["Gap 1 to 2 (%)"] = (
    #             (df_region["2nd Lowest"] - df_region["1st Lowest"]) / df_region["1st Lowest"] * 100
    #         ).round(1)

    #         # --- Hitung Median & Deviasi ---
    #         df_region["Median Price"] = df_region[vendor_cols].median(axis=1, numeric_only=True)
    #         for vendor in vendor_cols:
    #             df_region[f"{vendor} to Median (%)"] = (
    #                 (df_region[vendor] - df_region["Median Price"]) / df_region["Median Price"] * 100
    #             ).round(1)

    #         # --- 🎯 Tambahkan dua slicer terpisah untuk 1st Vendor dan 2nd Vendor
    #         all_1st = sorted(df_region["1st Vendor"].dropna().unique())
    #         all_2nd = sorted(df_region["2nd Vendor"].dropna().unique())

    #         col_sel_1, col_sel_2 = st.columns(2)
    #         with col_sel_1:
    #             selected_1st = st.multiselect(
    #                 "Filter: 1st vendor",
    #                 options=all_1st,
    #                 default=None,
    #                 placeholder="Choose one or more vendors",
    #                 key=f"filter_1st_vendor_{region}"
    #             )
    #         with col_sel_2:
    #             selected_2nd = st.multiselect(
    #                 "Filter: 2nd vendor",
    #                 options=all_2nd,
    #                 default=None,
    #                 placeholder="Choose one or more vendors",
    #                 key=f"filter_2nd_vendor_{region}"
    #             )

    #         # --- Terapkan filter dengan logika AND
    #         if selected_1st and selected_2nd:
    #             df_filtered = df_region[
    #                 df_region["1st Vendor"].isin(selected_1st) &
    #                 df_region["2nd Vendor"].isin(selected_2nd)
    #             ]
    #         elif selected_1st:
    #             df_filtered = df_region[df_region["1st Vendor"].isin(selected_1st)]
    #         elif selected_2nd:
    #             df_filtered = df_region[df_region["2nd Vendor"].isin(selected_2nd)]
    #         else:
    #             df_filtered = df_region.copy()

    #         # --- Format rupiah & persentase hanya untuk df_filtered
    #         num_cols = df_filtered.select_dtypes(include=["number"]).columns
    #         format_dict = {col: format_rupiah for col in num_cols}
    #         format_dict.update({"Gap 1 to 2 (%)": "{:.1f}%"})
    #         for v in vendor_cols:
    #             format_dict[f"{v} to Median (%)"] = "{:+.1f}%"

    #         df_styled = (
    #             df_filtered.style
    #             .format(format_dict)
    #             .apply(lambda row: highlight_1st_2nd_vendor(row, df_filtered.columns), axis=1)
    #         )

    #         st.caption(f"✨ Total number of data entries: **{len(df_filtered)}**")
    #         st.dataframe(df_styled, hide_index=True)

    #         # Simpan hasil ke variabel
    #         excel_data = get_excel_download_highlight_1st_2nd_lowest(df_filtered, sheet_name="Regional Comparison")

    #         # Layout tombol (rata kanan)
    #         col1, col2, col3 = st.columns([2.3,2,1])
    #         with col3:
    #             st.download_button(
    #                 label="Download",
    #                 data=excel_data,
    #                 file_name=f"Regional Comparison - {region}.xlsx",
    #                 mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    #                 icon=":material/download:",
    #                 key=f"download_regional_{region}_comparison"  # unik per tab
    #             )

    tab1.divider()
    
    # GABUNGG SEMUA REGION JADI SATUU
    # tab1.markdown("##### 🌍 Regional Bid & Price Summary")
    tab1.markdown("##### 🧠 Bid & Price Summary Analysis — Region")

    all_regions_combined = []

    for region in region_cols:
        df_region = pd.DataFrame()
        ref_order = None

        # Ambil kolom region dari semua vendor
        for j, (vendor, df_vendor) in enumerate(result.items()):
            scope_col = df_vendor.columns[0]
            df_temp = df_vendor[[scope_col, region]].copy()
            df_temp.rename(columns={region: f"{vendor}"}, inplace=True)

            if j == 0:
                ref_order = df_temp[scope_col].tolist()
                df_region = df_temp
            else:
                df_region = df_region.merge(df_temp, on=scope_col, how="outer")

        # Urutkan sesuai vendor pertama
        if ref_order is not None:
            df_region[scope_col] = pd.Categorical(df_region[scope_col], categories=ref_order, ordered=True)
            df_region = df_region.sort_values(scope_col).reset_index(drop=True)

        # Hitung metrik seperti sebelumnya
        vendor_cols = [c for c in df_region.columns if c != scope_col]
        df_region["1st Lowest"] = df_region[vendor_cols].min(axis=1, numeric_only=True)
        df_region["1st Vendor"] = df_region[vendor_cols].idxmin(axis=1)
        df_region["2nd Lowest"] = df_region[vendor_cols].apply(
            lambda row: sorted([v for v in row if pd.notna(v)])[1]
            if len([v for v in row if pd.notna(v)]) > 1 else None,
            axis=1
        )
        df_region["2nd Vendor"] = df_region.apply(
            lambda row: next((col for col in vendor_cols if row[col] == row["2nd Lowest"]), None),
            axis=1
        )
        df_region["Gap 1 to 2 (%)"] = (
            (df_region["2nd Lowest"] - df_region["1st Lowest"]) / df_region["1st Lowest"] * 100
        ).round(1)

        # Median & Deviasi
        df_region["Median Price"] = df_region[vendor_cols].median(axis=1, numeric_only=True)
        for vendor in vendor_cols:
            df_region[f"{vendor} to Median (%)"] = (
                (df_region[vendor] - df_region["Median Price"]) / df_region["Median Price"] * 100
            ).round(1)

        # Tambahkan kolom region
        df_region.insert(0, "REGION", region)

        # Simpan ke list
        all_regions_combined.append(df_region)

    # --- Gabungkan semua region jadi satu DataFrame besar ---
    df_all_regions = pd.concat(all_regions_combined, ignore_index=True)

    # --- 🎯 Tambahkan slicer
    all_1st = sorted(df_all_regions["1st Vendor"].dropna().unique())
    all_2nd = sorted(df_all_regions["2nd Vendor"].dropna().unique())

    with tab1:
        col_sel_1, col_sel_2 = st.columns(2)
        with col_sel_1:
            selected_1st = st.multiselect(
                "Filter: 1st vendor",
                options=all_1st,
                default=[],
                placeholder="Choose vendors",
                key="filter_1st_region"
            )
        with col_sel_2:
            selected_2nd = st.multiselect(
                "Filter: 2nd vendor",
                options=all_2nd,
                default=[],
                placeholder="Choose vendors",
                key="filter_2nd_region"
            )

        # --- Terapkan filter AND secara dinamis
        df_filtered_region = df_all_regions.copy()

        if selected_1st:
            df_filtered_region = df_filtered_region[df_filtered_region["1st Vendor"].isin(selected_1st)]

        if selected_2nd:
            df_filtered_region = df_filtered_region[df_filtered_region["2nd Vendor"].isin(selected_2nd)]

    # --- Styling (Rupiah & Persen) ---
    num_cols = df_filtered_region.select_dtypes(include=["number"]).columns
    format_dict = {col: format_rupiah for col in num_cols}
    format_dict.update({"Gap 1 to 2 (%)": "{:.1f}%"})
    for vendor in vendor_cols:
        format_dict[f"{vendor} to Median (%)"] = "{:+.1f}%"

    df_filtered_region_styled = (
        df_filtered_region.style
        .format(format_dict)
        .apply(lambda row: highlight_1st_2nd_vendor(row, df_filtered_region.columns), axis=1)
    )

    # --- Tampilkan di Streamlit ---
    tab1.caption(f"Successfully consolidated all {len(region_cols)} regional tabs into **{len(df_filtered_region):,} total rows**.")
    tab1.dataframe(df_filtered_region_styled, hide_index=True)

    excel_data = get_excel_download_highlight_1st_2nd_lowest(df_filtered_region)
    with tab1:
        # --- Optional: tombol download ---
        col1, col2, col3 = st.columns([2.3,2,1])
        with col3:
            st.download_button(
                label="Download",
                data=excel_data,
                file_name="Regional Comparison - ALL.xlsx",
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
            win1_counts = df_all_regions["1st Vendor"].value_counts(dropna=True).reset_index()
            win1_counts.columns = ["Vendor", "Wins1"]

            win2_counts = df_all_regions["2nd Vendor"].value_counts(dropna=True).reset_index()
            win2_counts.columns = ["Vendor", "Wins2"]

            # --- Hitung total partisipasi vendor ---
            vendor_counts = (
                df_all_regions[vendor_cols]
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
                        key="download_Win_Rate_Trend_Tab1"
                    )

        with subtab2:
            # --- AVERAGE GAP VISUALIZATION ---
            # --- Hitung rata-rata Gap 1 to 2 (%) per Vendor (hanya untuk 1st Vendor)
            df_gap = df_all_regions.copy()

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

        # Gunakan kolom pertama sebagai index (dinamis)
        first_col = df_clean_transposed.columns[0]
        df_clean_transposed = df_clean_transposed.set_index(first_col)

        # Transpose
        df_transposed = df_clean_transposed.T.reset_index().rename(columns={"index": "REGION"})

        # Tambah kolom total
        if "TOTAL" not in df_transposed.columns:
            df_transposed["TOTAL"] = df_transposed.sum(axis=1, numeric_only=True)

        # Pilih kolom numeric baru setelah transpose
        num_cols_transposed = df_transposed.select_dtypes(include=["number"]).columns

        # Format Rupiah
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
    
    #     # Download button per sheet
    #     excel_data = get_excel_download(df_transposed, sheet_name=name)

    #     # st.markdown("<div style='margin-top:0px'></div>", unsafe_allow_html=True)
    #     col1, col2, col3 = tab2.columns([2.3,2,1])
    #     with col3:
    #         tab2.download_button(
    #             label="Download",
    #             data=excel_data,
    #             file_name="Transposed_{name}.xlsx",
    #             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    #             icon=":material/download:",
    #             key=f"download_transposed_{name}"
    #         )
    #         tab2.write("")

    # tab2.divider()

    # MERGE TRANSPOSED
    tab2.markdown("##### 🗃️ Merge Transposed")

    # --- Merge all vendor data into one DataFrame ---
    merged_df_transposed = []

    for vendor_name, df_vendor in result.items():
        df_temp = df_vendor.copy()

        # Tambahkan baris total di akhir
        total_row = {df_temp.columns[0]: "TOTAL"}
        for col in df_temp.columns[1:]:
            total_row[col] = df_temp[col].sum(numeric_only=True)
        df_temp = pd.concat([df_temp, pd.DataFrame([total_row])], ignore_index=True)

        df_temp.insert(0, "VENDOR", vendor_name)  # Tambahkan kolom vendor di depan
        merged_df_transposed.append(df_temp)

    df_all_vendors_transposed = pd.concat(merged_df_transposed, ignore_index=True)

    # Simpan ke session_state jika perlu digunakan di halaman lain
    st.session_state["merged_all_data_transposed_tco_by_region"] = df_all_vendors_transposed

    # Format Rupiah
    num_cols = df_all_vendors_transposed.select_dtypes(include=["number"]).columns
    df_all_vendors_transposed_styled = (
        df_all_vendors_transposed.style
        .format({col: format_rupiah for col in num_cols})
        .apply(highlight_total_row_v2, axis=1)
    )

    tab2.caption(f"Data from **{total_sheets} vendors** have been successfully consolidated, analyzing a total of **{len(df_all_vendors_transposed):,} combined records**.")
    tab2.dataframe(df_all_vendors_transposed_styled, hide_index=True)

    # Download button to Excel
    excel_data = get_excel_download_highlight_total(df_all_vendors_transposed)

    with tab2:
        # Layout tombol (rata kanan)
        col1, col2, col3 = st.columns([2.3,2,1])
        with col3:
            st.download_button(
                label="Download",
                data=excel_data,
                file_name="Merged_All_Transposed_Dataset.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                icon=":material/download:",
            )

    tab2.divider()

    # TCO per Regionn
    tab2.markdown("##### 💸 TCO Summary — Region")
    tab2.caption(f"The following table presents the TCO summary across all regions.")
    
    # Merge semua sheet berdasarkan TCO Component (indeks kolom pertama)
    merged_transposed = None
    ref_order = None  # simpan urutan referensi dari sheet pertama

    for i, (name, df_sub) in enumerate(result.items()):
        # Kolom pertama (Region)
        first_col = df_sub.columns[0]
        total_col = "TOTAL"
        df_merge_transposed = df_sub[[first_col, total_col]].rename(columns={total_col: name})

        # Simpan urutan referensi dari sheet pertama
        if i == 0:
            ref_order = df_merge_transposed[first_col].tolist()
            merged_transposed = df_merge_transposed
        else:
            merged_transposed = merged_transposed.merge(df_merge_transposed, on=first_col, how="outer")

    # Reorder Region sesuai sheet pertama
    if ref_order is not None:
        merged_transposed[first_col] = pd.Categorical(merged_transposed[first_col], categories=ref_order, ordered=True)
        merged_transposed = merged_transposed.sort_values(first_col).reset_index(drop=True)

    # Tambahkan baris TOTAL di akhir
    total_row = {first_col: "TOTAL"}
    for col in merged_transposed.columns[1:]:
        total_row[col] = merged_transposed[col].sum(numeric_only=True)

    merged_transposed_total = pd.concat([merged_transposed, pd.DataFrame([total_row])], ignore_index=True)

    # Format Rupiah & highlight baris TOTAL
    num_cols = merged_transposed_total.select_dtypes(include=["number"]).columns

    merged_transposed_styled = (
        merged_transposed_total.style
        .format({col: format_rupiah for col in num_cols})
        .apply(highlight_total_row, axis=1)
    )

    st.session_state["merged_transposed_tco_by_region"] = merged_transposed_total
    tab2.dataframe(merged_transposed_styled, hide_index=True)

    # Download button to Excel
    excel_data = get_excel_download(merged_transposed_total)

    with tab2:
        # Layout tombol (rata kanan)
        col1, col2, col3 = st.columns([2.3,2,1])
        with col3:
            st.download_button(
                label="Download",
                data=excel_data,
                file_name="Merged_Transformed_Region.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                icon=":material/download:",
            )

    # # SCOPEE -> REGION COMPARISONN
    # tab2.markdown("##### 🧠 Bid & Price Analysis — Region")

    # Ambil daftar scope dari salah satu vendor (asumsi sama di semua)
    sample_vendor = next(iter(result.values()))
    scope_cols = sample_vendor.columns[1:-1]  # kolom setelah "Region"

    # # Loop per Scope
    # tabs = tab2.tabs(scope_cols.tolist())

    # for i, scope in enumerate(scope_cols):
    #     with tabs[i]:
    #         df_scope = pd.DataFrame()
    #         ref_order = None  # simpan urutan scope vendor pertama

    #         # Ambil kolom scope dari semua vendor
    #         for j, (vendor, df_vendor) in enumerate(result.items()):
    #             scope_col = df_vendor.columns[0]
    #             df_temp = df_vendor[[scope_col, scope]].copy()
    #             df_temp.rename(columns={scope: f"{vendor}"}, inplace=True)

    #             # Simpan urutan scope vendor pertama
    #             if j == 0:
    #                 ref_order = df_temp[scope_col].tolist()
    #                 df_scope = df_temp
    #             else:
    #                 df_scope = df_scope.merge(df_temp, on=scope_col, how="outer")

    #         # Reorder kembali sesuai urutan vendor pertama
    #         if ref_order is not None:
    #             df_scope[scope_col] = pd.Categorical(df_scope[scope_col], categories=ref_order, ordered=True)
    #             df_scope = df_scope.sort_values(scope_col).reset_index(drop=True)

    #         # --- Hitung ranking harga per baris ---
    #         vendor_cols = [c for c in df_scope.columns if c != scope_col]
    #         df_scope["1st Lowest"] = df_scope[vendor_cols].min(axis=1, numeric_only=True)
    #         df_scope["1st Vendor"] = df_scope[vendor_cols].idxmin(axis=1)
    #         df_scope["2nd Lowest"] = df_scope[vendor_cols].apply(
    #             lambda row: sorted([v for v in row if pd.notna(v)])[1]
    #             if len([v for v in row if pd.notna(v)]) > 1
    #             else None,
    #             axis=1,
    #         )
    #         df_scope["2nd Vendor"] = df_scope.apply(
    #             lambda row: next(
    #                 (col for col in vendor_cols if row[col] == row["2nd Lowest"]), None
    #             ),
    #             axis=1,
    #         )

    #         # Gap 1 to 2 (%)
    #         df_scope["Gap 1 to 2 (%)"] = (
    #             (df_scope["2nd Lowest"] - df_scope["1st Lowest"]) / df_scope["1st Lowest"] * 100
    #         ).round(1)

    #         # --- Hitung Median & Deviasi ---
    #         df_scope["Median Price"] = df_scope[vendor_cols].median(axis=1, numeric_only=True)
    #         for vendor in vendor_cols:
    #             df_scope[f"{vendor} to Median (%)"] = (
    #                 (df_scope[vendor] - df_scope["Median Price"]) / df_scope["Median Price"] * 100
    #             ).round(1)

    #         # --- 🎯 Tambahkan dua slicer terpisah untuk 1st Vendor dan 2nd Vendor
    #         all_1st = sorted(df_scope["1st Vendor"].dropna().unique())
    #         all_2nd = sorted(df_scope["2nd Vendor"].dropna().unique())

    #         col_sel_1, col_sel_2 = st.columns(2)
    #         with col_sel_1:
    #             selected_1st = st.multiselect(
    #                 "Filter: 1st vendor",
    #                 options=all_1st,
    #                 default=None,
    #                 placeholder="Choose one or more vendors",
    #                 key=f"filter_1st_vendor_{scope}"
    #             )
    #         with col_sel_2:
    #             selected_2nd = st.multiselect(
    #                 "Filter: 2nd vendor",
    #                 options=all_2nd,
    #                 default=None,
    #                 placeholder="Choose one or more vendors",
    #                 key=f"filter_2nd_vendor_{scope}"
    #             )

    #         # --- Terapkan filter dengan logika AND
    #         if selected_1st and selected_2nd:
    #             df_filtered = df_scope[
    #                 df_scope["1st Vendor"].isin(selected_1st) &
    #                 df_scope["2nd Vendor"].isin(selected_2nd)
    #             ]
    #         elif selected_1st:
    #             df_filtered = df_scope[df_scope["1st Vendor"].isin(selected_1st)]
    #         elif selected_2nd:
    #             df_filtered = df_scope[df_scope["2nd Vendor"].isin(selected_2nd)]
    #         else:
    #             df_filtered = df_scope.copy()

    #         # --- Format rupiah & persentase hanya untuk df_filtered
    #         num_cols = df_filtered.select_dtypes(include=["number"]).columns
    #         format_dict = {col: format_rupiah for col in num_cols}
    #         format_dict.update({"Gap 1 to 2 (%)": "{:.1f}%"})
    #         for v in vendor_cols:
    #             format_dict[f"{v} to Median (%)"] = "{:+.1f}%"

    #         df_styled = (
    #             df_filtered.style
    #             .format(format_dict)
    #             .apply(lambda row: highlight_1st_2nd_vendor(row, df_filtered.columns), axis=1)
    #         )

    #         st.caption(f"✨ Total number of data entries: **{len(df_filtered)}**")
    #         st.dataframe(df_styled, hide_index=True)

    #         # Download button to Excel
    #         excel_data = get_excel_download_highlight_1st_2nd_lowest(df_filtered, sheet_name="Scope Comparison")

    #         # Layout tombol (rata kanan)
    #         col1, col2, col3 = st.columns([2.3,2,1])
    #         with col3:
    #             st.download_button(
    #                 label="Download",
    #                 data=excel_data,
    #                 file_name=f"Scope Comparison - {scope}.xlsx",
    #                 mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    #                 icon=":material/download:",
    #                 key=f"download_Scope_{scope}_Comparison"  # unik per tab
    #             )

    tab2.divider()

    # GABUNGG SEMUA SCOPE JADI SATU
    # tab2.markdown("##### 🌍 Scope Bid & Price Summary")
    tab2.markdown("##### 🧠 Bid & Price Summary Analysis — Scope")

    all_scopes_combined = []

    for scope in scope_cols:
        df_scope = pd.DataFrame()
        ref_order = None

        # Ambil kolom scope dari semua vendor
        for j, (vendor, df_vendor) in enumerate(result.items()):
            scope_col = df_vendor.columns[0]
            df_temp = df_vendor[[scope_col, scope]].copy()
            df_temp.rename(columns={scope: f"{vendor}"}, inplace=True)

            if j == 0:
                ref_order = df_temp[scope_col].tolist()
                df_scope = df_temp 
            else: 
                df_scope = df_scope.merge(df_temp, on=scope_col, how="outer")
        
        # Urutkan sesuai vendor pertama
        if ref_order is not None:
            df_scope[scope_col] = pd.Categorical(df_scope[scope_col], categories=ref_order, ordered=True)
            df_scope = df_scope.sort_values(scope_col).reset_index(drop=True)

        # Hitung metrik seperti sebelumnya
        vendor_cols = [c for c in df_scope.columns if c != scope_col]
        df_scope["1st Lowest"] = df_scope[vendor_cols].min(axis=1, numeric_only=True)
        df_scope["1st Vendor"] = df_scope[vendor_cols].idxmin(axis=1)
        df_scope["2nd Lowest"] = df_scope[vendor_cols].apply(
            lambda row: sorted([v for v in row if pd.notna(v)])[1]
            if len([v for v in row if pd.notna(v)]) > 1 else None,
            axis=1
        )
        df_scope["2nd Vendor"] = df_scope.apply(
            lambda row: next((col for col in vendor_cols if row[col] == row["2nd Lowest"]), None),
            axis=1
        )
        df_scope["Gap 1 to 2 (%)"] = (
            (df_scope["2nd Lowest"] - df_scope["1st Lowest"]) / df_scope["1st Lowest"] * 100
        ).round(1)

        # Median & Deviasi
        df_scope["Median Price"] = df_scope[vendor_cols].median(axis=1, numeric_only=True)
        for vendor in vendor_cols:
            df_scope[f"{vendor} to Median (%)"] = (
                (df_scope[vendor] - df_scope["Median Price"]) / df_scope["Median Price"] * 100
            ).round(1)

        # Tambahkan kolom scope
        df_scope.insert(0, "SCOPE", scope)
        all_scopes_combined.append(df_scope)

    # Gabungkan semua scope jadi satu DataFrame besar
    df_all_scopes = pd.concat(all_scopes_combined, ignore_index=True)

    # --- 🎯 Tambahkan slicer
    all_1st = sorted(df_all_scopes["1st Vendor"].dropna().unique())
    all_2nd = sorted(df_all_scopes["2nd Vendor"].dropna().unique())

    with tab2:
        col_sel_1, col_sel_2 = st.columns(2)
        with col_sel_1:
            selected_1st = st.multiselect(
                "Filter: 1st vendor",
                options=all_1st,
                default=[],
                placeholder="Choose vendors",
                key="filter_1st_scope"
            )
        with col_sel_2:
            selected_2nd = st.multiselect(
                "Filter: 2nd vendor",
                options=all_2nd,
                default=[],
                placeholder="Choose vendors",
                key="filter_2nd_scope"
            )

        # --- Terapkan filter AND secara dinamis
        df_filtered_scope = df_all_scopes.copy()

        if selected_1st:
            df_filtered_scope = df_filtered_scope[df_filtered_scope["1st Vendor"].isin(selected_1st)]

        if selected_2nd:
            df_filtered_scope = df_filtered_scope[df_filtered_scope["2nd Vendor"].isin(selected_2nd)]


    # --- Styling (Rupiah & Persen) ---
    num_cols = df_filtered_scope.select_dtypes(include=["number"]).columns
    format_dict = {col: format_rupiah for col in num_cols}
    format_dict.update({"Gap 1 to 2 (%)": "{:.1f}%"})
    for vendor in vendor_cols:
        format_dict[f"{vendor} to Median (%)"] = "{:+.1f}%"

    df_filtered_scope_styled = (
        df_filtered_scope.style
        .format(format_dict)
        .apply(lambda row: highlight_1st_2nd_vendor(row, df_filtered_scope.columns), axis=1)
    )

    tab2.caption(f"Successfully consolidated all {len(scope_cols)} regional tabs into **{len(df_filtered_scope):,} total rows**.")
    tab2.dataframe(df_filtered_scope_styled, hide_index=True)

    excel_data = get_excel_download_highlight_1st_2nd_lowest(df_filtered_scope)
    with tab2:
        col1, col2, col3 = st.columns([2.3,2,1])
        with col3:
            st.download_button(
                label="Download",
                data=excel_data,
                file_name="Scope Comparison - ALL.xlsx",
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
            win1_counts = df_all_scopes["1st Vendor"].value_counts(dropna=True).reset_index()
            win1_counts.columns = ["Vendor", "Wins1"]

            win2_counts = df_all_scopes["2nd Vendor"].value_counts(dropna=True).reset_index()
            win2_counts.columns = ["Vendor", "Wins2"]

            # --- Hitung total partisipasi vendor ---
            vendor_counts = (
                df_all_scopes[vendor_cols]
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
                        key="download_Win_Rate_Trend_Tab2"
                    )

        with subtab2:
            # --- AVERAGE GAP VISUALIZATION ---
            # --- Hitung rata-rata Gap 1 to 2 (%) per Vendor (hanya untuk 1st Vendor)
            df_gap = df_all_scopes.copy()

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
