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

# Download Highlight 1st & 2nd Vendors
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
    st.header("1️⃣ TCO Comparison by Year")
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
    # st.subheader("📂 Upload File")
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

        # Deteksi kolom yang mengandung "total" (case-insensitive)
        total_cols = [col for col in df_clean.columns if "total" in col.lower()]

        if len(total_cols) == 0:
            # Jika tidak ada kolom total -> buat baru
            df_clean["TOTAL"] = df_clean.sum(axis=1, numeric_only=True)

        else:
            # Jika sudah ada -> rename hanya kolom pertama yg cocok
            first_total_col = total_cols[0]
            df_clean = df_clean.rename(columns={first_total_col: "TOTAL"})

        # # Tambah kolom total
        # if "TOTAL" not in df_clean.columns:
        #     df_clean["TOTAL"] = df_clean.sum(axis=1, numeric_only=True)

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

        # Tambahkan baris total di akhir
        total_row = {df_temp.columns[0]: "TOTAL"}
        for col in df_temp.columns[1:]:
            total_row[col] = df_temp[col].sum(numeric_only=True)
        df_temp = pd.concat([df_temp, pd.DataFrame([total_row])], ignore_index=True)

        df_temp.insert(0, "VENDOR", vendor)
        merged_list.append(df_temp)

    # Gabungkan semua vendor jadi satu DataFrame
    merged_overview = pd.concat(merged_list, ignore_index=True)

    # Pastikan kolom berurutan (vendor as index-0)
    cols = ["VENDOR"] + [c for c in merged_overview.columns if c != "VENDOR"]
    merged_overview = merged_overview[cols]

    # Simpan session
    st.session_state["merge_overview_tco_by_year"] = merged_overview

    # Format rupiah dan tampilkan
    num_cols = merged_overview.select_dtypes(include=["number"]).columns
    merged_overview_styled = (
        merged_overview.style
        .format({col: format_rupiah for col in num_cols})
        .apply(highlight_total_row_v2, axis=1)
    )
    st.dataframe(merged_overview_styled, hide_index=True)

    # Download
    excel_data = get_excel_download_highlight_total(merged_overview, sheet_name="TCO Overview")
    col1, col2, col3 = st.columns([2.3,2,1])
    with col3:
        st.download_button(
            label="Download",
            data=excel_data,
            file_name="TCO_Overview_Highlighted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            icon=":material/download:",
        )

    st.divider()

    # TCO SUMMARY
    st.markdown("##### 💸 TCO Summary")
    st.caption(f"The following table presents the summary of the analysis results.")

    # Merge semua sheet berdasarkan TCO Component (indeks kolom pertama)
    merged = None
    ref_order = None  # simpan urutan referensi dari sheet pertama

    for i, (name, df_sub) in enumerate(result.items()):
        first_col = df_sub.columns[0]   # TCO Component
        total_col = "TOTAL"             # Total cost 5Y

        df_merge = df_sub[[first_col, total_col]].rename(columns={total_col: name})

        # Simpan urutan referensi dari sheet pertama
        if i == 0:
            ref_order = df_merge[first_col].tolist()
            merged = df_merge
        else:
            merged = merged.merge(df_merge, on=first_col, how="outer")

    # Reorder baris sesuai urutan dari sheet pertama
    if ref_order is not None:
        merged[first_col] = pd.Categorical(merged[first_col], categories=ref_order, ordered=True)
        merged = merged.sort_values(first_col).reset_index(drop=True)

    # Menambahkan baris total di akhir
    total_row = {first_col: "TOTAL"}
    for col in merged.columns[1:]:
        total_row[col] = merged[col].sum(numeric_only=True)

    merged_total = pd.concat([merged, pd.DataFrame([total_row])], ignore_index=True)

    # Fomat Rupiah & fungsi untuk styling baris TOTAL
    num_cols = merged_total.select_dtypes(include=["number"]).columns

    merged_styled = (
        merged_total.style
        .format({col: format_rupiah for col in num_cols})
        .apply(highlight_total_row, axis=1)
    )

    st.session_state["merged_tco_by_year"] = merged_total
    st.dataframe(merged_styled, hide_index=True)

    # Download button to Excel
    excel_data = get_excel_download(merged_total)

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
            file_name="Merged_TCO.xlsx",
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

            # Kalikan semua kolom numerik (selain kolom pertama)
            df_converted = merged.copy(deep=True)
            num_cols = df_converted.columns[1:]
            df_converted.loc[:, num_cols] = (
                df_converted.loc[:, num_cols]
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
            st.caption(f"Summary of total bidders after currency conversion to {currency} at a rate of {amount}.")

            # Menambahkan baris total di akhir
            total_row = {first_col: "TOTAL"}
            for col in df_converted.columns[1:]:
                total_row[col] = df_converted[col].sum(numeric_only=True)

            converted_total = pd.concat([df_converted, pd.DataFrame([total_row])], ignore_index=True)

            # Fomat Rupiah & fungsi untuk styling baris TOTAL
            num_cols = converted_total.select_dtypes(include=["number"]).columns

            converted_styled = (
                converted_total.style
                .format({col: format_rupiah for col in num_cols})
                .apply(highlight_total_row, axis=1)
            )

            st.dataframe(converted_styled, hide_index=True)

            # Download button to Excel
            excel_data_converted = get_excel_download(converted_total)

            # Layout tombol (rata kanan)
            col1, col2, col3 = st.columns([2.3,2,1])
            with col3:
                st.download_button(
                    label="Download",
                    data=excel_data_converted,
                    file_name="Converted_TCO.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    icon=":material/download:",
                )
    
    st.divider()

    # BID & PRICE ANALYSIS
    st.markdown("##### 🧠 Bid & Price Analysis")
    # st.caption("Comparative analysis across vendors including lowest price, gap percentage, and deviation from median.")

    df_analysis = merged_total.copy()

    # Ambil nama kolom vendor (semua kecuali kolom pertama)
    tco_col = df_analysis.columns[0]
    vendor_cols = [c for c in df_analysis.columns if c != tco_col]

    # Hapus baris TOTAL untuk analisis per komponen
    df_no_total = df_analysis[df_analysis[tco_col].str.upper() != "TOTAL"].copy()

    # Hitung nilai analisis per baris
    df_no_total["1st Lowest"] = df_no_total[vendor_cols].min(axis=1)
    df_no_total["1st Vendor"] = df_no_total[vendor_cols].idxmin(axis=1)

    # Kedua terendah (2nd lowest)
    def second_lowest(row):
        sorted_vals = sorted([(v, k) for k, v in row[vendor_cols].items() if pd.notnull(v)])
        if len(sorted_vals) > 1:
            return sorted_vals[1]
        return (np.nan, np.nan)

    df_no_total[["2nd Lowest", "2nd Vendor"]] = df_no_total.apply(
        lambda row: pd.Series(second_lowest(row)), axis=1
    )

    # Hitung Gap 1 to 2 (%)
    df_no_total["Gap 1 to 2 (%)"] = (
        (df_no_total["2nd Lowest"] - df_no_total["1st Lowest"])
        / df_no_total["1st Lowest"] * 100
    ).round(2)

    # Hitung Median Price
    df_no_total["Median Price"] = df_no_total[vendor_cols].median(axis=1)

    # Hitung deviasi tiap vendor terhadap median
    for v in vendor_cols:
        df_no_total[f"{v} to Median (%)"] = (
            (df_no_total[v] - df_no_total["Median Price"])
            / df_no_total["Median Price"] * 100
        ).round(2)

    # Urutkan kolom sesuai struktur yang diinginkan
    analysis_cols = (
        [tco_col]
        + vendor_cols
        + ["1st Lowest", "1st Vendor", "2nd Lowest", "2nd Vendor", "Gap 1 to 2 (%)", "Median Price"]
        + [f"{v} to Median (%)" for v in vendor_cols]
    )

    df_analysis_final = df_no_total[analysis_cols]

    # Format rupiah
    num_cols = df_analysis_final.select_dtypes(include=["number"]).columns
    format_dic = {col: format_rupiah for col in num_cols}
    format_dic.update({"Gap 1 to 2 (%)": "{:.1f}%"})
    for v in vendor_cols:
        format_dic[f"{v} to Median (%)"] = "{:+.1f}%"

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
        df_filtered = df_analysis_final[
            df_analysis_final["1st Vendor"].isin(selected_1st) &
            df_analysis_final["2nd Vendor"].isin(selected_2nd)
        ]
    elif selected_1st:
        df_filtered = df_analysis_final[df_analysis_final["1st Vendor"].isin(selected_1st)]
    elif selected_2nd:
        df_filtered = df_analysis_final[df_analysis_final["2nd Vendor"].isin(selected_2nd)]
    else:
        df_filtered = df_analysis_final.copy()

    df_analysis_styled = (
        df_filtered.style
        .format(format_dic)
        .apply(lambda row: highlight_1st_2nd_vendor(row, df_analysis_final.columns), axis=1)
    )

    st.caption(f"✨ Total number of data entries: **{len(df_filtered)}**")
    st.dataframe(df_analysis_styled, hide_index=True)

    # Download button to Excel
    excel_data = get_excel_download_highlight_1st_2nd_lowest(df_filtered, sheet_name="Bid & Price Analysis")

    # Layout tombol (rata kanan)
    col1, col2, col3 = st.columns([2.3,2,1])
    with col3:
        st.download_button(
            label="Download",
            data=excel_data,
            file_name="Bid & Price Analysis.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            icon=":material/download:",
        )

    # RANK
    if merged is not None:
        st.markdown("##### 🏅 Rank Visualization")
        st.caption("Rangking is generated based on each vendor's overall total cost.")
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
        st.caption("Each chart shows total offer comparison and ranking for every requirement.")
        if merged is not None and not merged.empty:
            tab3, tab4 = st.tabs(["💲Original Price", "💱 Converted Price"])
            reqs = merged[merged.columns[0]]
            vendors = merged.columns[1:]

            tab3.caption("")

            n_cols = 2
            for i in range(0, len(reqs), n_cols):
                cols = tab3.columns(n_cols)

                for j in range(n_cols):
                    if i + j < len(reqs):
                        req = reqs[i + j]

                        # --- Data vendor + harga ---
                        prices = merged.loc[i + j, vendors].reset_index()
                        prices.columns = ["Vendor", "Total"]

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
                vendors = df_converted.columns[1:]

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