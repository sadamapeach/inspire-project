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

    st.divider()

    # BIDD & PRICEE ANALYSIS
    st.markdown("##### 🧠 Bid & Price Analysis")

    # Hapus row total 
    df_analysis = df_final[
        ~df_final.apply(lambda row: row.astype(str).str.contains("TOTAL", case=False))
        .any(axis=1)
    ].copy()

    df_no_total = df_analysis.copy()

    non_round_cols = [c for c in df_analysis.columns if c != "ROUND"]
    scope_cols = df_analysis[non_round_cols].select_dtypes(exclude=["number"]).columns.tolist()
    vendor_cols = df_analysis[non_round_cols].select_dtypes(include=['number']).columns.tolist()
    
    # Hitung 1st & 2nd lowest
    df_analysis["1st Lowest"] = df_analysis[vendor_cols].min(axis=1)
    df_analysis["1st Vendor"] = df_analysis[vendor_cols].idxmin(axis=1)

    # untuk 2nd lowest -> sort values per row
    df_analysis["2nd Lowest"] = df_analysis[vendor_cols].apply(
        lambda row: row.nsmallest(2).iloc[-1] if len(row.dropna()) >= 2 else np.nan,
        axis=1
    )
    df_analysis["2nd Vendor"] = df_analysis[vendor_cols].apply(
        lambda row: row.nsmallest(2).index[-1] if len(row.dropna()) >= 2 else "",
        axis=1
    )

    # Gap %
    df_analysis["Gap 1 to 2 (%)"] = (
        (df_analysis["2nd Lowest"] - df_analysis["1st Lowest"]) / df_analysis["1st Lowest"] * 100
    ).round(2)

    # Median price
    df_analysis["Median Price"] = df_analysis[vendor_cols].median(axis=1)

    # Vendor → Median (%)
    for v in vendor_cols:
        df_analysis[f"{v} to Median (%)"] = (
            (df_analysis[v] - df_analysis["Median Price"]) / df_analysis["Median Price"] * 100
        ).round(2)

    # Simpan ke session state
    st.session_state["bid_and_price_analysis_tco_by_round"] = df_analysis

    # --- 🎯 Tambahkan slicer
    all_round = sorted(df_analysis["ROUND"].dropna().unique())
    all_1st = sorted(df_analysis["1st Vendor"].dropna().unique())
    all_2nd = sorted(df_analysis["2nd Vendor"].dropna().unique())

    col_sel_1, col_sel_2, col_sel_3 = st.columns(3)
    with col_sel_1:
        selected_round = st.multiselect(
            "Filter: Round",
            options=all_round,
            default=[],
            placeholder="Choose rounds",
        )
    with col_sel_2:
        selected_1st = st.multiselect(
            "Filter: 1st vendor",
            options=all_1st,
            default=[],
            placeholder="Choose vendors",
        )
    with col_sel_3:
        selected_2nd = st.multiselect(
            "Filter: 2nd vendor",
            options=all_2nd,
            default=[],
            placeholder="Choose vendors",
        )

    # --- Terapkan filter AND secara dinamis
    df_filtered_analysis = df_analysis.copy()

    if selected_round:
        df_filtered_analysis = df_filtered_analysis[df_filtered_analysis["ROUND"].isin(selected_round)]

    if selected_1st:
        df_filtered_analysis = df_filtered_analysis[df_filtered_analysis["1st Vendor"].isin(selected_1st)]

    if selected_2nd:
        df_filtered_analysis = df_filtered_analysis[df_filtered_analysis["2nd Vendor"].isin(selected_2nd)]

    # Format
    num_cols = df_filtered_analysis.select_dtypes(include=["number"]).columns
    format_dict = {col: format_rupiah for col in num_cols}
    format_dict.update({"Gap 1 to 2 (%)": "{:.1f}%"})
    for vendor in vendor_cols:
        format_dict[f"{vendor} to Median (%)"] = "{:+.1f}%"

    df_analysis_styled = (
        df_filtered_analysis.style
        .format(format_dict)
        .apply(lambda row: highlight_1st_2nd_vendor(row, df_filtered_analysis.columns), axis=1)
    )

    st.caption(f"✨ Total number of data entries: **{len(df_filtered_analysis)}**")
    st.dataframe(df_analysis_styled, hide_index=True)

    # Download button to Excel
    excel_data = get_excel_download_highlight_1st_2nd_lowest(df_filtered_analysis)

    # Layout tombol (rata kanan)
    col1, col2, col3 = st.columns([2.3,2,1])
    with col3:
        st.download_button(
            label="Download",
            data=excel_data,
            file_name="Bid & Price Analysis - TCO by Round.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            icon=":material/download:"
        )
    
    st.divider()

    # PRICE MOVEMENTT ANALYSIS
    st.markdown("##### 💸 Price Movement Analysis")

    # Identifikasi kolom
    non_round_cols = [c for c in df_no_total.columns if c != "ROUND"]
    scope_cols = df_no_total[non_round_cols].select_dtypes(exclude=["number"]).columns.tolist()
    vendor_cols = df_no_total[non_round_cols].select_dtypes(include=['number']).columns.tolist()

    # Melt (Unpivot)
    df_long = df_no_total.melt(
        id_vars=["ROUND"] + scope_cols,
        value_vars=vendor_cols,
        var_name="VENDOR",
        value_name="PRICE"
    )

    # PIVOT
    df_pivot = df_long.pivot_table(
        index=["VENDOR"] + scope_cols,
        columns="ROUND",
        values="PRICE",
        aggfunc="first"
    ).reset_index()

    # Urutkan kolom ROUND
    round_order = sorted([c for c in df_pivot.columns if c not in ["VENDOR"] + scope_cols])
    df_pivot = df_pivot[["VENDOR"] + scope_cols + round_order]

    # PRICE REDUCTION
    def compute_price_reduction(row):
        prices = row[round_order]
        valid_prices = prices.dropna()

        if len(valid_prices) < 2:
            return pd.Series({
                "PRICE REDUCTION (VALUE)": np.nan, 
                "PRICE REDUCTION (%)": np.nan
            })

        first_val = valid_prices.iloc[0]
        last_val  = valid_prices.iloc[-1]

        reduction_value = last_val - first_val
        reduction_pct   = (reduction_value / first_val) * 100

        return pd.Series({
            "PRICE REDUCTION (VALUE)": reduction_value,
            "PRICE REDUCTION (%)": round(reduction_pct, 2)
        }) 
    
    df_pivot[["PRICE REDUCTION (VALUE)", "PRICE REDUCTION (%)"]] = (
        df_pivot.apply(compute_price_reduction, axis=1)
    )

    # PRICE TREND
    def detect_trend(row):
        prices = row[round_order].values.astype(float)
        prices = prices[~np.isnan(prices)]

        if len(prices) <= 1:
            return "Insufficient Data"
        if all(prices[i] > prices[i+1] for i in range(len(prices)-1)):
            return "Consistently Down"
        if all(prices[i] < prices[i+1] for i in range(len(prices)-1)):
            return "Consistently Up"
        if len(set(prices)) == 1:
            return "No Change"
        
        return "Fluctuating"
    
    df_pivot["PRICE TREND"] = df_pivot.apply(detect_trend, axis=1)

    # PRICE STABILITY INDEX (PSI)
    def compute_psi(row):
        prices = row[round_order].astype(float)
        prices = prices[~np.isnan(prices)]
        if len(prices) == 0:
            return np.nan
        return ((prices.max() - prices.min()) / prices.mean()) * 100
    
    df_pivot["STANDARD DEVIATION"] = df_pivot[round_order].std(axis=1, ddof=0).round(4)
    df_pivot["PRICE STABILITY INDEX (%)"] = df_pivot.apply(compute_psi, axis=1).round(2)

    # Adding "TOTAL" columns
    round_cols = round_order.copy()   # ambil kolom ROUND saja untuk dihitung

    total_rows = []
    for vendor in df_pivot["VENDOR"].unique():
        df_vendor = df_pivot[df_pivot["VENDOR"] == vendor]

        # Hitung sum per ROUND
        total_data = df_vendor[round_cols].sum(numeric_only=True)

        # Buat row kosong
        # total_row = {col: "" for col in df_pivot.columns}
        total_row = {col: np.nan for col in df_pivot.columns}

        total_row["VENDOR"] = vendor
        total_row[scope_cols[0]] = "TOTAL"

        # Masukkan total ROUND
        for col in round_cols:
            total_row[col] = total_data[col]

        total_rows.append(total_row)

    df_total_rows = pd.DataFrame(total_rows)
    df_pivot = pd.concat([df_pivot, df_total_rows], ignore_index=True)

    # Urutkan lagi: vendor tetap grouping
    df_pivot = df_pivot.sort_values(["VENDOR", scope_cols[0]], key=lambda s: s.replace("TOTAL", "ZZZ"))

    # Tambahkan slicer 
    all_vendor = sorted(df_pivot["VENDOR"].dropna().unique())
    all_trend  = sorted(df_pivot["PRICE TREND"].dropna().unique())

    col_sel_1, col_sel_2 = st.columns(2)
    with col_sel_1:
        selected_vendor = st.multiselect(
            "Filter: Vendor",
            options=all_vendor,
            default=None,
            placeholder="Choose one or more vendors",
            key="filter_vendor"
        )
    with col_sel_2:
        selected_trend = st.multiselect(
            "Filter: Price Trend",
            options=all_trend,
            default=None,
            placeholder="Choose one or more price trends",
            key="filter_price_trend"
        )

    # Terapkan filter dengan logika AND
    if selected_vendor and selected_trend:
        df_filter_pivot = df_pivot[
            df_pivot["VENDOR"].isin(selected_vendor) &
            df_pivot["PRICE TREND"].isin(selected_trend)
        ]
    elif selected_vendor:
        df_filter_pivot = df_pivot[df_pivot["VENDOR"].isin(selected_vendor)]
    elif selected_trend:
        df_filter_pivot = df_pivot[df_pivot["PRICE TREND"].isin(selected_trend)]
    else:
        df_filter_pivot = df_pivot.copy()

    # Format
    num_cols = df_filter_pivot.select_dtypes(include=["number"]).columns
    format_dict = {col: format_rupiah for col in num_cols}
    format_dict.update({
        "PRICE REDUCTION (%)": "{:+.1f}%",
        "PRICE STABILITY INDEX (%)": "{:.1f}%"
    })

    df_pivot_style = (
        df_filter_pivot.style
        .format(format_dict)
        .apply(highlight_total_row_v2, axis=1)
    )

    # Tampilkan
    st.caption(f"✨ Total number of data entries: **{len(df_filter_pivot)}**")
    st.dataframe(df_pivot_style, hide_index=True)

    # --- Prepare dataframe for Excel export ---
    df_export = df_filter_pivot.copy()

    # Replace NaN / Inf with empty string to avoid xlsxwriter error
    df_export = df_export.replace([np.nan, np.inf, -np.inf], "")

    # Download
    excel_data = get_excel_download_highlight_total(df_export)
    # excel_data = get_excel_download(df_filter_pivot)
    col1, col2, col3 = st.columns([2.3,2,1])
    with col3:
        st.download_button(
            label="Download",
            data=excel_data,
            file_name="Price Movement Trend - TCO by Round.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            icon=":material/download:"
        )

    # VISUALIZATIONN
    st.markdown("##### 📊 Visualization")
    tab1, tab2 = st.tabs(["Winning Performance", "Price Trend"])

    # WINNING PERFORMANCEE
    # Gabungkan semua round
    df_all = df_analysis.copy()

    # --- Normalisasi nama vendor (biar konsisten) ---
    df_all["1st Vendor"] = (
        df_all["1st Vendor"]
        .astype(str)
        .str.strip()      # hilangkan spasi depan/belakang
        .str.upper()      # ubah ke huruf besar semua
    )

    # --- Hitung jumlah kemenangan per vendor per round ---
    win_summary = (
        df_all.groupby(["ROUND", "1st Vendor"])
            .size()
            .reset_index(name="Wins")
            .rename(columns={"1st Vendor": "VENDOR"})
    )

    # --- Urutkan round ---
    round_order = sorted(df_all["ROUND"].unique(), key=lambda x: str(x))
    win_summary["ROUND"] = pd.Categorical(win_summary["ROUND"], categories=round_order, ordered=True)

    # --- Tambahan: pastikan kombinasi Round–Vendor yang hilang diisi 0 ---
    all_rounds = win_summary["ROUND"].unique()
    all_vendors = win_summary["VENDOR"].unique()

    # Buat semua kombinasi round–vendor
    full_index = pd.MultiIndex.from_product([all_rounds, all_vendors], names=["ROUND", "VENDOR"])
    win_summary = (
        win_summary.set_index(["ROUND", "VENDOR"])
        .reindex(full_index, fill_value=0)
        .reset_index()
    )

    # --- Buat chart dengan Altair ---
    y_min = win_summary["Wins"].min()
    y_max = win_summary["Wins"].max()

    # --- Hitung total kemenangan per vendor ---
    vendor_order = (
        win_summary.groupby("VENDOR")["Wins"]
        .sum()
        .sort_values(ascending=False)
        .index.tolist()
    )

    chart = (
        alt.Chart(win_summary)
        .mark_line(point=True)
        .encode(
            x=alt.X("ROUND:N", sort=round_order, title="Round"),
            y=alt.Y(
                "Wins:Q", 
                title="Number of Wins",
                scale=alt.Scale(domain=[y_min - 1.5, y_max + 1.5]),
                axis=alt.Axis(
                    tickMinStep=1,
                    tickCount=win_summary["Wins"].nunique() + 1
                )
            ),
            color=alt.Color("VENDOR:N", title="Vendor", sort=vendor_order)
        )
        .properties(
            height=400,
            width="container",
            title="Winning Performance Across Rounds"
        ).configure_title(
            anchor='middle', 
            offset=12,
            fontSize=14
        )
        .configure_view(stroke='gray', strokeWidth=1)
        .configure_point(size=60)
        .configure_axis(labelFontSize=12, titleFontSize=13)
        .configure_legend(
            titleFontSize=12,        
            titleFontWeight="bold",  
            labelFontSize=12,        
            labelLimit=300,   
            orient="right"
        )
    )

    # Table
    win_table = (
        win_summary
        .pivot_table(
            index="VENDOR",
            columns="ROUND",
            values="Wins",
            aggfunc="sum",
            fill_value=0,
            observed=False
        ).reset_index()
    )

    # Urutkan vendor berdasarkan total kemenangan
    win_table["Total Wins"] = win_table.drop(columns="VENDOR").sum(axis=1)
    win_table = win_table.sort_values("Total Wins", ascending=False).reset_index(drop=True)

    # st.dataframe(win_table, hide_index=True)
    tab1.write("")
    tab1.altair_chart(chart)

    with tab1:
        with st.expander("See explanation"):
            st.write('''
                The visualization above shows the number of wins each vendor
                achieves in every tender round. A win is counted based on which
                vendor becomes the best bidder **(1st Vendor)** for each scope.
                     
                **💡 How to interpret the chart**
                     
                - High Wins Value  
                     Vendor is highly competitive in that round and wins more scopes
                     than others.  
                - Increasing Wins Across Rounds  
                     Indicates improving perfomance or more competitive pricing in later 
                     rounds.  
                - Decreasing Wins Across Rounds  
                     Shows declining competitiveness, with the vendor losing more scopes
                     compared the previous rounds.  
                - Zero Wins in a Round  
                     Vendor did not win any scope in that round, indicating weak competitiveness
                     for that stage.
            ''')
            st.dataframe(win_table, hide_index=True)

            # Simpan hasil ke variabel
            excel_data = get_excel_download(win_table)

            # Layout tombol (rata kanan)
            col1, col2, col3 = st.columns([3,1,1])
            with col3:
                st.download_button(
                    label="Download",
                    data=excel_data,
                    file_name="Win Rate Trend Summary.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    icon=":material/download:"
                )

    # PRICEE TRENDD
    trend_order = ["No Change", "Consistently Down", "Consistently Up", "Fluctuating", "Insufficient Data"]

    vendor_col = df_pivot.columns[0]

    # Ringkas jumlah kemunculan per vendor per trend
    trend_summary = (
        df_pivot.groupby([vendor_col, "PRICE TREND"], observed=False)
                .size()
                .reset_index(name="Count")
    )

    trend_summary["PRICE TREND"] = pd.Categorical(
        trend_summary["PRICE TREND"],
        categories=trend_order,
        ordered=True
    )

    # Urutan vendor
    vendor_order = (
        trend_summary.groupby(vendor_col)["Count"]
        .sum()
        .sort_values(ascending=False)
        .index.tolist()
    )

    # Warna vendor konsisten (optional)
    vendor_colors = [
        "#32AAB5", "#E7D39A", "#F1AA60", "#F27B68",
        "#E54787", "#BF219A", "#8E0F9C", "#4B1D91",
        "#246BCE", "#1EC8A5", "#F8C537"
    ]

    # Mapping warna
    color_map = {
        v: vendor_colors[i % len(vendor_colors)]
        for i, v in enumerate(vendor_order)
    }

    # Format angka ribuan pada sumbu Y
    y_axis = alt.Axis(
        title="Number of Occurrences",
        # grid=False,
        labelPadding=12,
        format="d"   # angka bulat, bukan 1k/1m
    )

    trend_summary["VendorIndex"] = trend_summary[vendor_col].map(
        {v: i for i, v in enumerate(vendor_order)}
    )
    trend_summary["TotalVendors"] = trend_summary.groupby("PRICE TREND", observed=False)[vendor_col].transform("nunique")

    # posisi label = vendor index + 0.5 (tengah bar)
    trend_summary["LabelOffset"] = trend_summary["VendorIndex"] + 0.5

    # Chart
    bars = (
        alt.Chart(trend_summary)
            .mark_bar()
            .encode(
                x=alt.X(
                    "PRICE TREND:N",
                    sort=trend_order,
                    axis=alt.Axis(labelAngle=0),
                    title=None
                ),
                y=alt.Y(
                    "Count:Q",
                    axis=y_axis,
                    title="Number of Occurrences",
                    scale=alt.Scale(domain=[0, trend_summary["Count"].max() * 1.15])
                ),
                color=alt.Color(
                    f"{vendor_col}:N",
                    title="Vendor",
                    sort=vendor_order
                ),
                xOffset=f"{vendor_col}:N"
            )
    )

    labels = (
        alt.Chart(trend_summary)
            .mark_text(
                dy=-7,                # geser sedikit ke atas
                fontSize=10,
                fontWeight="bold",
                color="gray",
            )
            .encode(
                x=alt.X(
                    "PRICE TREND:N",
                    sort=trend_order
                ),
                y=alt.Y("Count:Q"),
                text="Count:Q",
                xOffset=f"{vendor_col}:N"   # penting: supaya posisinya sama dengan bar!
            )
    )

    trend_chart = (
        (bars + labels)
            .properties(
                height=400,
                padding={"right": 15},
                width="container",
                title="Price Trend Distribution per Vendor"
            )
            .configure_title(anchor="middle", offset=12, fontSize=14)
            .configure_view(stroke="gray", strokeWidth=1)
            .configure_axis(labelFontSize=12, titleFontSize=13)
            .configure_legend(
                titleFontSize=12,
                titleFontWeight="bold",
                labelFontSize=12,
                labelLimit=300,
                orient="bottom"
            )
    )

    tab2.write("")
    tab2.altair_chart(trend_chart)

    trend_table = (
        trend_summary
            .pivot_table(
                index="PRICE TREND",
                columns=vendor_col,
                values="Count",
                aggfunc="sum",
                fill_value=0,
                observed=False
            )
            .reindex(trend_order)   # pastikan urut
            .reset_index()
    )

    with tab2:
        with st.expander("See explanation"):
            st.write('''
                The chart above shows the number of occurrences of each **Price 
                Trend** for every vendor based on the pivoted tender data.
                     
                **💡 How to interpret the chart**
                     
                - No Change  
                     The vendor's price remains stable across all rounds or periods.
                - Consistently Down  
                     The vendor's price decreases continuously from one round to the next.
                - Consistently Up  
                     The vendor's price increases in every subsequent round.
                - Fluctuating  
                     The vendor's price moves up and down across the rounds.
            ''')
            st.dataframe(trend_table, hide_index=True)

            # Simpan hasil ke variabel
            excel_data = get_excel_download(trend_table)

            # Layout tombol (rata kanan)
            col1, col2, col3 = st.columns([3,1,1])
            with col3:
                st.download_button(
                    label="Download",
                    data=excel_data,
                    file_name="Price Trend Summary.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    icon=":material/download:",
                )
