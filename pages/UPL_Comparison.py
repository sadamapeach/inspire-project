import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
import time
import re
from io import BytesIO
from functools import reduce
    
def round_half_up_num(series):
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
        format_rupiah_xls = workbook.add_format({'num_format': '#,##0'})
        fmt_pct_rupiah   = workbook.add_format({'num_format': '#,##0.0"%"'})
        highlight_format = workbook.add_format({
            "bold": True,
            "bg_color": "#D9EAD3",  # hijau lembut
            "font_color": "#1A5E20",  # hijau tua
            "num_format": "#,##0"
        })

        # Terapkan format
        for col_num, col_name in enumerate(df.columns):
            if col_name in df.select_dtypes(include=["number"]).columns:
                worksheet.set_column(col_num, col_num, 15, format_rupiah_xls)  

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

        # Tentukan format
        format_rupiah_xls = workbook.add_format({'num_format': '#,##0'})
        fmt_pct_rupiah   = workbook.add_format({'num_format': '#,##0.0"%"'})

        # Terapkan format
        for col_num, col_name in enumerate(df.columns):
            if col_name in df.select_dtypes(include=["number"]).columns:
                worksheet.set_column(col_num, col_num, 15, format_rupiah_xls)

            if "%" in col_name:
                worksheet.set_column(col_num, col_num, 15, fmt_pct_rupiah)

        # --- Format umum ---
        format_first = workbook.add_format({'bg_color': '#C6EFCE', "num_format": "#,##0"})  # hijau Excel-style
        format_second = workbook.add_format({'bg_color': '#FFEB9C', "num_format": "#,##0"}) # kuning Excel-style

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
    st.header("5️⃣ UPL Comparison")
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
        st.session_state["uploaded_file_upl_comparison"] = upload_file

        # --- Animasi proses upload ---
        msg = st.toast("📂 Uploading file...")
        time.sleep(1.2)

        msg.toast("🔍 Reading sheets...")
        time.sleep(1.2)

        msg.toast("✅ File uploaded successfully!")
        time.sleep(0.5)

        # Baca semua sheet sekaligus
        all_df = pd.read_excel(upload_file, sheet_name=None)
        st.session_state["all_df_upl_comparison_raw"] = all_df  # simpan versi mentah

    elif "all_df_upl_comparison_raw" in st.session_state:
        all_df = st.session_state["all_df_upl_comparison_raw"]
    else:
        return
    
    st.divider()

    # OVERVIEW
    # st.subheader("🔍 Overview")
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

        # Pembulatan
        num_cols = df_clean.select_dtypes(include=["number"]).columns
        df_clean[num_cols] = df_clean[num_cols].apply(round_half_up_num)

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

    st.session_state["result_upl_comparison"] = result
    # st.divider()

    # MERGEE
    st.markdown("##### 🗃️ Merge Data")
    st.caption(f"Successfully consolidated data from **{total_sheets} vendors**.")

    merged_list = []
    for vendor, df_clean in result.items():
        df_temp = df_clean.copy()

        # Deteksi kolom numeric terakhir (last index)
        numeric_cols = df_temp.select_dtypes(include=["number"]).columns
        last_numeric_col = numeric_cols[-1] if len(numeric_cols) > 0 else None

        # Buat dictionary untuk baris TOTAL
        total_row = {col: "" for col in df_temp.columns}
        total_row[df_temp.columns[0]] = "TOTAL"

        # Hitung total untuk kolom numerik (last index)
        if last_numeric_col:
            total_row[last_numeric_col] = df_temp[last_numeric_col].sum(skipna=True)

        # Tambahkan baris total di akhir DataFrame
        df_temp = pd.concat([df_temp, pd.DataFrame([total_row])], ignore_index=True)

        # Tambahkan kolom vendor di depan as first index
        df_temp.insert(0, "VENDOR", vendor)

        merged_list.append(df_temp)

    # Gabungkan semua vendor jadi satu DataFrame
    merged_overview = pd.concat(merged_list, ignore_index=True)

    # Pastikan kolom berurutan (vendor as index-0)
    cols = ["VENDOR"] + [c for c in merged_overview.columns if c != "VENDOR"]
    merged_overview = merged_overview[cols]

    # Simpan session
    st.session_state["merge_overview_upl_comparison"] = merged_overview

    # Format rupiah dan tampilkan
    num_cols = merged_overview.select_dtypes(include=["number"]).columns
    merged_overview_styled = (
        merged_overview.style
        .format({col: format_rupiah for col in num_cols})
        .apply(highlight_total_row_v2, axis=1)
    )
    st.dataframe(merged_overview_styled, hide_index=True)

    # Download
    excel_data = get_excel_download_highlight_total(merged_overview)
    col1, col2, col3 = st.columns([2.3,2,1])
    with col3:
        st.download_button(
            label="Download",
            data=excel_data,
            file_name="Merge UPL Comparison.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            icon=":material/download:",
        )

    st.divider()

    # MERGEE TRANSPOSEE
    st.markdown("##### 🛸 Transpose Data")
    st.caption("Cross-vendor price mapping to simplify analysis and highlight pricing differences.")

    df_transposed_list = []

    for vendor, df_clean in result.items():
        df_temp = df_clean.copy().reset_index(drop=True)

        # Deteksi kolom numerik terakhir (last index)
        numeric_cols = df_temp.select_dtypes(include=["number"]).columns
        last_numeric_col = numeric_cols[-1] if len(numeric_cols) > 0 else None 

        if last_numeric_col:
            # Ambil kolom deskriptif (semua kolom selain last index)
            non_num_cols = [c for c in df_temp.columns if c != last_numeric_col]

            # Tambahkan kolom order untuk menjaga urutan
            df_temp["__ORDER__"] = range(len(df_temp))

            # Buat DF per vendor: deskriptif + harga
            df_temp = df_temp[non_num_cols + ["__ORDER__", last_numeric_col]].copy()
            df_temp.rename(columns={last_numeric_col: vendor}, inplace=True)

            df_transposed_list.append(df_temp)

    # Gabungkan semua vendor berdasarkan kolom deskriptif
    df_transposed_overview = reduce(
        lambda left, right: pd.merge(left, right, on=non_num_cols + ["__ORDER__"], how="outer"),
        df_transposed_list
    )

    # Urutkan berdasarkan urutan asli
    df_transposed_overview = df_transposed_overview.sort_values("__ORDER__").reset_index(drop=True)

    # Hapus kolom __ORDER__ (tidak perlu ditampilkan)
    df_transposed_overview.drop(columns="__ORDER__", inplace=True)

    # Tambahkan baris TOTAL (sum semua kolom vendor numerik)
    vendor_cols = [v for v in result.keys() if v in df_transposed_overview.columns]

    total_row = {col: "" for col in df_transposed_overview.columns}
    total_row[non_num_cols[0]] = "TOTAL"
    for v in vendor_cols:
        total_row[v] = df_transposed_overview[v].sum(skipna=True)

    df_transposed_overview = pd.concat(
        [df_transposed_overview, pd.DataFrame([total_row])],
        ignore_index=True 
    )

    # Format Rupiah
    df_transposed_styled = (
        df_transposed_overview.style
        .format({col: format_rupiah for col in vendor_cols})
        .apply(highlight_total_row_v2, axis=1)
    )

    st.dataframe(df_transposed_styled, hide_index=True)

    # Simpan ke session
    st.session_state["transposed_overview_upl_comparison"] = df_transposed_overview

    # Download
    excel_data = get_excel_download_highlight_total(df_transposed_overview)
    col1, col2, col3 = st.columns([2.3,2,1])
    with col3:
        st.download_button(
            label="Download",
            data=excel_data,
            file_name="Transpose UPL Comparison.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            icon=":material/download:",
        )

    st.divider()

    # BID & PRICE ANALYSIS
    st.markdown("##### 🧠 Bid & Price Analysis")

    df_analysis = df_transposed_overview.copy()

    # Buang baris TOTAL sebelum analisis
    df_analysis = df_analysis[df_analysis.iloc[:, 0] != "TOTAL"].reset_index(drop=True)

    # Deteksi kolom vendor (numerik)
    vendor_cols = [c for c in df_analysis.columns if c in result.keys()]

    # Pastikan kolom vendor bertipe numerik
    for v in vendor_cols:
        df_analysis[v] = pd.to_numeric(df_analysis[v], errors="coerce")

    # Hitung 1st dan 2nd lowest
    df_analysis["1st Lowest"] = df_analysis[vendor_cols].min(axis=1)
    df_analysis["1st Vendor"] = df_analysis[vendor_cols].idxmin(axis=1)

    # Untuk 2nd lowest, ganti nilai lowest jadi NaN dulu, lalu cari min lagi
    temp = df_analysis[vendor_cols].mask(df_analysis[vendor_cols].eq(df_analysis["1st Lowest"], axis=0))
    df_analysis["2nd Lowest"] = temp.min(axis=1)
    df_analysis["2nd Vendor"] = temp.idxmin(axis=1)

    # Hitung gap antara 1st dan 2nd lowest (%)
    df_analysis["Gap 1 to 2 (%)"] = ((df_analysis["2nd Lowest"] - df_analysis["1st Lowest"]) / df_analysis["1st Lowest"] * 100).round(2)

    # Hitung median price
    df_analysis["Median Price"] = df_analysis[vendor_cols].median(axis=1)

    # Hitung selisih tiap vendor dengan median (%)
    for v in vendor_cols:
        df_analysis[f"{v} to Median (%)"] = ((df_analysis[v] - df_analysis["Median Price"]) / df_analysis["Median Price"] * 100).round(2)

    # --- 🎯 Tambahkan dua slicer terpisah untuk 1st Vendor dan 2nd Vendor
    all_1st = sorted(df_analysis["1st Vendor"].dropna().unique())
    all_2nd = sorted(df_analysis["2nd Vendor"].dropna().unique())

    col_sel_1, col_sel_2 = st.columns(2)
    with col_sel_1:
        selected_1st = st.multiselect(
            "Filter: 1st vendor",
            options=all_1st,
            default=None,
            placeholder="Choose one or more vendors",
            key="filter_1st_vendor"
        )
    with col_sel_2:
        selected_2nd = st.multiselect(
            "Filter: 2nd vendor",
            options=all_2nd,
            default=None,
            placeholder="Choose one or more vendors",
            key="filter_2nd_vendor"
        )

    # --- Terapkan filter dengan logika AND
    if selected_1st and selected_2nd:
        df_filtered = df_analysis[
            df_analysis["1st Vendor"].isin(selected_1st) &
            df_analysis["2nd Vendor"].isin(selected_2nd)
        ]
    elif selected_1st:
        df_filtered = df_analysis[df_analysis["1st Vendor"].isin(selected_1st)]
    elif selected_2nd:
        df_filtered = df_analysis[df_analysis["2nd Vendor"].isin(selected_2nd)]
    else:
        df_filtered = df_analysis.copy()

    # --- Tambahkan styling ---
    num_cols = df_filtered.select_dtypes(include=["number"]).columns
    format_dict = {col: format_rupiah for col in num_cols}
    format_dict.update({"Gap 1 to 2 (%)": "{:.1f}%"})
    for v in vendor_cols:
        format_dict[f"{v} to Median (%)"] = "{:+.1f}%"

    df_analysis_styled = (
        df_filtered.style
        .format(format_dict)
        .apply(lambda row: highlight_1st_2nd_vendor(row, df_filtered.columns), axis=1)
    )

    st.caption(f"✨ Total number of data entries: **{len(df_filtered)}**")
    st.dataframe(df_analysis_styled, hide_index=True)

    # Simpan hasil ke variabel
    excel_data = get_excel_download_highlight_1st_2nd_lowest(df_filtered)

    # Layout tombol (rata kanan)
    col1, col2, col3 = st.columns([2.3,2,1])
    with col3:
        st.download_button(
            label="Download",
            data=excel_data,
            file_name="UPL Comparison Analysis.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            icon=":material/download:"
        )

    # VISUALIZATION
    st.markdown("##### 📊 Visualization")
    tab1, tab2 = st.tabs(["Win Rate Trend", "Average Gap Trend"])

    # --- WIN RATE VISUALIZATION ---
    df_win_rate = df_analysis.copy()
    # --- Hitung total kemenangan (1st & 2nd Vendor)
    win1_counts = df_win_rate["1st Vendor"].value_counts(dropna=True).reset_index()
    win1_counts.columns = ["Vendor", "Wins1"]

    win2_counts = df_win_rate["2nd Vendor"].value_counts(dropna=True).reset_index()
    win2_counts.columns = ["Vendor", "Wins2"]

    # --- Hitung total partisipasi vendor ---
    vendor_counts = (
        df_win_rate[vendor_cols]
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
    df_win_rate = win_rate.rename(columns={
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
            st.dataframe(df_win_rate, hide_index=True)

            # Simpan hasil ke variabel
            excel_data = get_excel_download(df_win_rate, sheet_name="Win Rate Trend Summary")

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
    df_avg_gap = df_analysis.copy()
    # --- Hitung rata-rata Gap 1 to 2 (%) per Vendor (hanya untuk 1st Vendor)
    df_gap = df_avg_gap.copy()

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

        
    #     # --- WIN RATE VISUALIZATION ---
    #     st.markdown("##### Win Rate Trend")
    #     st.caption("Visualizing 1st and 2nd place win rates across all vendors to assess competitiveness.")

    #     # --- Hitung total kemenangan (1st & 2nd Vendor)
    #     win1_counts = df_pivot["1st Vendor"].value_counts(dropna=True).reset_index()
    #     win1_counts.columns = ["Vendor", "Wins1"]

    #     win2_counts = df_pivot["2nd Vendor"].value_counts(dropna=True).reset_index()
    #     win2_counts.columns = ["Vendor", "Wins2"]

    #     # --- Hitung total partisipasi vendor ---
    #     vendor_counts = (
    #         df_pivot[vendor_cols]
    #         .notna()       # True kalau vendor berpartisipasi (ada harga)
    #         .sum()         # Hitung True per kolom
    #         .reset_index()
    #     )
    #     vendor_counts.columns = ["Vendor", "Total Participations"]

    #     # --- Gabungkan semua ---
    #     win_rate = (
    #         vendor_counts
    #         .merge(win1_counts, on="Vendor", how="left")
    #         .merge(win2_counts, on="Vendor", how="left")
    #         .fillna(0)
    #     )

    #     # --- Hitung Win Rate (%)
    #     win_rate["1st Win Rate (%)"] = np.where(
    #         win_rate["Total Participations"] > 0,
    #         (win_rate["Wins1"] / win_rate["Total Participations"] * 100).round(1),
    #         0
    #     )
    #     win_rate["2nd Win Rate (%)"] = np.where(
    #         win_rate["Total Participations"] > 0,
    #         (win_rate["Wins2"] / win_rate["Total Participations"] * 100).round(1),
    #         0
    #     )

    #     # --- Siapkan data long-format untuk visualisasi ---
    #     win_rate_long = win_rate.melt(
    #         id_vars=["Vendor"],
    #         value_vars=["1st Win Rate (%)", "2nd Win Rate (%)"],
    #         var_name="Metric",
    #         value_name="Percentage"
    #     )

    #     # --- Urutkan vendor berdasarkan 1st Win Rate tertinggi ---
    #     vendor_order = (
    #         win_rate.sort_values("1st Win Rate (%)", ascending=False)["Vendor"].tolist()
    #     )

    #     # --- Tentukan warna untuk kedua metrik ---
    #     metric_colors = {
    #         "1st Win Rate (%)": "#1f77b4",
    #         "2nd Win Rate (%)": "#ff7f0e"     # biru
    #     }

    #     # --- Highlight interaktif ---
    #     highlight = alt.selection_point(on='mouseover', fields=['Metric'], nearest=True)

    #     # --- Batas atas dan bawah sumbu Y ---
    #     y_max = win_rate_long["Percentage"].max()
    #     if not np.isfinite(y_max) or y_max <= 0:
    #         y_max = 1

    #     # Pastikan data diurutkan sesuai vendor_order
    #     win_rate_long["Vendor"] = pd.Categorical(win_rate_long["Vendor"], categories=vendor_order, ordered=True)
    #     win_rate_long = win_rate_long.sort_values(["Metric", "Vendor"])

    #     # --- Chart utama ---
    #     base = (
    #         alt.Chart(win_rate_long)
    #         .encode(
    #             x=alt.X("Vendor:N", sort=vendor_order, title=None),
    #             y=alt.Y(
    #                 "Percentage:Q",
    #                 title="Win Rate (%)",
    #                 scale=alt.Scale(domain=[0, y_max * 1.2])
    #             ),
    #             color=alt.Color(
    #                 "Metric:N",
    #                 scale=alt.Scale(
    #                     domain=list(metric_colors.keys()),
    #                     range=list(metric_colors.values())
    #                 ),
    #                 title="Rank"
    #             ),
    #             tooltip=[
    #                 alt.Tooltip("Vendor:N", title="Vendor"),
    #                 alt.Tooltip("Metric:N", title="Position"),
    #                 alt.Tooltip("Percentage:Q", title="Win Rate (%)", format=".1f")
    #             ]
    #         )
    #     )

    #     # --- Garis dengan titik ---
    #     lines = base.mark_line(point=alt.OverlayMarkDef(size=70, filled=True), strokeWidth=3)

    #     # --- Label persentase di atas titik ---
    #     labels = base.mark_text(
    #         align='center',
    #         baseline='bottom',
    #         dy=-7,
    #         fontWeight='bold',
    #         color='gray'
    #     ).encode(
    #         text=alt.Text("Percentage:Q", format=".1f")
    #     ).transform_calculate(
    #         label="format(datum.Percentage, '.1f') + '%'"
    #     ).encode(
    #         text="label:N"
    #     )

    #     # --- Gabungkan semua elemen ---
    #     chart = (
    #         lines + labels
    #     ).properties(
    #         height=400,
    #         padding={"right": 15},
    #         title="📈 Vendor Win Rate Comparison (1st vs 2nd Place)"
    #     ).configure_title(
    #         anchor="middle",
    #         offset=12
    #     ).configure_axis(
    #         labelFontSize=12,
    #         titleFontSize=13
    #     ).configure_view(
    #         stroke='gray',
    #         strokeWidth=1
    #     ).configure_legend(
    #         titleFontSize=12,
    #         titleFontWeight="bold",
    #         labelFontSize=12,
    #         labelLimit=300,
    #         orient="bottom"
    #     )

    #     # --- Tampilkan chart di Streamlit
    #     st.altair_chart(chart)

    #     # Kolom yang mau ditaruh di depan
    #     cols_front = ["Wins1", "Wins2"]

    #     # Sisanya (Vendor + kolom lain yang tidak ada di cols_front)
    #     cols_rest = [c for c in win_rate.columns if c not in cols_front]

    #     # Gabungkan urutannya
    #     win_rate = win_rate[cols_rest[:1] + cols_front + cols_rest[1:]]

    #     # --- Ganti nama kolom biar lebih konsisten & enak dibaca
    #     df_summary = win_rate.rename(columns={
    #         "Wins1": "1st Rank",
    #         "Wins2": "2nd Rank"
    #     })

    #     with st.expander("See explanation"):
    #         st.write('''
    #             The visualization above compares the win rate of each vendor
    #             based on how often they achieved 1st or 2nd place in all
    #             tender evaluations.  
                     
    #             **💡 How to interpret the chart**  
                     
    #             - High 1st Win Rate (%)  
    #               Vendor is highly competitive and often offers the best commercial terms.  
    #             - High 2nd Win Rate (%)  
    #               Vendor consistently performs well, often just slightly less competitive than the winner.  
    #             - Large Gap Between 1st & 2nd Win Rate  
    #               Shows clear market dominance by certain vendors.
    #         ''')
    #         st.dataframe(df_summary, hide_index=True)
            
    #         # Download button to Excel
    #         @st.cache_data
    #         def get_excel_download(df_summary, sheet_name="Win Rate Trend Summary"):
    #             output = BytesIO()
    #             with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
    #                 df_summary.to_excel(writer, index=False, sheet_name=sheet_name)
    #             return output.getvalue()

    #         # Simpan hasil ke variabel
    #         excel_data = get_excel_download(df_summary, sheet_name="Win Rate Trend Summary")

    #         # Layout tombol (rata kanan)
    #         col1, col2, col3 = st.columns([3,1,1])
    #         with col3:
    #             st.download_button(
    #                 label="Download",
    #                 data=excel_data,
    #                 file_name="Win Rate Trend Summary.xlsx",
    #                 mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    #                 icon=":material/download:",
    #                 key=f"download_Win_Rate_Trend"
    #             )

    #     st.write("")

    #     # --- AVERAGE GAP VISUALIZATION ---
    #     st.markdown("##### Average Gap Trend")
    #     st.caption("Visualizing trend of the average gap between 1st and 2nd lowest bids.")

    #     # --- Hitung rata-rata Gap 1 to 2 (%) per Vendor (hanya untuk 1st Vendor)
    #     df_gap = df_pivot.copy()

    #     # Ubah kolom 'Gap 1 to 2 (%)' ke numerik (hapus simbol %)
    #     df_gap["Gap 1 to 2 (%)"] = (
    #         df_gap["Gap 1 to 2 (%)"]
    #         .replace("%", "", regex=True)
    #         .astype(float)
    #     )

    #     # Hitung rata-rata gap per vendor (hanya vendor yang jadi 1st Lowest)
    #     avg_gap = (
    #         df_gap.groupby("1st Vendor", dropna=True)["Gap 1 to 2 (%)"]
    #         .mean()
    #         .reset_index()
    #         .rename(columns={"1st Vendor": "Vendor", "Gap 1 to 2 (%)": "Average Gap (%)"})
    #         .sort_values("Average Gap (%)", ascending=False)
    #     )

    #     # st.dataframe(avg_gap)

    #     # Warna per vendor (biar konsisten kalau kamu sudah punya color mapping)
    #     colors_list = ["#F94144", "#F3722C", "#F8961E", "#F9C74F", "#90BE6D",
    #                 "#43AA8B", "#577590", "#E54787", "#BF219A", "#8E0F9C", "#4B1D91"]
    #     vendor_colors = {v: c for v, c in zip(avg_gap["Vendor"].unique(), colors_list)}

    #     # Interaksi hover
    #     highlight = alt.selection_point(on='mouseover', fields=['Vendor'], nearest=True)

    #     # --- Chart utama ---
    #     bars = (
    #         alt.Chart(avg_gap)
    #         .mark_bar()
    #         .encode(
    #             x=alt.X("Vendor:N", sort='-y', title=None),
    #             y=alt.Y("Average Gap (%):Q", title="Average Gap (%)", scale=alt.Scale(domain=[0, avg_gap["Average Gap (%)"].max() * 1.2])),
    #             color=alt.Color("Vendor:N",
    #                             scale=alt.Scale(domain=list(vendor_colors.keys()), range=list(vendor_colors.values())),
    #                             legend=None),
    #             tooltip=[
    #                 alt.Tooltip("Vendor:N", title="Vendor"),
    #                 alt.Tooltip("Average Gap (%):Q", title="Average Gap (%)", format=".1f")
    #             ]
    #         )
    #         .add_params(highlight)
    #     )

    #     # Label teks di atas bar
    #     labels = (
    #         alt.Chart(avg_gap)
    #         .mark_text(dy=-7, fontWeight='bold', color='gray')
    #         .encode(
    #             x="Vendor:N",
    #             y="Average Gap (%):Q",
    #             text=alt.Text("Average Gap (%):Q", format=".1f")  # Format angka
    #         )
    #         .transform_calculate(  # Tambahkan simbol %
    #             label_text="format(datum['Average Gap (%)'], '.1f') + '%'"
    #         )
    #         .encode(
    #             text="label_text:N"
    #         )
    #     )

    #     # Frame luar untuk gaya rapi
    #     frame = alt.Chart().mark_rect(stroke='gray', strokeWidth=1, fillOpacity=0)

    #     avg_line = alt.Chart(avg_gap).mark_rule(color='gray', strokeDash=[4,2], size=1.75).encode(
    #         y='mean(Average Gap (%)):Q'
    #     )

    #     # Gabungkan semua elemen
    #     chart = (bars + labels + frame + avg_line).properties(
    #         title="Average Gap (%) per 1st Vendor",
    #         height=400
    #     ).configure_title(
    #         anchor='middle',
    #         offset=12
    #     ).configure_axis(
    #         labelFontSize=12,
    #         titleFontSize=13
    #     ).configure_view(
    #         stroke='gray',
    #         strokeWidth=1
    #     )

    #     # --- Tampilkan di Streamlit ---
    #     st.altair_chart(chart)

    #     avg_value = avg_gap["Average Gap (%)"].mean()
    #     with st.expander("See explanation"):
    #         st.write(f'''
    #             The chart above shows the average price difference between 
    #             the lowest and second-lowest bids for each vendor when they 
    #             rank 1st, indicating their pricing dominance or competitiveness.
                     
    #             **💡 How to interpret the chart**  
                     
    #             - High Gap  
    #               High gap indicates strong vendor dominance (much lower prices).  
    #             - Low Gap  
    #               Low gap indicates intense competition with similar pricing among vendors.  
                
    #             The dashed line represents the average gap across all vendors, serving as a benchmark ({avg_value:.1f}%).
    #         ''')

    #     st.divider()

    #     # ---------- MEDIAN ---------- 
    #     st.subheader("🎯 Median Price & Vendor Deviation")
    #     st.caption("Shows each vendor’s percentage deviation from the median price — lower means more competitive.")

    #     # --- Hitung median per baris (berdasarkan vendor columns) ---
    #     df_median = df_pivot.copy()

    #     # Pastikan kolom vendor numeric dulu sebelum hitung median
    #     for v in vendor_cols:
    #         df_median[v] = pd.to_numeric(df_median[v].replace("", np.nan).astype(str).str.replace(".", ""), errors="coerce")

    #     # Hitung median tiap baris (berdasarkan vendor)
    #     df_median["Median"] = df_median[vendor_cols].median(axis=1, skipna=True)

    #     # Hitung gap tiap vendor terhadap median (%)
    #     for v in vendor_cols:
    #         df_median[f"{v} to Median (%)"] = df_median.apply(
    #             lambda row: (
    #                 f"{((row[v] - row['Median']) / row['Median'] * 100):.1f}%"
    #                 if pd.notna(row[v]) and pd.notna(row['Median']) and row['Median'] != 0
    #                 else ""
    #             ),
    #             axis=1
    #         )

    #     # Ambil hanya kolom yang diperlukan
    #     cols_final = list(info_cols) + ["Median"] + [f"{v} to Median (%)" for v in vendor_cols]
    #     df_median = df_median[cols_final]

    #     df_export_median = df_median.copy()

    #     # Optional: format Median jadi ribuan
    #     df_median["Median"] = df_median["Median"].apply(
    #         lambda x: f"{int(round(x)):,}".replace(",", ".") if pd.notna(x) else ""
    #     )

    #     # --- Fungsi untuk meng-highlight vendor dengan gap (%) terendah ---
    #     def highlight_lowest_median(s):
    #         # ambil hanya kolom vendor yang berisi persentase ke median
    #         vendor_cols_pct = [c for c in s.index if c.endswith("to Median (%)")]

    #         # konversi ke float (abaikan % dan kosong)
    #         vals = s[vendor_cols_pct].replace("", np.nan).str.replace("%", "").astype(float)

    #         # cari nilai minimum
    #         if vals.notna().any():
    #             min_val = vals.min()
    #         else:
    #             min_val = None

    #         # siapkan style per kolom
    #         styles = []
    #         for c in s.index:
    #             if c in vendor_cols_pct and pd.notna(min_val) and float(str(s[c]).replace("%", "")) == min_val:
    #                 styles.append("background-color: #f8c8dc; color: #7a1f47;")
    #             else:
    #                 styles.append("")
    #         return styles

    #     # --- Terapkan styling ---
    #     df_median_styled = df_median.style.apply(highlight_lowest_median, axis=1)

    #     st.dataframe(df_median_styled, hide_index=True)

    #     # Download button to Excel
    #     @st.cache_data
    #     def get_excel_download(df_export_median, sheet_name="Median Analysis (%)"):
    #         output = BytesIO()
    #         with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
    #             df_export_median.to_excel(writer, index=False, sheet_name=sheet_name)
    #         return output.getvalue()

    #     # Simpan hasil ke variabel
    #     excel_data = get_excel_download(df_export_median, sheet_name=f"Median Analysis (%)")

    #     # Layout tombol (rata kanan)
    #     col1, col2, col3 = st.columns([3,1,1])
    #     with col3:
    #         st.download_button(
    #             label="Download",
    #             data=excel_data,
    #             file_name=f"Median Analysis (%).xlsx",
    #             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    #             icon=":material/download:",
    #             key=f"download_median"
    #         )
