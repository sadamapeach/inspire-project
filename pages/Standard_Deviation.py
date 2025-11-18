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

def format_rupiah_percent(x):
    if pd.isna(x):
        return ""                   # hilangkan None / NaN
    return f"{format_rupiah(x)}%"   # pakai format_rupiah + %

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
    st.header("6️⃣ Standard Deviation")
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
    # st.subheader("📂 Upload File")
    st.markdown("##### 📂 Upload File")
    upload_file = st.file_uploader("Upload your file here!", type=["xlsx", "xls"])

    if upload_file is not None:
        st.session_state["uploaded_file_standard_deviation"] = upload_file

        # --- Animasi proses upload ---
        msg = st.toast("📂 Uploading file...")
        time.sleep(1.2)

        msg.toast("🔍 Reading sheets...")
        time.sleep(1.2)

        msg.toast("✅ File uploaded successfully!")
        time.sleep(0.5)

        # Baca sheet
        df = pd.read_excel(upload_file)

        # Simpan versi mentah (setelah konversi)
        st.session_state["df_standard_deviation_raw"] = df

    elif "df_standard_deviation_raw" in st.session_state:
        df = st.session_state["df_standard_deviation_raw"]
    else:
        return
    
    st.divider()

    # OVERVIEW
    st.markdown("##### 🔍 Overview")

    rows_before, cols_before = df.shape
    # Data cleaning
    df_clean = df.replace(r'^\s*$', None, regex=True)
    df_clean = df_clean.dropna(how="all", axis=0).dropna(how="all", axis=1)
    
    rows_after, cols_after = df_clean.shape

    # Gunakan baris pertama sebagai header jika masih Unnamed
    if any("Unnamed" in str(c) for c in df_clean.columns):
        df_clean.columns = df_clean.iloc[0]
        df_clean = df_clean[1:].reset_index(drop=True)

    # Konversi tipe otomatis
    df_clean = df_clean.convert_dtypes()

    # Hapus dtype numpy
    def safe_convert(x):
        if isinstance(x, (np.generic, np.number)):
            return x.item()
        return x 
    
    df_clean = df_clean.map(safe_convert)
    df_clean.columns = [safe_convert(c) for c in df_clean.columns]
    df_clean.index = [safe_convert(i) for i in df_clean.index]

    # Paksa header & index ke string agar JSON safe
    df_clean.columns = df_clean.columns.map(str)
    df_clean.index = df_clean.index.map(str)

    # Format
    num_cols = df_clean.select_dtypes(include=["number"]).columns
    df_clean[num_cols] = df_clean[num_cols].apply(round_half_up)
    df_clean_styled = (
        df_clean.style
        .format({col: format_rupiah for col in num_cols})
    )

    st.caption(f"The are **{len(df_clean.columns)-1} participating bidders** in this session.")
    st.dataframe(df_clean_styled, hide_index=True)

    # --- NOTIFIKASI KHUSUS ---
    if (rows_after < rows_before) or (cols_after < cols_before):
        st.markdown(
            "<p style='font-size:12px; color:#808080; margin-top:-15px; margin-bottom:0;'>"
            "Preprocessing completed! Hidden rows and columns removed ✅</p>",
            unsafe_allow_html=True
        )

    # Simpan session
    st.session_state['result_standard_deviation'] = df_clean
    st.divider()

    # RANKKK
    st.markdown("##### 🥇 Bidder's Rank")
    st.caption("The bidder ranking process has been successfully completed.")

    vendor_cols = df_clean.columns[1:]          # numeric column vendor (dynamic)
    df_rank = df_clean[[df_clean.columns[0]]].copy()  # ambil kolom pertama (scope)
    df_rank[vendor_cols] = (
        df_clean[vendor_cols]
        .rank(axis=1, method="min", ascending=True)
        .astype('Int64')
    )

    st.dataframe(df_rank, hide_index=True)

    # Download
    excel_data = get_excel_download(df_rank)

    # Layout tombol (rata kanan)
    col1, col2, col3 = st.columns([2.3,2,1])
    with col3:
        st.download_button(
            label="Download",
            data=excel_data,
            file_name="Bidder_Rank.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            icon=":material/download:",
        )

    st.divider()

    # RANK-1 DEVIATIONNN
    st.markdown("##### 🛸 Rank-1 Deviation (%)")
    st.caption("This table shows each vendor’s price deviation (%) from the lowest-priced (Rank-1) vendor per item.")

    min_value = df_clean[vendor_cols].min(axis=1, skipna=True)
    # Cari vendor dengan nilai minimum
    best_vendor = df_clean[vendor_cols].idxmin(axis=1, skipna=True)
    
    # # Simpan hasil ke dataframe baru
    # df_min_vendor = pd.DataFrame({
    #     df_clean.columns[0]: df_clean[df_clean[0]],
    #     "best_vendor": best_vendor,
    #     "best_price": min_value 
    # })

    # Buat dataframe deviasi dalam persentase
    df_dev = df_clean[[df_clean.columns[0]]].copy()
    for col in vendor_cols:
        df_dev[col] = ((df_clean[col] - min_value) / min_value) * 100

    # Abaikan vendor yang tidak ikut (No-Bid)
    df_dev[vendor_cols] = df_dev[vendor_cols].where(~df_clean[vendor_cols].isna(), np.nan)

    # Format
    num_cols = df_dev.select_dtypes(include=["number"]).columns
    format_dict = {col: format_rupiah_percent for col in num_cols}

    df_dev_styled = df_dev.style.format(format_dict)

    st.dataframe(df_dev_styled, hide_index=True)

    # Download
    excel_data = get_excel_download(df_dev)

    # Layout tombol (rata kanan)
    col1, col2, col3 = st.columns([2.3,2,1])
    with col3:
        st.download_button(
            label="Download",
            data=excel_data,
            file_name="Rank-1 Deviation (%).xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            icon=":material/download:",
        )

    st.divider()

    # SUMMARYY DEVIATIONN
    st.markdown("##### 🌍 Summary Deviation (%)")
    st.caption("This table summarizes vendor rankings and their deviation (%) from the Rank-1 vendor for each item.")

    scope_col = df_clean.columns[0]
    vendor_cols = df_clean.columns[1:]

    # Ubah ke long format
    df_long = df_clean.melt(id_vars=[scope_col], var_name="Vendor", value_name="Price").dropna(subset=["Price"])
    df_long["Rank"] = (
        df_long.groupby(scope_col)["Price"]
        .rank(method="min", ascending=True)
    )

    summary_rows = []
    for comp, group in df_long.groupby(scope_col):
        group = group.sort_values("Rank").reset_index(drop=True)
        base_price = group.loc[0, "Price"]
        row_data = {
            df_clean.columns[0]: comp,
            "1st Rank": group.loc[0, "Vendor"],
            "Best Price": base_price
        }

        # Tambahkan 2nd, 3rd, dst secara horizontal
        for i in range(1, len(group)):
            rank = i+1
            vendor = group.loc[i, "Vendor"]
            price = group.loc[i, "Price"]
            deviation = ((price - base_price) / base_price) * 100
            row_data[f"{rank}th Rank"] = vendor
            row_data[f"Dev. {rank}th to 1st (%)"] = deviation
        summary_rows.append(row_data)

    df_summary = pd.DataFrame(summary_rows)

    # Format
    format_dict = {}

    # Kolom "Best Price"
    if "Best Price" in df_summary.columns:
        format_dict["Best Price"] = format_rupiah
    
    # Kolom deviasi (%)
    for col in df_summary.columns:
        if col.startswith("Dev. ") and col.endswith("(%)"):
            format_dict[col] = format_rupiah_percent
    
    df_summary_styled = df_summary.style.format(format_dict)

    st.dataframe(df_summary_styled, hide_index=True)

    # Download
    excel_data = get_excel_download(df_summary)

    # Layout tombol (rata kanan)
    col1, col2, col3 = st.columns([2.3,2,1])
    with col3:
        st.download_button(
            label="Download",
            data=excel_data,
            file_name="Summary Deviation (%).xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            icon=":material/download:",
        )

    # # RANK VISUALIZATION
    # st.subheader(f"📊 Rank Visualization")

    # # Rank Visualization per Component
    # col0 = df.columns[0]  # kolom index pertama
    # vendor_cols = df.columns[1:]

    # # Buat tab berdasarkan nilai unik kolom pertama
    # tab_names = df[col0].unique()
    # tabs = st.tabs([str(name) for name in tab_names])

    # for i, tab in enumerate(tabs):
    #     tab_name = tab_names[i]
    #     # tab.subheader(f"📊 Rank Visualization")
    #     # tab.caption("Hooray! You’ve got the bidder with the lowest offer 🎉")

    #     # Ambil data untuk tab ini
    #     df_tab = df[df[col0] == tab_name][vendor_cols].copy()

    #     # Hitung total per vendor (di sini pakai sum jika ada multiple rows per tab)
    #     ranked_sum = df_tab.sum().sort_values(ascending=True)

    #     # Siapkan dataframe chart
    #     df_chart = (
    #         ranked_sum.reset_index()
    #         .rename(columns={"index": "Vendor", 0: "Total"})
    #         .sort_values("Total", ascending=True)
    #     )

    #     # Filter vendor dengan nilai 0 atau None
    #     df_chart_filtered = df_chart[df_chart["Total"] > 0].copy()
    #     df_chart_filtered["Rank"] = range(1, len(df_chart_filtered) + 1)
    #     df_chart_filtered["Mid"] = df_chart_filtered["Total"] / 2

    #     # Format string ribuan
    #     df_chart_filtered["Total_str"] = df_chart_filtered["Total"].apply(lambda x: f"{int(x):,}".replace(",", "."))
    #     df_chart_filtered["Legend"] = df_chart_filtered.apply(lambda x: f"Rank {x['Rank']} - {x['Total_str']}", axis=1)

    #     # Warna manual per vendor
    #     colors_list = ["#F94144", "#F3722C", "#F8961E", "#F9C74F", "#90BE6D", "#43AA8B", "#577590", "#E54787", "#BF219A", "#8E0F9C", "#4B1D91"]
    #     vendor_colors = {v: c for v, c in zip(df_chart_filtered["Legend"], colors_list)}

    #     highlight = alt.selection_point(on='mouseover', fields=['Vendor'], nearest=True)

    #     # Bars
    #     bars = (
    #         alt.Chart(df_chart_filtered)
    #         .mark_bar()
    #         .encode(
    #             x=alt.X("Vendor:N", axis=alt.Axis(title=None)),
    #             y=alt.Y("Total:Q", axis=alt.Axis(title=None, grid=False),
    #                     scale=alt.Scale(domain=[0, df_chart_filtered["Total"].max() * 1.1])
    #             ),
    #             color=alt.Color("Legend:N", title="Total Offer by Rank",
    #             scale=alt.Scale(domain=list(vendor_colors.keys()), 
    #                             range=list(vendor_colors.values()))
    #             ),
    #             tooltip=[
    #                 alt.Tooltip("Vendor:N", title="Vendor"),
    #                 alt.Tooltip("Total_str:N", title="Total (USD)")
    #             ]
    #         ).add_params(highlight)
    #     )

    #     tab.caption(" ")
    #     # Rank text
    #     rank_text = (
    #         alt.Chart(df_chart_filtered)
    #         .mark_text(
    #             dy=-7,           # geser teks sedikit ke atas
    #             color="gray", 
    #             fontWeight="bold"
    #         )
    #         .encode(
    #             x="Vendor:N",
    #             y="Total:Q",     # di atas bar
    #             text="Rank:N"
    #         )
    #     )

    #     # Border frame
    #     frame = (
    #         alt.Chart()
    #         .mark_rect(stroke='gray', strokeWidth=1, fillOpacity=0)
    #     )

    #     # Gabungkan chart
    #     chart = (bars + rank_text + frame).properties(
    #         title=f"{tab_name}: Comparative Bidder Ranking"
    #     ).configure_title(
    #         anchor='middle', 
    #         offset=12
    #     ).configure_legend(
    #         titleFontSize=12,        
    #         titleFontWeight="bold",  
    #         labelFontSize=12,        
    #         labelLimit=300,        
    #         orient="right"
    #     )

    #     # Tampilkan
    #     tab.altair_chart(chart)

    # # st.divider()
