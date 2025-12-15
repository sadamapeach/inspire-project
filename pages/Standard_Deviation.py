import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
import time
import math
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

def highlight_min_cell(row):
    styles = []
    
    # Cari nilai minimum, abaikan NaN
    numeric_vals = row[row.apply(lambda x: isinstance(x, (int, float)))]
    if not numeric_vals.empty:
        min_val = numeric_vals.min()
    else:
        min_val = None

    # Buat style per cell
    for val in row:
        if val == min_val:
            styles.append("background-color: #C6EFCE; color: #006100;")
        else:
            styles.append("")
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
def get_excel_download_highlight(df, sheet_name="Sheet1"):
    output = BytesIO()

    # ===== COPY & COERCE NUMERIC =====
    df_to_write = df.copy()
    numeric_cols = []

    for col in df_to_write.columns:
        coerced = pd.to_numeric(df_to_write[col], errors="coerce")
        if coerced.notna().any():
            df_to_write[col] = coerced
            numeric_cols.append(col)

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_to_write.to_excel(writer, index=False, sheet_name=sheet_name)
        workbook  = writer.book
        worksheet = writer.sheets[sheet_name]

        # ===== FORMAT =====
        fmt_pct = workbook.add_format({'num_format': '#,##0.0"%"'})

        fmt_min = workbook.add_format({
            "bold": True,
            "bg_color": "#D9EAD3",
            "font_color": "#1A5E20",
            'num_format': '#,##0.0"%"'
        })

        # ===== COLUMN WIDTH =====
        for col_idx in range(len(df_to_write.columns)):
            worksheet.set_column(col_idx, col_idx, 15)

        # ===== LOOP ALL CELLS =====
        for row_idx, row in enumerate(df_to_write.itertuples(index=False), start=1):

            numeric_vals = {
                idx: val for idx, val in enumerate(row)
                if idx < len(df_to_write.columns)
                and df_to_write.columns[idx] in numeric_cols
                and pd.notna(val)
            }

            min_val = min(numeric_vals.values()) if numeric_vals else None

            for col_idx, col_name in enumerate(df_to_write.columns):
                value = row[col_idx]

                # ===== SAFETY =====
                if pd.isna(value):
                    worksheet.write_blank(row_idx, col_idx, None)
                    continue

                # ===== PICK FORMAT =====
                if col_name in numeric_cols:
                    fmt = fmt_min if value == min_val else fmt_pct
                    worksheet.write_number(row_idx, col_idx, value, fmt)
                else:
                    worksheet.write(row_idx, col_idx, value)

        # ===== AUTOFIT =====
        for i, col in enumerate(df_to_write.columns):
            worksheet.set_column(
                i, i,
                max(len(str(col)), df_to_write[col].astype(str).map(len).max()) + 2
            )

    output.seek(0)
    return output.getvalue()

def page():
    # Header Title
    st.markdown(
        """
        <div style="font-size:2.25rem; font-weight:700; margin-bottom:9px">
            7️⃣ Standard Deviation
        </div>
        """,
        unsafe_allow_html=True
    )
    # st.header("7️⃣ Standard Deviation")
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

    # === Identifikasi kolom ===
    non_num_cols = df_clean.select_dtypes(exclude=["number"]).columns.tolist()
    vendor_cols = df_clean.select_dtypes(include=["number"]).columns.tolist()

    # RANKKK
    st.markdown("##### 🥇 Bidder's Rank")
    st.caption("The bidder ranking process has been successfully completed.")

    # Copy non-numeric col
    df_rank = df_clean[non_num_cols].copy()

    # Hitung rank
    df_rank[vendor_cols] = (
        df_clean[vendor_cols]
        .rank(axis=1, method="min", ascending=True)
        .astype("Int64")
    )

    st.dataframe(df_rank, hide_index=True)

    # Download
    excel_data = get_excel_download(df_rank)
    col1, col2, col3 = st.columns([2.3,2,1])
    with col3:
        st.download_button(
            label="Download",
            data=excel_data,
            file_name="Bidder's Rank - Standard Deviation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            icon=":material/download:",
        )

    st.divider()

    # RANK-1 DEVIATIONNN
    st.markdown("##### 🛸 Rank-1 Deviation (%)")
    st.caption("This table shows each vendor’s price deviation (%) from the lowest-priced (Rank-1) vendor per item.")

    min_value = df_clean[vendor_cols].min(axis=1, skipna=True)
    
    # Buat dataframe deviasi dalam persentase
    df_dev = df_clean[non_num_cols].copy()
    for col in vendor_cols:
        df_dev[col] = ((df_clean[col] - min_value) / min_value) * 100

    # Abaikan vendor yang tidak ikut (No-Bid)
    df_dev[vendor_cols] = df_dev[vendor_cols].where(~df_clean[vendor_cols].isna(), np.nan)

    # Format
    num_cols = df_dev.select_dtypes(include=["number"]).columns
    format_dict = {col: format_rupiah_percent for col in num_cols}

    df_dev_styled = (
        df_dev.style
        .format(format_dict)
        .apply(highlight_min_cell, axis=1)
    )

    st.dataframe(df_dev_styled, hide_index=True)

    # Download
    excel_data = get_excel_download_highlight(df_dev)
    col1, col2, col3 = st.columns([2.3,2,1])
    with col3:
        st.download_button(
            label="Download",
            data=excel_data,
            file_name="Rank-1 Deviation (%) - Standard Deviation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            icon=":material/download:",
        )

    st.divider()

    # SUMMARYY DEVIATIONN
    st.markdown("##### 🌍 Summary Deviation (%)")
    st.caption("This table summarizes vendor rankings and their deviation (%) from the Rank-1 vendor for each item.")

    # Ambil kolom utama -> non_num[0]
    main_col = non_num_cols[0
                            ]
    # Ubah ke long format
    df_long = df_clean.melt(
        id_vars=non_num_cols, 
        var_name="Vendor", 
        value_name="[PRICE]"
    ).dropna(subset=["[PRICE]"])

    # Rank
    df_long["Rank"] = df_long.groupby(non_num_cols)["[PRICE]"].rank(method="min")

    # Fungsi ordinal
    def ordinal(n):
        if 10 <= n % 100 <= 20:
            suf = "th"
        else:
            suf = {1:"st",2:"nd",3:"rd"}.get(n % 10, "th")
        return f"{n}{suf}"

    summary_rows = []

    # Grouping
    for keys, group in df_long.groupby(non_num_cols):
        group = group.sort_values("Rank").reset_index(drop=True)
        base_price = group.loc[0, "[PRICE]"]

        row_data = {}

        # isi semua kolom non-num
        for i, col in enumerate(non_num_cols):
            row_data[col] = keys[i]

        row_data["1st Rank"] = group.loc[0, "Vendor"]
        row_data["Best Price"] = base_price

        # 2nd, 3rd dst
        for i in range(1, len(group)):
            r = i + 1
            vendor = group.loc[i, "Vendor"]
            price = group.loc[i, "[PRICE]"]

            deviation = (
                ((price - base_price) / base_price) * 100
                if base_price not in (0, np.nan)
                else np.nan
            )

            row_data[f"{ordinal(r)} Rank"] = vendor
            row_data[f"Dev. {ordinal(r)} to 1st (%)"] = deviation

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
    col1, col2, col3 = st.columns([2.3,2,1])
    with col3:
        st.download_button(
            label="Download",
            data=excel_data,
            file_name="Summary Deviation (%).xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            icon=":material/download:",
        )

    # VISUALIZATIONN
    st.markdown("##### 📊 Visualization")

    scope_col = non_num_cols[0]

    # Tab
    tab_names = df_clean[scope_col].unique()
    tabs = st.tabs([str(name) for name in tab_names])

    for i, tab in enumerate(tabs):
        tab_name = tab_names[i]

        # Ambil data untuk tab ini
        df_tab = df_clean[df_clean[scope_col] == tab_name][vendor_cols].copy()

        # Hitung total per vendor
        ranked_sum = df_tab.sum().sort_values(ascending=True)

        df_chart = (
            ranked_sum.reset_index()
            .rename(columns={"index": "Vendor", 0:"Total"})
            .sort_values("Total", ascending=True)
        )

        # Filter vendor dengan nilai 0 atau None
        df_chart_filtered = df_chart[df_chart["Total"] > 0].copy()
        df_chart_filtered["Rank"] = range(1, len(df_chart_filtered) + 1)
        df_chart_filtered["Mid"] = df_chart_filtered["Total"] / 2

        # Format string ribuan
        df_chart_filtered["Total_str"] = df_chart_filtered["Total"].apply(lambda x: f"{int(x):,}".replace(",", "."))
        df_chart_filtered["Legend"] = df_chart_filtered.apply(lambda x: f"Rank {x['Rank']} - {x['Total_str']}", axis=1)

        # Warna manual per vendor
        colors_list = ["#F94144", "#F3722C", "#F8961E", "#F9C74F", "#90BE6D", "#43AA8B", "#577590", "#E54787", "#BF219A", "#8E0F9C", "#4B1D91"]
        vendor_colors = {v: c for v, c in zip(df_chart_filtered["Legend"], colors_list)}

        # Format angka besar di sumbu Y → jadi singkat (1K, 1M)
        y_axis = alt.Axis(title=None, grid=False, format=".0s", labelPadding=12)

        highlight = alt.selection_point(on='mouseover', fields=['Vendor'], nearest=True)

        # Bars
        bars = (
            alt.Chart(df_chart_filtered)
            .mark_bar()
            .encode(
                x=alt.X("Vendor:N", axis=alt.Axis(title=None)),
                y=alt.Y("Total:Q", axis=y_axis,
                        scale=alt.Scale(domain=[0, df_chart_filtered["Total"].max() * 1.1])
                ),
                color=alt.Color("Legend:N", title="Total Offer by Rank",
                scale=alt.Scale(domain=list(vendor_colors.keys()), 
                                range=list(vendor_colors.values()))
                ),
                tooltip=[
                    alt.Tooltip("Vendor:N", title="Vendor"),
                    alt.Tooltip("Total_str:N", title="Total (USD)")
                ]
            ).add_params(highlight)
        )

        # Rank text
        rank_text = (
            alt.Chart(df_chart_filtered)
            .mark_text(
                dy=-7,           # geser teks sedikit ke atas
                color="gray", 
                fontWeight="bold"
            )
            .encode(
                x="Vendor:N",
                y="Total:Q",     # di atas bar
                text="Rank:N"
            )
        )

        # Border frame
        frame = (
            alt.Chart()
            .mark_rect(stroke='gray', strokeWidth=1, fillOpacity=0)
        )

        # Gabungkan chart
        chart = (bars + rank_text + frame).properties(
            title=f"{tab_name}: Comparative Bidder Ranking"
        ).configure_title(
            anchor='middle',
            fontSize=13, 
            offset=12,
            fontWeight="bold"
        ).configure_legend(
            titleFontSize=11,        
            titleFontWeight="bold",  
            labelFontSize=10,        
            labelLimit=300,        
            orient="right"
        )

        # Tampilkan
        tab.altair_chart(chart)

    st.divider()

    # SUPERR BUTTON
    st.markdown("##### 🧑‍💻 Super Download — Export Selected Sheets")
    dataframes = {
        "Bidder's Rank": df_rank,
        "Rank-1 Deviation (%)": df_dev,
        "Summary Deviation (%)": df_summary,
    }

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
                df = df_dict[sheet].copy()

                # ========= SPECIAL SHEET =========
                if sheet == "Rank-1 Deviation (%)":
                    df_to_write = df.copy()
                    numeric_cols = [
                        c for c in df_to_write.columns
                        if pd.api.types.is_numeric_dtype(df_to_write[c])
                    ]

                    df_to_write.to_excel(writer, index=False, sheet_name=sheet)
                    workbook  = writer.book
                    worksheet = writer.sheets[sheet]

                    fmt_pct = workbook.add_format({'num_format': '#,##0.0"%"'})
                    fmt_min = workbook.add_format({
                        "bold": True,
                        "bg_color": "#D9EAD3",
                        "font_color": "#1A5E20",
                        'num_format': '#,##0.0"%"'
                    })

                    # ===== WRITE + FORMAT CELL =====
                    for r, row in enumerate(df_to_write.itertuples(index=False), start=1):
                        numeric_vals = {
                            i: val for i, val in enumerate(row)
                            if df_to_write.columns[i] in numeric_cols and pd.notna(val)
                        }
                        min_val = min(numeric_vals.values()) if numeric_vals else None

                        for c, col in enumerate(df_to_write.columns):
                            val = row[c]

                            if pd.isna(val):
                                worksheet.write_blank(r, c, None)
                            elif col in numeric_cols:
                                fmt = fmt_min if val == min_val else fmt_pct
                                worksheet.write_number(r, c, val, fmt)
                            else:
                                worksheet.write(r, c, val)

                    # ===== AUTOFIT (KEEP FORMAT!) =====
                    for i, col in enumerate(df_to_write.columns):
                        width = max(
                            len(str(col)),
                            df_to_write[col].astype(str).map(len).max()
                        ) + 2

                        worksheet.set_column(
                            i, i,
                            width,
                            fmt_pct if col in numeric_cols else None
                        )

                    continue

                # ========= DEFAULT SHEETS =========
                df.to_excel(writer, index=False, sheet_name=sheet)
                workbook  = writer.book
                worksheet = writer.sheets[sheet]

                fmt_rp  = workbook.add_format({'num_format': '#,##0'})
                fmt_pct = workbook.add_format({'num_format': '#,##0.0"%"'})

                numeric_cols = df.select_dtypes(include=["number"]).columns.tolist()
                pct_cols = [c for c in df.columns if "%" in c]

                # ===== REWRITE CELLS (SAFE) =====
                for r, row in enumerate(df.itertuples(index=False), start=1):
                    for c, col in enumerate(df.columns):
                        val = row[c]

                        if pd.isna(val) or (isinstance(val, float) and np.isinf(val)):
                            worksheet.write_blank(r, c, None)
                        elif col in pct_cols:
                            worksheet.write_number(r, c, val, fmt_pct)
                        elif col in numeric_cols:
                            worksheet.write_number(r, c, val, fmt_rp)
                        else:
                            worksheet.write(r, c, val)

                # ===== AUTOFIT (KEEP FORMAT!) =====
                for i, col in enumerate(df.columns):
                    width = max(
                        len(str(col)),
                        df[col].astype(str).map(len).max()
                    ) + 2

                    if col in pct_cols:
                        worksheet.set_column(i, i, width, fmt_pct)
                    elif col in numeric_cols:
                        worksheet.set_column(i, i, width, fmt_rp)
                    else:
                        worksheet.set_column(i, i, width)

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
            file_name="Standard Deviation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            on_click=release_the_balloons,
            type="primary",
            use_container_width=True,
        )

