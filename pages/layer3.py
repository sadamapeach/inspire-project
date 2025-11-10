import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
import time
import re
from io import BytesIO

def page():
    # Header Title
    st.title("3️⃣ Standard Deviation")
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
    st.subheader("📂 Upload File")
    upload_file = st.file_uploader("Upload your file here!", type=["xlsx", "xls"])

    if upload_file is not None:
        st.session_state["uploaded_file_layer3"] = upload_file

        # --- Animasi proses upload ---
        msg = st.toast("📂 Uploading file...")
        time.sleep(1.2)

        msg.toast("🔍 Reading sheets...")
        time.sleep(1.2)

        msg.toast("✅ File uploaded successfully!")
        time.sleep(0.5)

        # Baca sheet
        df = pd.read_excel(upload_file)

        # # Konversi kolom numerik ke float
        # for col in df.select_dtypes(include=["number"]).columns:
        #     df[col] = df[col].astype(float)

        # Simpan versi mentah (setelah konversi)
        st.session_state["df_layer3_raw"] = df

    elif "df_layer3_raw" in st.session_state:
        df = st.session_state["df_layer3_raw"]
    else:
        return
    
    st.divider()

    # OVERVIEW
    st.subheader("🔍 Overview")

    # Fungsi styling untuk highlight 'No-Bid'
    def highlight_no_bid(val):
        if val == "No-Bid":
            # warna pink lembut dengan teks bold
            # return "background-color: #ffe6f2; color: #d63384; font-weight: bold; text-align: center;"
            # return "background-color: #fff3cd; color: #856404; font-weight: bold; text-align: center;"
            # return "background-color: #ffeb99; color: #7a5c00; font-weight: bold; text-align: center;"
            # return "background-color: #ffcccc; color: #b30000; font-weight: bold; text-align: center;"
            return "background-color: #f8c8dc; color: #7a1f47; font-weight: bold; text-align: center;"
        return "text-align: center;"

    # Hide case
    if df.isna().all(axis=1).any() or df.isna().all(axis=0).any():
        # Pre-processing: hapus baris & kolom kosong total
        df_clean = df.dropna(how="all").dropna(axis=1, how="all")

        # Gunakan baris pertama sebagai header
        df_clean.columns = df_clean.iloc[0]
        df_clean = df_clean[1:].reset_index(drop=True)

        # Konversi tipe data otomatis
        df = df_clean.convert_dtypes()

        # Bersihkan semua kemungkinan tipe numpy di kolom, index, dan isi
        def safe_convert(x):
            if isinstance(x, (np.generic, np.number)):
                return x.item()
            return x

        df = df.map(safe_convert)  # untuk isi dataframe
        df.columns = [safe_convert(c) for c in df.columns]  # kolom
        df.index = [safe_convert(i) for i in df.index]      # index

        # Paksa semua header & index ke string agar JSON safe
        df.columns = df.columns.map(str)
        df.index = df.index.map(str)

        # Total bidders
        total_bidders = len(df.columns) - 1  # minus TCO Component (indeks[0])
        st.caption(f"There are **{total_bidders} bidders** competing for the best offer in this session 🤯")

        # Tampilan 'No-Bid'
        df_hide_display = df.copy()    

        # Konversi angka float ke integer string (agar tidak tampil .0)
        df_hide_display = df_hide_display.apply(
            lambda col: col.map(
                lambda x: f"{int(x)}" if isinstance(x, (int, float)) and pd.notna(x) and float(x).is_integer() else x
            )
        )

        # Ubah NaN jadi 'No-Bid'
        df_hide_display = df_hide_display.fillna("No-Bid").astype("object")

        # Tampilkan di Streamlit
        st.dataframe(
            df_hide_display.style.map(highlight_no_bid),
            hide_index=True
        )

        # st.dataframe(df, hide_index=True)
        st.markdown(
            """
            <p style='font-size:12px; color:#808080; margin-top:-15px; margin-bottom:0;'>
                Preprocessing completed! Hidden rows and columns removed ✅
            </p>
            """,
            unsafe_allow_html=True
        )
    
    else:
        # Total bidders
        total_bidders = len(df.columns) - 1  # minus TCO Component (indeks[0])
        st.caption(f"There are **{total_bidders} bidders** competing for the best offer in this session 🤯")

        # Tampilan 'No-Bid'
        df_raw_display = df.copy()

        # Konversi angka float ke integer string (agar tidak tampil .0)
        df_raw_display = df_raw_display.apply(
            lambda col: col.map(
                lambda x: f"{int(x)}" if isinstance(x, (int, float)) and pd.notna(x) and float(x).is_integer() else x
            )
        )

        # Ubah NaN jadi 'No-Bid'
        df_raw_display = df_raw_display.fillna("No-Bid").astype("object")

        # Tampilkan di Streamlit
        st.dataframe(
            df_raw_display.style.map(highlight_no_bid),
            hide_index=True
        )

        # st.dataframe(df, hide_index=True)
        st.markdown(
            """
            <p style='font-size:12px; color:#808080; margin-top:-15px; margin-bottom:0;'>
                No hidden rows and columns detected. Proceeding with raw data ✅
            </p>
            """,
            unsafe_allow_html=True
        )

    st.divider()

    # Rank
    st.subheader("🥇 Bidder's Rank")
    st.caption("Awesome! All bidders have been successfully ranked 🎉")

    # Logic
    vendor_cols = df.columns[1:]    # kolom vendor
    
    df_rank = df[[df.columns[0]]].copy()    # ambil kolom pertama (WP)
    df_rank[vendor_cols] = (
        df[vendor_cols]
        .rank(axis=1, method='min', ascending=True)
        .astype('Int64')
    )

    # Tampilan 'No-Bid'
    df_rank_display = df_rank.copy()

    # Konversi angka float ke integer string (agar tidak tampil .0)
    df_rank_display = df_rank_display.apply(
        lambda col: col.map(
            lambda x: f"{int(x)}" if isinstance(x, (int, float)) and pd.notna(x) and float(x).is_integer() else x
        )
    )

    # Ubah NaN jadi 'No-Bid'
    df_rank_display = df_rank_display.fillna("No-Bid").astype("object")

    # Tampilkan di Streamlit
    st.dataframe(
        df_rank_display.style.map(highlight_no_bid),
        hide_index=True
    )

    # Download button to Excel
    @st.cache_data
    def get_excel_download(df_rank_display, sheet_name="Bidder Rank"):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_rank_display.to_excel(writer, index=False, sheet_name=sheet_name)
        return output.getvalue()

    # Simpan hasil ke variabel
    excel_data = get_excel_download(df_rank_display)

    # Layout tombol (rata kanan)
    col1, col2, col3 = st.columns([3,1,1])
    with col3:
        st.download_button(
            label="Download",
            data=excel_data,
            file_name="Bidder_Rank.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            icon=":material/download:",
        )

    # RANK VISUALIZATION
    st.subheader(f"📊 Rank Visualization")

    # Rank Visualization per Component
    col0 = df.columns[0]  # kolom index pertama
    vendor_cols = df.columns[1:]

    # Buat tab berdasarkan nilai unik kolom pertama
    tab_names = df[col0].unique()
    tabs = st.tabs([str(name) for name in tab_names])

    for i, tab in enumerate(tabs):
        tab_name = tab_names[i]
        # tab.subheader(f"📊 Rank Visualization")
        # tab.caption("Hooray! You’ve got the bidder with the lowest offer 🎉")

        # Ambil data untuk tab ini
        df_tab = df[df[col0] == tab_name][vendor_cols].copy()

        # Hitung total per vendor (di sini pakai sum jika ada multiple rows per tab)
        ranked_sum = df_tab.sum().sort_values(ascending=True)

        # Siapkan dataframe chart
        df_chart = (
            ranked_sum.reset_index()
            .rename(columns={"index": "Vendor", 0: "Total"})
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

        highlight = alt.selection_point(on='mouseover', fields=['Vendor'], nearest=True)

        # Bars
        bars = (
            alt.Chart(df_chart_filtered)
            .mark_bar()
            .encode(
                x=alt.X("Vendor:N", axis=alt.Axis(title=None)),
                y=alt.Y("Total:Q", axis=alt.Axis(title=None, grid=False),
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

        tab.caption(" ")
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
            offset=12
        ).configure_legend(
            titleFontSize=12,        
            titleFontWeight="bold",  
            labelFontSize=12,        
            labelLimit=300,        
            orient="right"
        )

        # Tampilkan
        tab.altair_chart(chart)

    # st.divider()

    # STANDARD DEVIATION
    st.subheader("📈 Standard Deviation")
    # st.caption("blbalanbsmlaslka")
    tab1, tab2 = st.tabs(["🏅 Rank-1 Deviation (%)", "🌐 Summary Deviation (%)"])

    # Rank-1 Deviation
    min_value = df[vendor_cols].min(axis=1, skipna=True)

    # Cari vendor dengan nilai minimum
    best_vendor = df[vendor_cols].idxmin(axis=1, skipna=True)

    # Simpan hasil ke dataframe baru
    df_min_vendor = pd.DataFrame({
        df.columns[0]: df[df.columns[0]],
        "best_vendor": best_vendor,
        "best_price": min_value
    })

    # Buat dataframe deviasi dalam persentase
    df_dev = df[[df.columns[0]]].copy()
    for col in vendor_cols:
        df_dev[col] = ((df[col] - min_value) / min_value) * 100

    # Abaikan vendor yang tidak ikut
    df_dev[vendor_cols] = df_dev[vendor_cols].where(~df[vendor_cols].isna(), np.nan)

    # --- Highlight pakai referensi df_min_vendor ---
    def highlight_cells(row):
        comp_name = row[df.columns[0]]
        best_v = df_min_vendor.loc[
            df_min_vendor[df.columns[0]] == comp_name, "best_vendor"
        ].values[0]

        styles = {}
        for col in vendor_cols:
            val = row[col]
            if val == "No-Bid":
                styles[col] = "background-color: #f8c8dc; color: #7a1f47; font-weight: bold;"
            elif col == best_v:
                styles[col] = "background-color: #d7c6f3; color: #402e72; font-weight: bold;"
            else:
                styles[col] = ""
        return pd.Series(styles)

    # --- Format tampilan ---
    df_dev_display = df_dev.copy()

    # Ubah NaN jadi 'No-Bid'
    df_dev_display = df_dev_display.fillna("No-Bid").astype("object")

    for col in vendor_cols:
        df_dev_display[col] = df_dev_display[col].apply(
            lambda x: f"{x:.2f}%" if isinstance(x, (int, float)) else x
        )

    # Terapkan highlight berdasarkan df_min_vendor
    styled_df = df_dev_display.style.apply(highlight_cells, axis=1)

    # Dataframe
    tab1.dataframe(styled_df, hide_index=True)

    # Download button to Excel
    @st.cache_data
    def get_excel_download2(df_dev_display, sheet_name="Rank-1 Deviation"):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_dev_display.to_excel(writer, index=False, sheet_name=sheet_name)
        return output.getvalue()

    # Simpan hasil ke variabel
    excel_data2 = get_excel_download2(df_dev_display)
    tab1.download_button(
        label="Download",
        data=excel_data2,
        file_name="Rank1_Deviation.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        icon=":material/download:",
    )

    # Overall Deviation (%)
    col0 = df.columns[0]
    vendor_cols = df.columns[1:]

    # Ubah ke long format
    df_long = df.melt(id_vars=[col0], var_name="Vendor", value_name="Price").dropna(subset=["Price"])

    # Rank per komponen
    df_long["Rank"] = df_long.groupby(col0)["Price"].rank(method="min", ascending=True)

    # Fungsi buat hasil per komponen
    summary_rows = []

    for comp, group in df_long.groupby(col0):
        group = group.sort_values("Rank").reset_index(drop=True)
        base_price = group.loc[0, "Price"]
        row_data = {
            df.columns[0]: comp,
            "1st Rank": group.loc[0, "Vendor"],
            "Best Price": base_price
        }
        # Tambahkan 2nd, 3rd, dst secara horizontal
        for i in range(1, len(group)):
            rank = i + 1
            vendor = group.loc[i, "Vendor"]
            price = group.loc[i, "Price"]
            deviation = ((price - base_price) / base_price) * 100
            row_data[f"{rank}th Rank"] = vendor
            row_data[f"Dev. {rank}th vs 1st (%)"] = deviation
        summary_rows.append(row_data)

    # Gabung jadi dataframe
    df_overall = pd.DataFrame(summary_rows)

    # Format deviasi
    for col in df_overall.columns:
        if "Dev." in col:
            df_overall[col] = df_overall[col].apply(lambda x: f"{x:.2f}%" if pd.notna(x) else "")

    # Ganti NaN dengan string kosong
    df_overall = df_overall.fillna("")

    # Konversi angka float ke integer string (agar tidak tampil .0)
    df_overall = df_overall.apply(
        lambda col: col.map(
            lambda x: f"{int(x)}" if isinstance(x, (int, float)) and pd.notna(x) and float(x).is_integer() else x
        )
    )

    # --- Fungsi highlight pink untuk cell kosong ---
    def highlight_empty(val):
        if val == "" or pd.isna(val):
            return "background-color: #f8c8dc; color: #7a1f47;"
        return ""

    # --- Tampilkan tabel di Streamlit ---
    tab2.dataframe(
        df_overall.style.map(highlight_empty),
        hide_index=True
    )

    # Download button to Excel
    @st.cache_data
    def get_excel_download3(df_overall, sheet_name="Overall Deviation"):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_overall.to_excel(writer, index=False, sheet_name=sheet_name)
        return output.getvalue()

    # Simpan hasil ke variabel
    excel_data2 = get_excel_download3(df_overall)
    tab2.download_button(
        label="Download",
        data=excel_data2,
        file_name="Overall_Deviation.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        icon=":material/download:",
    )
