import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
import time
import re
from io import BytesIO
from decimal import Decimal, ROUND_HALF_UP

def format_thousand(value):
    """Format angka ke format Indonesia (1.000.000), hilangkan 0 jadi kosong."""
    try:
        if pd.isna(value) or value == 0:
            return ""  # kosongkan nilai 0 atau NaN
        return f"{int(value):,}".replace(",", ".")
    except (ValueError, TypeError):
        return value
                
def page():
    # Header Title
    st.title("4️⃣ UPL Maker")
    st.markdown(
        ":red-badge[Indosat] :orange-badge[Ooredoo] :green-badge[Hutchison]"
    )
    st.caption("Drop in your annual pricing template and let this analytics system work its magic ✨")

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
        st.session_state["uploaded_file_layer4"] = upload_file

        # --- Animasi proses upload ---
        msg = st.toast("📂 Uploading file...")
        time.sleep(1.2)

        msg.toast("🔍 Reading sheets...")
        time.sleep(1.2)

        msg.toast("✅ File uploaded successfully!")
        time.sleep(0.5)

        # Baca semua sheet sekaligus
        all_df = pd.read_excel(upload_file, sheet_name=None)
        st.session_state["all_df_layer4_raw"] = all_df  # simpan versi mentah

    elif "all_df_layer4_raw" in st.session_state:
        all_df = st.session_state["all_df_layer4_raw"]
    else:
        return
    
    st.divider()

    # OVERVIEW
    st.subheader("🔍 Overview")
    # Total bidders
    total_sheets = len(all_df)
    st.caption(f"You're analyzing offers from **{total_sheets} participating bidders** in this session 🧐")

    result = {}

    for name, df in all_df.items():
        # Hide case
        if df.isna().all(axis=1).any() or df.isna().all(axis=0).any():
            # Pre-processing: hapus baris & kolom kosong total
            df_clean = df.dropna(how="all").dropna(axis=1, how="all")

            # Gunakan baris pertama sebagai header
            df_clean.columns = df_clean.iloc[0]
            df_clean = df_clean[1:].reset_index(drop=True)

            # Konversi tipe data otomatis
            df_clean = df_clean.convert_dtypes()

            # Bersihkan semua kemungkinan tipe numpy di kolom, index, dan isi
            def safe_convert(x):
                if isinstance(x, (np.generic, np.number)):
                    return x.item()
                return x

            df_clean = df_clean.map(safe_convert)  # untuk isi dataframe
            df_clean.columns = [safe_convert(c) for c in df_clean.columns]  # kolom
            df_clean.index = [safe_convert(i) for i in df_clean.index]      # index

            # Paksa semua header & index ke string agar JSON safe
            df_clean.columns = df_clean.columns.map(str)
            df_clean.index = df_clean.index.map(str)

            df_styled = df_clean.copy()
            for col in df_styled.columns:
                # Terapkan hanya untuk kolom numeric
                if pd.api.types.is_numeric_dtype(df_styled[col]):
                    df_styled[col] = df_styled[col].apply(format_thousand)

            st.markdown(
                f"""
                <div style='display: flex; justify-content: space-between; 
                            align-items: center; margin-bottom: 8px;'>
                    <span style='font-size:15px;'>✨ {name}</span>
                    <span style='font-size:12px; color:#808080;'>
                        Total rows: <b>{len(df_clean):,}</b>
                    </span>
                </div>
                """,
                unsafe_allow_html=True
            )

            result[name] = df_clean
            st.dataframe(df_styled, hide_index=True)
            st.markdown(
                """
                <p style='font-size:12px; color:#808080; margin-top:-15px; margin-bottom:0;'>
                    Preprocessing completed! Hidden rows and columns removed ✅
                </p>
                """,
                unsafe_allow_html=True
            )

        else:
            # Jika tidak ada yang di-hide
            df = df.loc[:, ~df.columns.duplicated()].copy()

            # --- 🔹 Terapkan format ribuan di sini juga ---
            df_styled = df.copy()
            for col in df_styled.columns:
                if pd.api.types.is_numeric_dtype(df_styled[col]):
                    df_styled[col] = df_styled[col].apply(format_thousand)

            st.markdown(
                f"""
                <div style='display: flex; justify-content: space-between; 
                            align-items: center; margin-bottom: 8px;'>
                    <span style='font-size:15px;'>✨ {name}</span>
                    <span style='font-size:12px; color:#808080;'>
                        Total rows: <b>{len(df):,}</b>
                    </span>
                </div>
                """,
                unsafe_allow_html=True
            )

            result[name] = df
            st.dataframe(df_styled, hide_index=True)

    st.session_state["result_layer4"] = result
    st.divider()

    # MERGE
    st.subheader("🗃️ Merge Data")

    # --- Gabungkan semua sheet ---
    merged = None

    for i, (name, df_sub) in enumerate(result.items()):
        info_cols = df_sub.columns[:-1]  # kolom template (semua kecuali terakhir)
        value_col = df_sub.columns[-1]   # kolom terakhir = nilai vendor (UPL)

        # Buat DataFrame vendor
        df_vendor = df_sub[info_cols].copy()
        df_vendor[name] = df_sub[value_col]

        # Simpan urutan baris dari sheet pertama sebagai referensi
        if i == 0:
            df_vendor["__order"] = range(len(df_vendor))
            merged = df_vendor
        else:
            merged = merged.merge(df_vendor, on=list(info_cols), how="outer")

    # Urutkan kembali sesuai urutan aslinya
    if "__order" in merged.columns:
        merged = merged.sort_values("__order").drop(columns="__order")

    merged_display = merged.copy()
    for col in merged_display.columns:
        if pd.api.types.is_numeric_dtype(merged_display[col]):
            merged_display[col] = merged_display[col].apply(format_thousand)

    total_bidders =  len(result)
    total_rows = len(merged)
    st.caption(f"Voilà! Here’s the summary of your analysis — **{total_bidders} bidders** and **{total_rows:,} total rows** analyzed!")

    # Simpan dan tampilkan
    st.session_state["merged_layer4"] = merged
    st.dataframe(merged_display, hide_index=True)

    # Download button to Excel
    @st.cache_data
    def get_excel_download(merged, sheet_name="Merged Data UPL"):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            merged.to_excel(writer, index=False, sheet_name=sheet_name)
        return output.getvalue()
    
    # Simpan hasil ke variabel
    excel_data = get_excel_download(merged)

    # Layout tombol (rata kanan)
    col1, col2, col3 = st.columns([3,1,1])
    with col3:
        st.download_button(
            label="Download",
            data=excel_data,
            file_name="Merged_UPL.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            icon=":material/download:",
        )

    st.divider()  

    # LOWEST PRICE & GAP (%) 
    st.subheader("📑 Lowest Price & Gap (%)")  

    # --- Ambil DataFrame hasil merged sebelumnya ---
    df_merged = st.session_state.get("merged_layer4", None)

    if df_merged is not None:
        df_pivot = df_merged.copy()

        # --- Identifikasi kolom vendor (semua setelah kolom info) ---
        # Asumsi: kolom vendor adalah semua kolom numeric atau non-template
        # Jika kamu tahu persis kolom info (misal ['Scope', 'Category']), bisa ganti di sini
        info_cols = df_pivot.columns[: -len(result)] if len(result) > 0 else []
        vendor_cols = [c for c in df_pivot.columns if c not in info_cols]

        # --- Hitung 1st & 2nd lowest ---
        def first_second(row):
            s = row[vendor_cols].dropna()
            if len(s) == 0:
                return np.nan, np.nan, np.nan, np.nan
            elif len(s) == 1:
                return s.iloc[0], s.index[0], np.nan, np.nan
            # Sort values karena bisa campuran dtype
            ns = s.sort_values()
            return ns.iloc[0], ns.index[0], ns.iloc[1], ns.index[1]

        res = df_pivot.apply(first_second, axis=1)
        df_pivot["1st Lowest"] = res.apply(lambda x: x[0])
        df_pivot["1st Vendor"] = res.apply(lambda x: x[1])
        df_pivot["2nd Lowest"] = res.apply(lambda x: x[2])
        df_pivot["2nd Vendor"] = res.apply(lambda x: x[3])

        # --- Hitung Gap 1 to 2 (%) ---
        def calc_gap(row):
            a, b = row["1st Lowest"], row["2nd Lowest"]
            if pd.isna(a) or pd.isna(b) or a == 0:
                return ""
            return f"{int(round((b - a) / a * 100, 0))}%"

        df_pivot["Gap 1 to 2 (%)"] = df_pivot.apply(calc_gap, axis=1)

        # --- Fungsi pembulatan dan format ribuan ---
        def round_half_up(n):
            if pd.isna(n):
                return n
            return int(Decimal(n).quantize(0, rounding=ROUND_HALF_UP))

        for c in vendor_cols + ["1st Lowest", "2nd Lowest"]:
            df_pivot[c] = df_pivot[c].apply(
                lambda x: f"{round_half_up(x):,}".replace(",", ".") if pd.notna(x) else ""
            )

        # --- Urutkan kolom (info → vendor → hasil analisis) ---
        ordered_cols = list(info_cols) + vendor_cols + [
            "1st Lowest",
            "1st Vendor",
            "2nd Lowest",
            "2nd Vendor",
            "Gap 1 to 2 (%)"
        ]
        df_pivot = df_pivot[ordered_cols]

        # --- 🎯 Tambahkan dua slicer terpisah untuk 1st Vendor dan 2nd Vendor
        all_1st = sorted(df_pivot["1st Vendor"].dropna().unique())
        all_2nd = sorted(df_pivot["2nd Vendor"].dropna().unique())

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
            df_filtered = df_pivot[
                df_pivot["1st Vendor"].isin(selected_1st) &
                df_pivot["2nd Vendor"].isin(selected_2nd)
            ]
        elif selected_1st:
            df_filtered = df_pivot[df_pivot["1st Vendor"].isin(selected_1st)]
        elif selected_2nd:
            df_filtered = df_pivot[df_pivot["2nd Vendor"].isin(selected_2nd)]
        else:
            df_filtered = df_pivot.copy()

        # Highlight pemenang
        def highlight_winners(row):
            styles = [""] * len(row)
            for i, col in enumerate(df_pivot.columns):
                if col == row["1st Vendor"]:
                    styles[i] = "background-color: #d7c6f3; color: #402e72;"
                elif col == row["2nd Vendor"]:
                    styles[i] = "background-color: #f8c8dc; color: #7a1f47;"
            return styles

        styled_df = df_filtered.style.apply(highlight_winners, axis=1)
        st.caption(f"✨ Total number of data entries: **{len(df_filtered)}**")
        st.dataframe(styled_df, hide_index=True)

        # Simpan hasil analisis ke session_state
        st.session_state["price_analysis"] = df_pivot

        # --- Buat salinan numerik untuk ekspor ---
        df_export = df_filtered.copy()

        # Ubah kolom angka dari format teks ke numerik kembali
        numeric_cols = vendor_cols + ["1st Lowest", "2nd Lowest"]
        for c in numeric_cols:
            df_export[c] = (
                df_export[c]
                .replace({r"\.": "", r",": "."}, regex=True)   # hilangkan pemisah ribuan, ubah koma jadi titik jika ada
                .replace("", np.nan)
                .astype(float)
            )

        # Download button to Excel
        @st.cache_data
        def get_excel_download(df_export, sheet_name="Lowest Price & Gap (%)"):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_export.to_excel(writer, index=False, sheet_name=sheet_name)
            return output.getvalue()

        # Simpan hasil ke variabel
        excel_data = get_excel_download(df_export, sheet_name="Lowest Price & Gap (%)")

        # Layout tombol (rata kanan)
        col1, col2, col3 = st.columns([3,1,1])
        with col3:
            st.download_button(
                label="Download",
                data=excel_data,
                file_name="Lowest Price & Gap (%).xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                icon=":material/download:",
                key=f"download_UPL_Maker"  # unik per tab
            )
        
        # --- WIN RATE VISUALIZATION ---
        st.markdown("##### Win Rate Trend")
        st.caption("Visualizing 1st and 2nd place win rates across all vendors to assess competitiveness.")

        # --- Hitung total kemenangan (1st & 2nd Vendor)
        win1_counts = df_pivot["1st Vendor"].value_counts(dropna=True).reset_index()
        win1_counts.columns = ["Vendor", "Wins1"]

        win2_counts = df_pivot["2nd Vendor"].value_counts(dropna=True).reset_index()
        win2_counts.columns = ["Vendor", "Wins2"]

        # --- Hitung total partisipasi vendor ---
        vendor_counts = (
            df_pivot[vendor_cols]
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
            title="📈 Vendor Win Rate Comparison (1st vs 2nd Place)"
        ).configure_title(
            anchor="middle",
            offset=12
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
            
            # Download button to Excel
            @st.cache_data
            def get_excel_download(df_summary, sheet_name="Win Rate Trend Summary"):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_summary.to_excel(writer, index=False, sheet_name=sheet_name)
                return output.getvalue()

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

        st.write("")

        # --- AVERAGE GAP VISUALIZATION ---
        st.markdown("##### Average Gap Trend")
        st.caption("Visualizing trend of the average gap between 1st and 2nd lowest bids.")

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
            offset=12
        ).configure_axis(
            labelFontSize=12,
            titleFontSize=13
        ).configure_view(
            stroke='gray',
            strokeWidth=1
        )

        # --- Tampilkan di Streamlit ---
        st.altair_chart(chart)

        avg_value = avg_gap["Average Gap (%)"].mean()
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

        # ---------- MEDIAN ---------- 
        st.subheader("🎯 Median Price & Vendor Deviation")
        st.caption("Shows each vendor’s percentage deviation from the median price — lower means more competitive.")

        # --- Hitung median per baris (berdasarkan vendor columns) ---
        df_median = df_pivot.copy()

        # Pastikan kolom vendor numeric dulu sebelum hitung median
        for v in vendor_cols:
            df_median[v] = pd.to_numeric(df_median[v].replace("", np.nan).astype(str).str.replace(".", ""), errors="coerce")

        # Hitung median tiap baris (berdasarkan vendor)
        df_median["Median"] = df_median[vendor_cols].median(axis=1, skipna=True)

        # Hitung gap tiap vendor terhadap median (%)
        for v in vendor_cols:
            df_median[f"{v} to Median (%)"] = df_median.apply(
                lambda row: (
                    f"{((row[v] - row['Median']) / row['Median'] * 100):.1f}%"
                    if pd.notna(row[v]) and pd.notna(row['Median']) and row['Median'] != 0
                    else ""
                ),
                axis=1
            )

        # Ambil hanya kolom yang diperlukan
        cols_final = list(info_cols) + ["Median"] + [f"{v} to Median (%)" for v in vendor_cols]
        df_median = df_median[cols_final]

        df_export_median = df_median.copy()

        # Optional: format Median jadi ribuan
        df_median["Median"] = df_median["Median"].apply(
            lambda x: f"{int(round(x)):,}".replace(",", ".") if pd.notna(x) else ""
        )

        # --- Fungsi untuk meng-highlight vendor dengan gap (%) terendah ---
        def highlight_lowest_median(s):
            # ambil hanya kolom vendor yang berisi persentase ke median
            vendor_cols_pct = [c for c in s.index if c.endswith("to Median (%)")]

            # konversi ke float (abaikan % dan kosong)
            vals = s[vendor_cols_pct].replace("", np.nan).str.replace("%", "").astype(float)

            # cari nilai minimum
            if vals.notna().any():
                min_val = vals.min()
            else:
                min_val = None

            # siapkan style per kolom
            styles = []
            for c in s.index:
                if c in vendor_cols_pct and pd.notna(min_val) and float(str(s[c]).replace("%", "")) == min_val:
                    styles.append("background-color: #f8c8dc; color: #7a1f47;")
                else:
                    styles.append("")
            return styles

        # --- Terapkan styling ---
        df_median_styled = df_median.style.apply(highlight_lowest_median, axis=1)

        st.dataframe(df_median_styled, hide_index=True)

        # Download button to Excel
        @st.cache_data
        def get_excel_download(df_export_median, sheet_name="Median Analysis (%)"):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_export_median.to_excel(writer, index=False, sheet_name=sheet_name)
            return output.getvalue()

        # Simpan hasil ke variabel
        excel_data = get_excel_download(df_export_median, sheet_name=f"Median Analysis (%)")

        # Layout tombol (rata kanan)
        col1, col2, col3 = st.columns([3,1,1])
        with col3:
            st.download_button(
                label="Download",
                data=excel_data,
                file_name=f"Median Analysis (%).xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                icon=":material/download:",
                key=f"download_median"
            )
