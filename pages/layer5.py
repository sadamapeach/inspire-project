import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
import numbers
import time
import math
import re
from io import BytesIO
from decimal import Decimal, ROUND_HALF_UP

def page():
    # Header Title
    st.title("5️⃣ Round UPL")
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
        st.session_state["uploaded_file_layer5"] = upload_file

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
        st.session_state["df_layer5_raw"] = df

    elif "df_layer5_raw" in st.session_state:
        df = st.session_state["df_layer5_raw"]
    else:
        return
    
    st.divider()

    # OVERVIEW
    st.subheader("🔍 Overview")

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

        # Tampilkan di Streamlit
        st.caption(f"Your dataset contains **{len(df)} rows** ready for analysis!")
        st.dataframe(df, hide_index=True)

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
        # Tampilkan di Streamlit
        st.caption(f"Data loaded! Your dataset contains **{len(df)} rows** ready for analysis 🤓")
        st.dataframe(df, hide_index=True)

        # st.dataframe(df, hide_index=True)
        st.markdown(
            """
            <p style='font-size:12px; color:#808080; margin-top:-15px; margin-bottom:0;'>
                No hidden rows and columns detected. Proceeding with raw data ✅
            </p>
            """,
            unsafe_allow_html=True
        )

    st.write("")

    st.subheader("🧮 Round: Lowest Price & Gap (%)")

    col_vendor = df.columns[0]
    col_round  = df.columns[1]
    col_scope  = df.columns[2]
    col_price  = df.columns[3]

    # --- Tab per round
    rounds = df[col_round].unique()
    tabs = st.tabs([f"{r}" for r in rounds])

    # Inisialisasi list untuk simpan hasil semua round
    df_all_rounds = []

    for i, r in enumerate(rounds):
        with tabs[i]:
            df_r = df[df[col_round] == r].copy().reset_index(drop=True)

            # --- Tambah kolom order untuk handle duplicate Scope per vendor
            df_r["__order"] = df_r.groupby([col_scope, col_vendor]).cumcount()

            # --- Pivot horizontal: Scope di baris, Vendor di kolom
            df_pivot = df_r.pivot_table(
                index=[col_scope, "__order"],
                columns=col_vendor,
                values=col_price,
                aggfunc="first"
            ).reset_index()

            # Vendor columns
            vendor_cols = [c for c in df_pivot.columns if c not in [col_scope, "__order"]]

            # --- Pastikan semua numeric, coerce errors
            df_pivot[vendor_cols] = df_pivot[vendor_cols].apply(pd.to_numeric, errors='coerce')

            # --- Hitung 1st & 2nd lowest
            def first_second(row):
                s = row[vendor_cols].dropna()
                if len(s) == 0:
                    return np.nan, np.nan, np.nan, np.nan
                elif len(s) == 1:
                    return s.iloc[0], s.index[0], np.nan, np.nan
                # Gunakan sort_values karena nsmallest kadang error jika dtype object
                ns = s.sort_values().iloc[:2]
                return ns.iloc[0], ns.index[0], ns.iloc[1], ns.index[1]

            res = df_pivot.apply(first_second, axis=1)
            df_pivot["1st Lowest"] = res.apply(lambda x: x[0])
            df_pivot["1st Vendor"] = res.apply(lambda x: x[1])
            df_pivot["2nd Lowest"] = res.apply(lambda x: x[2])
            df_pivot["2nd Vendor"] = res.apply(lambda x: x[3])

            # --- Hitung Gap 1 to 2 (%)
            def calc_gap(row):
                a, b = row["1st Lowest"], row["2nd Lowest"]
                if pd.isna(a) or pd.isna(b) or a == 0:
                    return ""
                return f"{int(round((b-a)/a*100,0))}%"

            df_pivot["Gap 1 to 2 (%)"] = df_pivot.apply(calc_gap, axis=1)

            def round_half_up(n):
                if pd.isna(n):
                    return n
                return int(Decimal(n).quantize(0, rounding=ROUND_HALF_UP))

            for c in vendor_cols + ["1st Lowest","2nd Lowest"]:
                df_pivot[c] = df_pivot[c].apply(
                    lambda x: f"{round_half_up(x):,}".replace(",", ".") if pd.notna(x) else ""
                )

            # --- Kembalikan urutan & hapus kolom pembantu
            df_pivot = df_pivot.sort_values(["__order", col_scope]).drop(columns="__order").reset_index(drop=True)

            # --- Urutkan kolom akhir
            ordered_cols = [col_scope] + vendor_cols + ["1st Lowest","1st Vendor","2nd Lowest","2nd Vendor","Gap 1 to 2 (%)"]
            df_pivot = df_pivot[ordered_cols]

            # Urutkan scope sesuai urutan kemunculan di df_r
            scope_order = df_r[col_scope].drop_duplicates()
            df_pivot = df_pivot.set_index(col_scope).loc[scope_order].reset_index()

            # LINE CHART (simpan semua hasil df_pivot + info round)
            df_line_chart = df_pivot.copy()
            df_line_chart["Round"] = r
            df_all_rounds.append(df_line_chart)

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
                    key=f"filter_1st_vendor_{r}"
                )
            with col_sel_2:
                selected_2nd = st.multiselect(
                    "Filter: 2nd vendor",
                    options=all_2nd,
                    default=None,
                    placeholder="Choose one or more vendors",
                    key=f"filter_2nd_vendor_{r}"
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

            # --- Styling function untuk highlight
            def highlight_winners(row):
                styles = [""] * len(row)
                if "1st Vendor" in row and "2nd Vendor" in row:
                    for i, col in enumerate(df_pivot.columns):
                        if col == row["1st Vendor"]:
                            styles[i] = "background-color: #d7c6f3; color: #402e72; font-weight: bold;"
                            # styles[i] = "background-color: #b7e4c7; color: #1b4332; font-weight: bold;"
                        elif col == row["2nd Vendor"]:
                            styles[i] = "background-color: #f8c8dc; color: #7a1f47; font-weight: bold;"
                            # styles[i] = "background-color: #fff3b0; color: #665c00; font-weight: bold;"
                return styles

            # --- Terapkan styling (setelah df_pivot final)
            styled_df = df_filtered.style.apply(highlight_winners, axis=1)

            st.caption(f"✨ Total number of data entries: **{len(df_filtered)}**")
            st.dataframe(styled_df, hide_index=True)

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
            excel_data = get_excel_download(df_export, sheet_name=f"Round_{r}")

            # Layout tombol (rata kanan)
            col1, col2, col3 = st.columns([3,1,1])
            with col3:
                st.download_button(
                    label="Download",
                    data=excel_data,
                    file_name=f"Lowest Price & Gap ({r}).xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    icon=":material/download:",
                    key=f"download_{r}"  # unik per tab
                )

    st.write("")

    # TABEL JUMLAH KEMENANGAN
    st.markdown("##### Trend of Vendor Wins Across Rounds")
    st.caption("This line chart visualization shows how vendor wins change across bidding rounds and highlights consistent leaders.")

    # Line Chart Visualization
    # Gabungkan semua round
    df_all = pd.concat(df_all_rounds, ignore_index=True)

    # --- Normalisasi nama vendor (biar konsisten) ---
    df_all["1st Vendor"] = (
        df_all["1st Vendor"]
        .astype(str)
        .str.strip()      # hilangkan spasi depan/belakang
        .str.upper()      # ubah ke huruf besar semua
    )

    # --- Hitung jumlah kemenangan per vendor per round ---
    win_summary = (
        df_all.groupby(["Round", "1st Vendor"])
            .size()
            .reset_index(name="Wins")
            .rename(columns={"1st Vendor": "Vendor"})
    )

    # --- Urutkan round ---
    round_order = sorted(df_all["Round"].unique(), key=lambda x: str(x))
    win_summary["Round"] = pd.Categorical(win_summary["Round"], categories=round_order, ordered=True)

    # --- Tambahan: pastikan kombinasi Round–Vendor yang hilang diisi 0 ---
    all_rounds = win_summary["Round"].unique()
    all_vendors = win_summary["Vendor"].unique()

    # Buat semua kombinasi round–vendor
    full_index = pd.MultiIndex.from_product([all_rounds, all_vendors], names=["Round", "Vendor"])
    win_summary = (
        win_summary.set_index(["Round", "Vendor"])
        .reindex(full_index, fill_value=0)
        .reset_index()
    )

    # --- Buat chart dengan Altair ---
    y_min = win_summary["Wins"].min()
    y_max = win_summary["Wins"].max()

    # --- Hitung total kemenangan per vendor ---
    vendor_order = (
        win_summary.groupby("Vendor")["Wins"]
        .sum()
        .sort_values(ascending=False)
        .index.tolist()
    )

    chart = (
        alt.Chart(win_summary)
        .mark_line(point=True)
        .encode(
            x=alt.X("Round:N", sort=round_order, title="Round"),
            y=alt.Y(
                "Wins:Q", 
                title="Number of Wins",
                scale=alt.Scale(domain=[y_min - 1.5, y_max + 1.5]),
                axis=alt.Axis(
                    tickMinStep=1,
                    tickCount=win_summary["Wins"].nunique() + 1
                )
            ),
            color=alt.Color("Vendor:N", title="Vendor", sort=vendor_order)
        )
        .properties(
            height=400,
            width="container",
            title="Winning Performance Across Rounds"
        ).configure_title(
            anchor='middle', 
            offset=12
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
            index="Vendor",
            columns="Round",
            values="Wins",
            aggfunc="sum",
            fill_value=0,
            observed=False
        ).reset_index()
    )

    # Urutkan vendor berdasarkan total kemenangan
    win_table["Total Wins"] = win_table.drop(columns="Vendor").sum(axis=1)
    win_table = win_table.sort_values("Total Wins", ascending=False).reset_index(drop=True)

    # st.dataframe(win_table, hide_index=True)
    st.altair_chart(chart)
    st.dataframe(win_table, hide_index=True)

    # Price Movement
    st.write(" ")
    st.markdown("##### Trend of Price Movement per Scope")

    col_vendor = df.columns[0]
    col_round  = df.columns[1]
    col_scope  = df.columns[2]
    col_price  = df.columns[3]

    # --- Normalisasi nama vendor biar konsisten ---
    df[col_vendor] = (
        df[col_vendor]
        .astype(str)
        .str.strip()      # Hilangkan spasi depan/belakang
        .str.upper()      # Ubah ke huruf besar semua
    )

    # --- Tab per vendor
    vendors = df[col_vendor].unique()
    tabs = st.tabs([f"{v}" for v in vendors])

    # Inisialisasi list untuk simpan hasil semua vendor
    df_all_vendors = []

    for i, v in enumerate(vendors):
        with tabs[i]:
            # Filter data vendor
            df_vendor = df[df[col_vendor] == v].copy().reset_index(drop=True)

            # Tambahkan kolom order untuk handle duplicate Scope
            df_vendor["__order"] = df_vendor.groupby([col_scope, col_round]).cumcount()

            # Pivot: scope sebagai index, round sebagai kolom, value = UPL
            df_pivot = (
                df_vendor
                .pivot_table(
                    index=[col_scope, "__order"],
                    columns=col_round,
                    values=col_price,
                    aggfunc="first",
                    sort=False
                )
                .fillna(0)
                .reset_index()
            )

            # Pembulatan
            def round_half_up(n):
                if pd.isna(n):
                    return n
                try:
                    return int(Decimal(str(n)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
                except:
                    return n
            
            # Terapkan pembulatan
            num_cols = [c for c in df_pivot.columns if c not in [col_scope, "__order"]]
            for c in num_cols:
                df_pivot[c] = df_pivot[c].apply(lambda x: round_half_up(x) if pd.notna(x) else x)

            # Tambahkan kolom chart (list dari tiap baris)
            # df_pivot["Price Trend"] = df_pivot[[c for c in num_cols if c != "__order"]].values.tolist()
            df_pivot["Price Trend"] = (
                df_pivot[[c for c in num_cols if c != "__order"]]
                .apply(lambda x: str(list(x)), axis=1)
            )

            # Ambil hanya kolom numerik (rounds) untuk hitung min/max
            numeric_cols = df_pivot.select_dtypes(include="number").columns
            if len(numeric_cols) > 0:
                y_min = int(df_pivot[numeric_cols].min().min())
                y_max = int(df_pivot[numeric_cols].max().max())
            else:
                y_min, y_max = 0, 0

            # Hapus kolom pembantu dan urutkan ulang
            df_pivot = df_pivot.sort_values(["__order", col_scope]).drop(columns="__order").reset_index(drop=True)

            # Urutkan scope sesuai kemunculan awal
            scope_order = df_vendor[col_scope].drop_duplicates()
            df_pivot = df_pivot.set_index(col_scope).loc[scope_order].reset_index()

            # MENAMBAHKAN SLICER
            all_scope = sorted(df_pivot[col_scope].dropna().unique())
            selected_scope = st.multiselect(
                f"Filter: {col_scope}",
                options=all_scope,
                default=None,
                placeholder="Choose one or more scope",
                key=f"scope_filter_{v}"
            )

            # --- Terapkan filter kalau user memilih scope
            if selected_scope:
                df_filtered = df_pivot[df_pivot[col_scope].isin(selected_scope)]
            else:
                df_filtered = df_pivot  # tampilkan semua kalau belum difilter

            df_styled = df_filtered.copy()
            def format_thousand(value):
                """Format angka ke format Indonesia (1.000.000), hilangkan 0 jadi kosong."""
                try:
                    if pd.isna(value) or value == 0:
                        return ""  # kosongkan nilai 0 atau NaN
                    return f"{int(value):,}".replace(",", ".")
                except (ValueError, TypeError):
                    return value
                
            # Pilih kolom round (semua numerik kecuali kolom scope dan Price Trend)
            round_cols = [c for c in df_styled.columns if c not in [col_scope, "Price Trend"]]

            # Terapkan format ribuan
            for c in round_cols:
                df_styled[c] = df_styled[c].apply(format_thousand)

            # Fungsi styling untuk highlight pink pada nilai kosong (bekas 0)
            def highlight_zero(val):
                if val == "":
                    return "background-color: #f8c8dc"
                return ""
            
            # Terapkan styling ke kolom round saja
            df_styled = df_styled.style.map(highlight_zero, subset=round_cols)

            # --- Tampilkan tabel dengan konfigurasi kolom Streamlit ---
            st.caption(f"✨ Total number of data entries: **{len(df_filtered)}**")
            st.dataframe(
                df_styled,
                hide_index=True,
                column_config={
                    col_scope: "Scope",
                    "Price Trend": st.column_config.LineChartColumn(
                        "Price Trend",
                        help="Shows price changes across rounds",
                        y_min=y_min,
                        y_max=y_max,
                    ),
                }
            )

            # Simpan hasil (opsional)
            df_all_vendors.append(df_pivot)

            # Download button to Excel
            @st.cache_data
            def get_excel_download(df_filtered, sheet_name="Trend of Price Movement per Scope"):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_filtered.to_excel(writer, index=False, sheet_name=sheet_name)
                return output.getvalue()
            
            # Simpan hasil ke variabel
            excel_data = get_excel_download(df_filtered, sheet_name=f"{v}")

            # Layout
            col1, col2, col3 = st.columns([3,1,1])
            with col3:
                st.download_button(
                    label="Download",
                    data=excel_data,
                    file_name=f"Trend of Price Movement per Scope ({v}).xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    icon=":material/download:",
                    key=f"download_{v}"  # unik per tab
                )

    # MEDIAN
    st.subheader("🎯 Round: Bidder to Median (%)")

    col_vendor = df.columns[0]
    col_round  = df.columns[1]
    col_scope  = df.columns[2]
    col_price  = df.columns[3]

    # --- Tab per round
    rounds = df[col_round].unique()
    tabs = st.tabs([f"{r}" for r in rounds])

    for i, r in enumerate(rounds):
        with tabs[i]:
            df_r = df[df[col_round] == r].copy().reset_index(drop=True)

            # --- Tambah kolom order untuk handle duplicate Scope per vendor
            df_r["__order"] = df_r.groupby([col_scope, col_vendor]).cumcount()

            # --- Pivot horizontal: Scope di baris, Vendor di kolom
            df_pivot = df_r.pivot_table(
                index=[col_scope, "__order"],
                columns=col_vendor,
                values=col_price,
                aggfunc="first"
            ).reset_index()

            vendor_cols = [c for c in df_pivot.columns if c not in [col_scope, "__order"]]

            # Pastikan numeric
            df_pivot[vendor_cols] = df_pivot[vendor_cols].apply(pd.to_numeric, errors='coerce')

            # --- Hitung median per baris
            df_pivot["Median"] = df_pivot[vendor_cols].median(axis=1).round().astype("Int64")

            # --- Fungsi pembulatan (half up)
            def round_half_up(n):
                if pd.isna(n):
                    return n
                return float(Decimal(n).quantize(Decimal("0.1"), rounding=ROUND_HALF_UP))

            # --- Hitung deviasi tiap vendor terhadap median
            for v in vendor_cols:
                df_pivot[f"{v} to Median (%)"] = df_pivot.apply(
                    lambda row: (
                        f"{round_half_up(((row[v] - row['Median']) / row['Median']) * 100)}%"
                        if pd.notna(row[v]) and pd.notna(row['Median']) and row['Median'] != 0
                        else ""
                    ),
                    axis=1
                )

            # --- Urutkan & rapikan kolom
            ordered_cols = [col_scope] + ["Median"] + [f"{v} to Median (%)" for v in vendor_cols]
            df_pivot = df_pivot[ordered_cols]

            # --- Urutkan scope sesuai urutan kemunculan di df_r
            scope_order = df_r[col_scope].drop_duplicates()
            df_pivot = df_pivot.set_index(col_scope).loc[scope_order].reset_index()

            # Buat df_export2 untuk donwload
            df_export2 = df_pivot.copy()

            # Pemisah ribuan titik
            df_pivot["Median"] = df_pivot["Median"].apply(
                lambda x: f"{x:,.0f}".replace(",", ".") if pd.notna(x) else ""
            )

            # --- Highlight vendor dengan deviasi terkecil (paling negatif)
            def highlight_lowest(row):
                styles = [""] * len(row)
                values = [
                    float(str(row[f"{v} to Median (%)"]).replace("%", "")) 
                    if str(row[f"{v} to Median (%)"]).replace("%", "").strip() != "" 
                    else np.nan
                    for v in vendor_cols
                ]
                if all(np.isnan(values)):
                    return styles
                min_val = np.nanmin(values)
                for idx, val in enumerate(values):
                    if val == min_val:
                        styles[idx + 2] = "background-color: #f8c8dc; color: #7a1f47; font-weight: bold;"
                return styles

            styled_df = df_pivot.style.apply(highlight_lowest, axis=1)

            st.caption(f"Round **{r}** contains **{len(df_pivot)} scopes** for median analysis!")
            st.dataframe(styled_df, hide_index=True)

            # Download button to Excel
            @st.cache_data
            def get_excel_download(df_export2, sheet_name="Median Analysis (%)"):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_export2.to_excel(writer, index=False, sheet_name=sheet_name)
                return output.getvalue()

            # Simpan hasil ke variabel
            excel_data = get_excel_download(df_export2, sheet_name=f"Round_{r}")

            # Layout tombol (rata kanan)
            col1, col2, col3 = st.columns([3,1,1])
            with col3:
                st.download_button(
                    label="Download",
                    data=excel_data,
                    file_name=f"Median Analysis (%) ({r}).xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    icon=":material/download:",
                    key=f"download_med_{r}"  # unik per tab
                )