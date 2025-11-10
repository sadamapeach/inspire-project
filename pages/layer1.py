# import streamlit as st
# import pandas as pd
# import numpy as np
# import altair as alt

# def page():
#     # Header Title
#     st.title("1️⃣ Single Layer TCO")
#     st.markdown(
#         ":red-badge[Indosat] :orange-badge[Ooredoo] :green-badge[Hutchison]"
#     )
#     st.caption("Hai Team! Drop in your annual pricing template and let this analytics system work its magic ✨")

#     # Divider custom
#     st.markdown(
#         """
#         <hr style="margin-top:-5px; margin-bottom:10px; border: none; height: 2px; background-color: #ddd;">
#         """,
#         unsafe_allow_html=True
#     )

#     # File Uploader
#     st.subheader("🗃️ Upload File")
#     upload_file = st.file_uploader("Upload your file here!", type=["xlsx", "xls"])

#     if upload_file is not None:
#         st.session_state["uploaded_file_layer1"] = upload_file
#         st.success('File uploaded successfully!', icon="✅")

#         # Baca semua sheet sekaligus
#         df = pd.read_excel(upload_file)
#         st.session_state["df_layer1"] = df

#         # result = {}
#         # st.divider()

#     elif "df_layer1" in st.session_state:
#         df = st.session_state["df_layer1"]
#     else:
#         return
    
#     st.divider()

#     # OVERVIEW
#     st.subheader("🗃️ Overview")
#     # Total bidders
#     total_bidders = len(df.columns) - 1     # minus TCO Component (indeks[0])
#     st.caption(f"There are {total_bidders} bidders competing 🤯")
#     st.dataframe(df, hide_index=True)

#     st.divider()

#     # RANK
#     cols_to_sum = df.columns[1:]
#     sum_series = df[cols_to_sum].sum(numeric_only=True)
#     rank_sum = sum_series.sort_values(ascending=True)

#     # Tabel
#     st.subheader("🗃️ Bidder's Rank")
#     st.caption("Hooray! You've got the bidder with the lowest offer 🎉")

#     # Siapkan data dengan ranking
#     df_chart = (
#         rank_sum.reset_index()
#         .rename(columns={"index": "Vendor", 0: "Total"})
#         .sort_values("Total", ascending=True)
#     )
#     df_chart["Rank"] = range(1, len(df_chart) + 1)
#     df_chart["Mid"] = df_chart["Total"] / 2     # Tambahkan kolom posisi tengah bar

#     # String format dengan titik ribuan
#     df_chart["Total_str"] = df_chart["Total"].apply(lambda x: f"{int(x):,}".replace(",", "."))
#     # df_chart["Legend"] = df_chart["Total_str"] + " USD"
#     df_chart["Legend"] = df_chart.apply(
#         lambda x: f"Rank {x['Rank']} - {x['Total_str']} USD", axis=1
#     )

#     # Warna manual tetap sama
#     vendor_colors = {
#         v: c for v, c in zip(df_chart["Legend"], [
#             # "#32AAB5", "#E7D39A", "#F1AA60", "#F27B68", "#E54787", "#BF219A", "#8E0F9C", "#4B1D91"
#             # "#E37083", "#FFB876", "#F49AA2", "#A8BF8A", "#FFCB7C", "#89B7C2"
#             "#F94144", "#F3722C", "#F8961E", "#F9C74F", "#90BE6D", "#43AA8B", "#577590", "#E54787", "#BF219A", "#8E0F9C", "#4B1D91"
#         ])
#     }

#     highlight = alt.selection_point(on='mouseover', fields=['Vendor'], nearest=True)

#     # Chart bar dengan warna berbeda per vendor
#     bars = (
#         alt.Chart(df_chart)
#         .mark_bar()
#         .encode(
#             x=alt.X("Vendor:N", axis=alt.Axis(title=None)),
#             y=alt.Y(
#                 "Total:Q", 
#                 axis=alt.Axis(title=None, grid=False),
#                 scale=alt.Scale(domain=[0, df_chart["Total"].max() * 1.1])
#             ),
#             color=alt.Color(
#                 "Legend:N", 
#                 title="Total Offer by Rank",
#                 scale=alt.Scale(domain=list(vendor_colors.keys()),
#                                 range=list(vendor_colors.values()))  
#             ),
#             tooltip=[
#                 alt.Tooltip("Vendor:N", title="Vendor"),
#                 alt.Tooltip("Total_str:N", title="Total (USD)")
#             ]
#         ).add_params(highlight)
#     )

#     # Rank di dalam bar (tengah)
#     rank_text = (
#         alt.Chart(df_chart)
#         .mark_text(color="white", fontWeight="bold")
#         .encode(
#             x="Vendor:N",
#             y="Mid:Q",       # pakai posisi tengah
#             text="Rank:N"
#         )
#     )

#     # Tambahkan border penuh (pakai layer persegi)
#     frame = (
#         alt.Chart()
#         .mark_rect(
#             stroke='gray',
#             strokeWidth=1,
#             fillOpacity=0
#         )
#     )

#     # Gabungkan semua layer
#     chart = (bars + rank_text + frame).properties(
#         title="✨ Bidder's Rank and Total Offer ✨"
#     ).configure_title(
#         anchor='middle',
#         offset=15
#     ).configure_legend(
#         titleFontSize=12,        
#         titleFontWeight="bold",  
#         labelFontSize=12,        
#         labelLimit=300,          
#         # labelFont="Montserrat",
#         orient="right"
#     )

#     st.altair_chart(chart, use_container_width=True)

#     st.divider()

#     # COMPARISON
#     reqs = df[df.columns[0]]
#     vendors = df.columns[1:]

#     st.subheader("🗃️ Comparison of Components")
#     st.caption("Here’s a comparison of each component to help you make better decisions. Good luck!")

#     n_cols = 2
#     for i in range(0, len(reqs), n_cols):
#         cols = st.columns(n_cols)

#         for j in range(n_cols):
#             if i + j < len(reqs):
#                 req = reqs[i + j]

#                 # Data vendor + harga
#                 prices = df.loc[i + j, vendors].reset_index()
#                 prices.columns = ["Vendor", "Total"]

#                 # Urutkan harga + tambahkan rank
#                 df_chart = prices.sort_values("Total", ascending=True).reset_index(drop=True)
#                 df_chart["Rank"] = range(1, len(df_chart) + 1)
#                 df_chart["Mid"] = df_chart["Total"] / 2

#                 # String format dengan titik ribuan
#                 df_chart["Total_str"] = df_chart["Total"].apply(lambda x: f"{int(x):,}".replace(",", "."))
#                 # df_chart["Legend"] = df_chart["Total_str"] + " USD"
#                 df_chart["Legend"] = df_chart.apply(
#                     lambda x: f"R{x['Rank']} - {x['Total_str']} USD", axis=1
#                 )

#                 # Warna manual tetap sama
#                 vendor_colors = {
#                     v: c for v, c in zip(df_chart["Legend"], [
#                         # "#32AAB5", "#E7D39A", "#F1AA60", "#F27B68", "#E54787", "#BF219A", "#8E0F9C", "#4B1D91"
#                         # "#E37083", "#FFB876", "#F49AA2", "#A8BF8A", "#FFCB7C", "#89B7C2"
#                         "#F94144", "#F3722C", "#F8961E", "#F9C74F", "#90BE6D", "#43AA8B", "#577590", "#E54787", "#BF219A", "#8E0F9C", "#4B1D91"
#                     ])
#                 }

#                 highlight = alt.selection_point(on='mouseover', fields=['Vendor'], nearest=True)

#                 # Bar chart dengan warna berbeda tiap vendor
#                 bars = (
#                     alt.Chart(df_chart)
#                     .mark_bar()
#                     .encode(
#                         x=alt.X("Vendor:N", axis=alt.Axis(title=None)),
#                         y=alt.Y(
#                             "Total:Q",
#                             axis=alt.Axis(title=None, grid=False),
#                             # scale=alt.Scale(domain=[y_min, y_max])
#                             scale=alt.Scale(domain=[0, df_chart["Total"].max() * 1.1])
#                         ),
#                         color=alt.Color(
#                             "Legend:N", 
#                             title=f"Total {req} Offer by Rank",  # beda warna otomatis
#                             scale=alt.Scale(domain=list(vendor_colors.keys()),
#                                             range=list(vendor_colors.values()))
#                         ),
#                         tooltip=[
#                             alt.Tooltip("Vendor:N", title="Vendor"),
#                             alt.Tooltip("Total_str:N", title="Total (USD)")
#                         ]
#                     )
#                     .add_params(highlight)
#                 )

#                 # Rank di dalam bar
#                 rank_text = (
#                     alt.Chart(df_chart)
#                     .mark_text(color="white", fontWeight="bold")
#                     .encode(
#                         x="Vendor:N",
#                         y="Mid:Q",
#                         # y=alt.Y("Scaled_Total:Q", stack="zero"),
#                         text="Rank:N"
#                     )
#                 )

#                 # Tambahkan border penuh (pakai layer persegi)
#                 frame = (
#                     alt.Chart()
#                     .mark_rect(
#                         stroke='gray',
#                         strokeWidth=1,
#                         fillOpacity=0
#                     )
#                 )

#                 chart = (bars + rank_text + frame).properties(
#                     title=f"✨ Comparison - {req} ✨"
#                 ).configure_title(
#                     anchor='middle',
#                     offset=15
#                 ).configure_legend(
#                     orient="bottom",
#                     direction="horizontal",
#                     columns=2,
#                     titleFontSize=12,
#                     titleFontWeight="bold",
#                     labelFontSize=12,
#                     labelLimit=300
#                 )

#                 # Tampilkan chart di kolom
#                 cols[j].altair_chart(chart, use_container_width=True)

