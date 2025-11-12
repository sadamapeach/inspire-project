import streamlit as st
import pages.home as home 
# import pages.layer1 as layer1
import pages.TCO_by_Year as TCO_by_Year
import pages.TCO_by_Region as TCO_by_Region
import pages.TCO_by_Year_Region as TCO_by_Year_Region
import pages.UPL_Comparison as UPL_Comparison
# import pages.layer3 as layer3
# import pages.layer5 as layer5

pg = st.navigation([
    st.Page(home.page, title="🏡 Home", url_path="home"),
    # st.Page(layer1.page, title="1️⃣ Single Layer TCO", url_path="layer1"),
    st.Page(TCO_by_Year.page, title="1️⃣ TCO Comparison by Year", url_path="1_TCO_Comparison_by_Year"),
    st.Page(TCO_by_Region.page, title="2️⃣ TCO Comparison by Region", url_path="2_TCO_Comparison_by_Region"),
    st.Page(TCO_by_Year_Region.page, title="3️⃣ TCO Comparison by Year + Region", url_path="3_TCO_Comparison_by_Year_Region"),
    st.Page(UPL_Comparison.page, title="4️⃣ UPL Comparison", url_path="4_UPL_Comparison"),
    # st.Page(layer3.page, title="3️⃣ Standard Deviation", url_path="layer3"),
    # st.Page(layer5.page, title="5️⃣ Round UPL", url_path="layer5"),
])

pg.run()
