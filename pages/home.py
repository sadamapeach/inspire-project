import streamlit as st
import pandas as pd

def page():
    st.header("🏡 Intro: Bid & Price Analytics Tool")
    st.markdown(
        ":red-badge[Indosat] :orange-badge[Ooredoo] :green-badge[Hutchison]"
    )
    st.caption("INSPIRE 2025 | Oktaviana Sadama Nur Azizah")

    # Divider custom
    st.markdown(
        """
        <hr style="margin-top:-5px; margin-bottom:10px; border: none; height: 2px; background-color: #ddd;">
        """,
        unsafe_allow_html=True
    )

    # st.markdown("On progress..")
    # st.image("https://i.pinimg.com/originals/61/8f/08/618f083c61a7460ce0a6064319af41bd.gif")

    st.markdown("""
        <div style="text-align: justify; font-size: 15px; margin-bottom: 30px">
            <span style="color: orange; font-weight: 600;"> Bid & Price Analytics Tool</span>
            is a <a href="https://streamlit.io" target="_blank" 
            style="color:#4DA3FF;">
            Streamlit</a>-powered application that automates
            commercial bid comparison by analyzing multi-vendor submissions, identifying the 
            most competitive offers, and generating clear, ready-to-share analytical outputs.
        </div>
    """, unsafe_allow_html=True)
    
    # Main Features
    # st.markdown("#### Main Features")
    st.subheader("Main Features")
    st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

    # 1️⃣ TCO Comparison by Year
    col1, col2 = st.columns([1, 9])  # kiri lebih sempit

    # background-color: #EC008C;
    with col1:
        st.markdown(
            """
            <div style="
                background-color: #BC13FE;
                border-radius: 8px;
                padding: 20px;
                width: 65px;
                height: 65px;
                display: flex;
                align-items: center;
                justify-content: center;
                box-shadow: 0 2px 5px rgba(0,0,0,0.15);
            ">
                <img src="https://cdn-icons-png.flaticon.com/512/14991/14991730.png" width="35" />
            </div>
            """,
            unsafe_allow_html=True
        )

    with col2:
        st.markdown(
            """
            <div style="
                display: flex;
                align-items: center;
                height: 65px;
            ">
                <div style="text-align: justify; font-size: 15px;">
                    <span style="color: #BC13FE; font-weight: 800;">TCO Comparison by Year</span>
                    compares TCO accross vendors by analyzing year-over-year price changes to
                    identify long-term cost trends.
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.divider()

    # 2️⃣ TCO Comparison by Region
    col1, col2 = st.columns([1, 9])  # kiri lebih sempit

    # background-color: #C6168D;
    with col1:
        st.markdown(
            """
            <div style="
                background-color: #FF2EC4; 
                border-radius: 8px;
                padding: 20px;
                width: 65px;
                height: 65px;
                display: flex;
                align-items: center;
                justify-content: center;
                box-shadow: 0 2px 5px rgba(0,0,0,0.15);
            ">
                <img src="https://cdn-icons-png.flaticon.com/512/15366/15366033.png" width="35" />
            </div>
            """,
            unsafe_allow_html=True
        )

    with col2:
        st.markdown(
            """
            <div style="
                display: flex;
                align-items: center;
                height: 65px;
            ">
                <div style="text-align: justify; font-size: 15px;">
                    <span style="color: #FF2EC4; font-weight: 800;">TCO Comparison by Region</span>
                    analyzes TCO differences across regions to highlight cost variations based on
                    geographic requirements.
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.divider()

    # 3️⃣ TCO Comparison by Year + Region
    col1, col2 = st.columns([1, 9])  # kiri lebih sempit

    # background-color: #ED1C24;
    with col1:
        st.markdown(
            """
            <div style="
                background-color: #FF1BF1; 
                border-radius: 8px;
                padding: 20px;
                width: 65px;
                height: 65px;
                display: flex;
                align-items: center;
                justify-content: center;
                box-shadow: 0 2px 5px rgba(0,0,0,0.15);
            ">
                <img src="https://cdn-icons-png.flaticon.com/512/4624/4624053.png" width="35" />
            </div>
            """,
            unsafe_allow_html=True
        )

    with col2:
        st.markdown(
            """
            <div style="
                display: flex;
                align-items: center;
                height: 65px;
            ">
                <div style="text-align: justify; font-size: 15px;">
                    <span style="color: #FF1BF1; font-weight: 800;">TCO Comparison by Year + Region</span>
                    combines yearly and regional TCO analysis to provide a more comprehensive view of 
                    vendor pricing across time and location.
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.divider()

    # 4️⃣ TCO Comparison Round by Round
    col1, col2 = st.columns([1, 9])  # kiri lebih sempit

    # background-color: #26BDAD;
    with col1:
        st.markdown(
            """
            <div style="
                background-color: #26BDAD; 
                border-radius: 8px;
                padding: 20px;
                width: 65px;
                height: 65px;
                display: flex;
                align-items: center;
                justify-content: center;
                box-shadow: 0 2px 5px rgba(0,0,0,0.15);
            ">
                <img src="https://cdn-icons-png.flaticon.com/512/4624/4624081.png" width="35" />
            </div>
            """,
            unsafe_allow_html=True
        )

    with col2:
        st.markdown(
            """
            <div style="
                display: flex;
                align-items: center;
                height: 65px;
            ">
                <div style="text-align: justify; font-size: 15px;">
                    <span style="color: #26BDAD; font-weight: 800;">TCO Comparison Round by Round</span>
                    evaluates TCO changes across negotiation rounds to track vendor pricing progress and
                    identify the most competitive movements.
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.divider()

    # 5️⃣ UPL Comparison
    col1, col2 = st.columns([1, 9])  # kiri lebih sempit

    # background-color: #FFCB09;
    with col1:
        st.markdown(
            """
            <div style="
                background-color: #C7FF00; 
                border-radius: 8px;
                padding: 20px;
                width: 65px;
                height: 65px;
                display: flex;
                align-items: center;
                justify-content: center;
                box-shadow: 0 2px 5px rgba(0,0,0,0.15);
            ">
                <img src="https://cdn-icons-png.flaticon.com/512/4624/4624116.png" width="35" />
            </div>
            """,
            unsafe_allow_html=True
        )

    with col2:
        st.markdown(
            """
            <div style="
                display: flex;
                align-items: center;
                height: 65px;
            ">
                <div style="text-align: justify; font-size: 15px;">
                    <span style="color: #C7FF00; font-weight: 800;">UPL Comparison</span>
                    compares UPL values across vendors to determine the most competitive 
                    item-level pricing.
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.divider()

    # 6️⃣ UPL Comparison Round by Round
    col1, col2 = st.columns([1, 9])  # kiri lebih sempit

    # background-color: #FF77A9;
    with col1:
        st.markdown(
            """
            <div style="
                background-color: #FFCB09;
                border-radius: 8px;
                padding: 20px;
                width: 65px;
                height: 65px;
                display: flex;
                align-items: center;
                justify-content: center;
                box-shadow: 0 2px 5px rgba(0,0,0,0.15);
            ">
                <img src="https://cdn-icons-png.flaticon.com/512/4624/4624030.png" width="35" />
            </div>
            """,
            unsafe_allow_html=True
        )

    with col2:
        st.markdown(
            """
            <div style="
                display: flex;
                align-items: center;
                height: 65px;
            ">
                <div style="text-align: justify; font-size: 15px;">
                    <span style="color: #FFCB09; font-weight: 800;">UPL Comparison Round by Round</span>
                    tracks UPL adjustments throughout negotiation rounds to understand pricing dynamics
                    at the item level.
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.divider()

    # 7️⃣ Standard Deviation
    col1, col2 = st.columns([1, 9])  # kiri lebih sempit

    with col1:
        st.markdown(
            """
            <div style="
                background-color: #ED1C24; 
                border-radius: 8px;
                padding: 20px;
                width: 65px;
                height: 65px;
                display: flex;
                align-items: center;
                justify-content: center;
                box-shadow: 0 2px 5px rgba(0,0,0,0.15);
            ">
                <img src="https://cdn-icons-png.flaticon.com/512/4624/4624098.png" width="35" />
            </div>
            """,
            unsafe_allow_html=True
        )

    with col2:
        st.markdown(
            """
            <div style="
                display: flex;
                align-items: center;
                height: 65px;
            ">
                <div style="text-align: justify; font-size: 15px;">
                    <span style="color: #ED1C24; font-weight: 800;">Standard Deviation</span>
                    measures the price variation across vendors to assess pricing stability 
                    and the consistency of commercial offers.
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.divider()

    # # 1️⃣ Single Layer TCO
    # st.markdown("##### 1️⃣ Single Layer TCO")
    # # st.markdown("""
    # # This feature is designed to process pricing data where all bidders are presented within a single sheet.
    # # The required format is:
    # # """)

    # st.markdown(
    #     """
    #     <div style="text-align: justify;">
    #         This feature is designed to process pricing data where all bidders are presented within a single sheet.  
    #         The required format is:
    #     </div>
    #     """,
    #     unsafe_allow_html=True
    # )

    # st.write("")
    # st.image("assets/single-layer-tco.png", caption="Single Layer TCO Template Structure")

    # # st.markdown("""
    # # Each component's value for every bidder is summed to calculate the total TCO per bidder, and then ranked to identify the most cost-efficient option.
    # # Additionaly, the tool sums each component across all bidders to provide a component-level cost comparison, visualized through a bar chart.
    # # """)

    # st.markdown(
    #     """
    #     <div style="text-align: justify;">
    #         The column structure must be followed, but the layout is flexible —
    #         you can have any number of <span style="color:#FF69B4;">components</span> 
    #         or <span style="color:#FF69B4;">bidders</span>, 
    #         and the column names are fully customizable.
    #     </div>
    #     """,
    #     unsafe_allow_html=True
    # )
    
    # st.write("")

    # # st.markdown("""
    # # The column structure must be followed, but the layout is flexible —
    # # you can have any number of bidders or components, and the column names are fully customizable.
    # # """)

    # # 2️⃣ Double Layer TCO
    # st.markdown("##### 2️⃣ Double Layer TCO")
    # # st.markdown("""
    # # This feature is designed for pricing evaluations where each bidder's data is stored in a separate sheet, 
    # # and pricing is spread across multiple years.
    # # The required format is: 
    # # """)

    # st.markdown(
    #     """
    #     <div style="text-align: justify;">
    #         This feature is designed for pricing evaluations where each bidder's data is stored in a separate sheet, 
    #         and pricing is spread across multiple years.
    #         The required format is: 
    #     </div>
    #     """,
    #     unsafe_allow_html=True
    # )

    # st.write("")

    # st.image("assets/double-layer-tco.png")
    # st.image("assets/all-bidders.png", caption="Double Layer TCO Template Structure")

    # # st.markdown("""
    # # The tool automatically reads all sheets, adds a 'Total' column (the sum of all year columns per component), 
    # # and merges all bidders' total columns into a single consolidated table.
    # # It then ranks the bidders based on total TCO and displays a bar chart comparing total and component-level costs.
    # # """)

    # st.markdown(
    #     """
    #     <div style="text-align: justify;">
    #         The structure and column order must remain consistent across all sheets —
    #         you can have any number of <span style="color:#FF69B4;">components</span>
    #         or <span style="color:#FF69B4;">years</span> (3-Year or 5-Year TCO, etc), 
    #         and the column names are fully customizable.
    #     </div>
    #     """,
    #     unsafe_allow_html=True
    # )

    # st.write("")

    # # st.markdown("""
    # # The structure and column order must remain consistent across all sheets —
    # # you can have any number of components or years (3-Year or 5-Year TCO, etc), and the column names are fully customizable.
    # # """)
