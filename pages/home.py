import streamlit as st
import pandas as pd

def page():
    st.markdown(
        """
        <div style="font-size:2.25rem; font-weight:700; margin-bottom:9px">
            🏡 Intro: Bid & Price Analytics Tool
        </div>
        """,
        unsafe_allow_html=True
    )
    # st.header("🏡 Intro: Bid & Price Analytics Tool")
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
        <div style="text-align: justify; font-size: 15px; margin-bottom: 20px">
            <span style="color: orange; font-weight: 600;"> Bid & Price Analytics Tool</span>
            is a <a href="https://streamlit.io" target="_blank" 
            style="color:#4DA3FF;">
            Streamlit</a>-powered application that automates
            commercial bid comparison by analyzing multi-vendor submissions, identifying the 
            most competitive offers, and generating clear, ready-to-share analytical outputs.
        </div>
    """, unsafe_allow_html=True)

    # Main Features
    # st.subheader("Main Features")
    st.markdown("#### Main Features")

    columns = ["Features", "Description"]
    data = [
        [
            "Data Cleaning",
            "Provides a single menu to extract multiple table at once"
        ],

        [
            "Analysis Modules",
            "Provides 7 menus to generate comprehensive comparisons and data summaries"
        ]
    ]
    df = pd.DataFrame(data, columns=columns)
    st.dataframe(df, hide_index=True)

    st.divider()

    # Data Cleaning
    # st.subheader("Data Cleaning")
    st.markdown("#### Data Cleaning")
    st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

    # ✂️ Table Extraction
    col1, col2 = st.columns([1, 9])  # kiri lebih sempit

    with col1:
        st.markdown(
            """
            <style>
            .hover-div {
                background-color: #0073FF;
                border-radius: 8px;
                padding: 20px;
                width: 65px;
                height: 65px;
                display: flex;
                align-items: center;
                justify-content: center;
                box-shadow: 0 2px 5px rgba(0,0,0,0.15);
                transition: background-color 0.3s, transform 0.3s;
            }

            .hover-div:hover {
                background-color: #99C2FF;
                transform: scale(1.05);
            }
            </style>

            <a href="https://inspire-project-user-guide-table-extraction.streamlit.app/">
                <div class="hover-div">
                    <img src="https://cdn-icons-png.flaticon.com/512/4624/4624188.png" width="35" />
                </div>
            </a>
            """,
            unsafe_allow_html=True
        )

    with col2:
        st.markdown(
            """
            <style>
            a.hover-link {
                color: #0073FF;
                font-weight: 800;
                text-decoration: none;
            }
            a.hover-link:hover {
                color: #99C2FF;
            }
            </style>

            <div style="
                display: flex;
                align-items: center;
                height: 65px;
            ">
                <div style="text-align: justify; font-size: 15px;">
                    <a href="https://inspire-project-user-guide-table-extraction.streamlit.app/"
                    class="hover-link">Table Extraction</a>
                    is used to extract tables from multi-sheet files where each sheet contain multiple
                    tables, arranged either horizontally or vertically.
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.divider()
    
    # Analysis Modules
    # st.subheader("Analysis Modules")
    st.markdown("#### Analysis Modules")
    st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

    # 1️⃣ TCO Comparison by Year
    col1, col2 = st.columns([1, 9])  # kiri lebih sempit

    with col1:
        st.markdown(
            """
            <style>
            .hover-div1 {
                background-color: #BC13FE;
                border-radius: 8px;
                padding: 20px;
                width: 65px;
                height: 65px;
                display: flex;
                align-items: center;
                justify-content: center;
                box-shadow: 0 2px 5px rgba(0,0,0,0.15);
                transition: background-color 0.3s, transform 0.3s;
            }

            .hover-div1:hover {
                background-color: #E299FF;
                transform: scale(1.05);
            }
            </style>

            <a href="https://inspire-project-user-guide-tco-comparison-by-year.streamlit.app/">
                <div class="hover-div1">
                    <img src="https://cdn-icons-png.flaticon.com/512/14991/14991730.png" width="35" />
                </div>
            </a>
            """,
            unsafe_allow_html=True
        )

    with col2:
        st.markdown(
            """
            <style>
            a.hover-link1 {
                color: #BC13FE;
                font-weight: 800;
                text-decoration: none;
            }
            a.hover-link1:hover {
                color: #E299FF;
            }
            </style>

            <div style="
                display: flex;
                align-items: center;
                height: 65px;
            ">
                <div style="text-align: justify; font-size: 15px;">
                    <a href="https://inspire-project-user-guide-tco-comparison-by-year.streamlit.app/"
                    class="hover-link1">TCO Comparison by Year</a>
                    compares TCO across vendors by analyzing year-over-year price changes to identify long-
                    term cost trends.
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
            <style>
            .hover-div2 {
                background-color: #FF2EC4;
                border-radius: 8px;
                padding: 20px;
                width: 65px;
                height: 65px;
                display: flex;
                align-items: center;
                justify-content: center;
                box-shadow: 0 2px 5px rgba(0,0,0,0.15);
                transition: background-color 0.3s, transform 0.3s;
            }

            .hover-div2:hover {
                background-color: #FF99E0;
                transform: scale(1.05);
            }
            </style>

            <a href="https://inspire-project-user-guide-tco-comparison-by-region.streamlit.app/">
                <div class="hover-div2">
                    <img src="https://cdn-icons-png.flaticon.com/512/15366/15366033.png" width="35" />
                </div>
            </a>
            """,
            unsafe_allow_html=True
        )

    with col2:
        st.markdown(
            """
            <style>
            a.hover-link2 {
                color: #FF2EC4;
                font-weight: 800;
                text-decoration: none;
            }
            a.hover-link2:hover {
                color: #FF99E0;
            }
            </style>

            <div style="
                display: flex;
                align-items: center;
                height: 65px;
            ">
                <div style="text-align: justify; font-size: 15px;">
                    <a href="https://inspire-project-user-guide-tco-comparison-by-region.streamlit.app/"
                    class="hover-link2">TCO Comparison by Region</a>
                    analyzes TCO differences across regions to highlight cost variations based on geographic 
                    requirements.
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
            <style>
            .hover-div3 {
                background-color: #FF1BF1;
                border-radius: 8px;
                padding: 20px;
                width: 65px;
                height: 65px;
                display: flex;
                align-items: center;
                justify-content: center;
                box-shadow: 0 2px 5px rgba(0,0,0,0.15);
                transition: background-color 0.3s, transform 0.3s;
            }

            .hover-div3:hover {
                background-color: #FF7AF7;
                transform: scale(1.05);
            }
            </style>

            <a href="https://inspire-project-user-guide-tco-comparison-by-year-region.streamlit.app/">
                <div class="hover-div3">
                    <img src="https://cdn-icons-png.flaticon.com/512/4624/4624053.png" width="35" />
                </div>
            </a>
            """,
            unsafe_allow_html=True
        )

    with col2:
        st.markdown(
            """
            <style>
            a.hover-link3 {
                color: #FF1BF1;
                font-weight: 800;
                text-decoration: none;
            }
            a.hover-link3:hover {
                color: #FF7AF7;
            }
            </style>

            <div style="
                display: flex;
                align-items: center;
                height: 65px;
            ">
                <div style="text-align: justify; font-size: 15px;">
                    <a href="https://inspire-project-user-guide-tco-comparison-by-year-region.streamlit.app/"
                    class="hover-link3">TCO Comparison by Year + Region</a>
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
            <style>
            .hover-div4 {
                background-color: #26BDAD;
                border-radius: 8px;
                padding: 20px;
                width: 65px;
                height: 65px;
                display: flex;
                align-items: center;
                justify-content: center;
                box-shadow: 0 2px 5px rgba(0,0,0,0.15);
                transition: background-color 0.3s, transform 0.3s;
            }

            .hover-div4:hover {
                background-color: #7FE2DC;
                transform: scale(1.05);
            }
            </style>

            <a href="https://inspire-project-user-guide-tco-comparison-round-by-round.streamlit.app/">
                <div class="hover-div4">
                    <img src="https://cdn-icons-png.flaticon.com/512/4624/4624081.png" width="35" />
                </div>
            </a>
            """,
            unsafe_allow_html=True
        )

    with col2:
        st.markdown(
            """
            <style>
            a.hover-link4 {
                color: #26BDAD;
                font-weight: 800;
                text-decoration: none;
            }
            a.hover-link4:hover {
                color: #7FE2DC;
            }
            </style>

            <div style="
                display: flex;
                align-items: center;
                height: 65px;
            ">
                <div style="text-align: justify; font-size: 15px;">
                    <a href="https://inspire-project-user-guide-tco-comparison-round-by-round.streamlit.app/"
                    class="hover-link4">TCO Comparison Round by Round</a>
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
            <style>
            .hover-div5 {
                background-color: #C7FF00;
                border-radius: 8px;
                padding: 20px;
                width: 65px;
                height: 65px;
                display: flex;
                align-items: center;
                justify-content: center;
                box-shadow: 0 2px 5px rgba(0,0,0,0.15);
                transition: background-color 0.3s, transform 0.3s;
            }

            .hover-div5:hover {
                background-color: #E5FF66;
                transform: scale(1.05);
            }
            </style>

            <a href="https://inspire-project-user-guide-upl-comparison.streamlit.app/">
                <div class="hover-div5">
                    <img src="https://cdn-icons-png.flaticon.com/512/4624/4624116.png" width="35" />
                </div>
            </a>
            """,
            unsafe_allow_html=True
        )

    with col2:
        st.markdown(
            """
            <style>
            a.hover-link5 {
                color: #C7FF00;
                font-weight: 800;
                text-decoration: none;
            }
            a.hover-link5:hover {
                color: #E5FF66;
            }
            </style>

            <div style="
                display: flex;
                align-items: center;
                height: 65px;
            ">
                <div style="text-align: justify; font-size: 15px;">
                    <a href="https://inspire-project-user-guide-upl-comparison.streamlit.app/"
                    class="hover-link5">UPL Comparison</a>
                    compares UPL values across multiple vendors in order to identify and determine the most 
                    competitive item-level pricing.
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.divider()

    # 6️⃣ UPL Comparison Round by Round
    col1, col2 = st.columns([1, 9])  # kiri lebih sempit

    with col1:
        st.markdown(
            """
            <style>
            .hover-div6 {
                background-color: #FFCB09;
                border-radius: 8px;
                padding: 20px;
                width: 65px;
                height: 65px;
                display: flex;
                align-items: center;
                justify-content: center;
                box-shadow: 0 2px 5px rgba(0,0,0,0.15);
                transition: background-color 0.3s, transform 0.3s;
            }

            .hover-div6:hover {
                background-color: #FFE066;
                transform: scale(1.05);
            }
            </style>

            <a href="https://inspire-project-user-guide-upl-comparison-round-by-round.streamlit.app/">
                <div class="hover-div6">
                    <img src="https://cdn-icons-png.flaticon.com/512/4624/4624030.png" width="35" />
                </div>
            </a>
            """,
            unsafe_allow_html=True
        )

    with col2:
        st.markdown(
            """
            <style>
            a.hover-link6 {
                color: #FFCB09;
                font-weight: 800;
                text-decoration: none;
            }
            a.hover-link6:hover {
                color: #FFE066;
            }
            </style>

            <div style="
                display: flex;
                align-items: center;
                height: 65px;
            ">
                <div style="text-align: justify; font-size: 15px;">
                    <a href="https://inspire-project-user-guide-upl-comparison-round-by-round.streamlit.app/"
                    class="hover-link6">UPL Comparison Round by Round</a>
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
            <style>
            .hover-div7 {
                background-color: #ED1C24;
                border-radius: 8px;
                padding: 20px;
                width: 65px;
                height: 65px;
                display: flex;
                align-items: center;
                justify-content: center;
                box-shadow: 0 2px 5px rgba(0,0,0,0.15);
                transition: background-color 0.3s, transform 0.3s;
            }

            .hover-div7:hover {
                background-color: #F56667;
                transform: scale(1.05);
            }
            </style>

            <a href="https://inspire-project-user-guide-standard-deviation.streamlit.app/">
                <div class="hover-div7">
                    <img src="https://cdn-icons-png.flaticon.com/512/4624/4624098.png" width="35" />
                </div>
            </a>
            """,
            unsafe_allow_html=True
        )

    with col2:
        st.markdown(
            """
            <style>
            a.hover-link7 {
                color: #ED1C24;
                font-weight: 800;
                text-decoration: none;
            }
            a.hover-link7:hover {
                color: #F56667;
            }
            </style>

            <div style="
                display: flex;
                align-items: center;
                height: 65px;
            ">
                <div style="text-align: justify; font-size: 15px;">
                    <a href="https://inspire-project-user-guide-standard-deviation.streamlit.app/"
                    class="hover-link7">Standard Deviation</a>
                    measures the price variation across vendors to assess pricing stability 
                    and the consistency of commercial offers.
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.divider()

    # Video Tutorial
    # st.subheader("Video Tutorial")
    st.markdown("#### Video Tutorial")

    st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            I have also included a video tutorial, which you can access through the 
            <span style="background:#FF0000; padding:2px 4px; border-radius:6px; font-weight:600; 
            font-size: 0.75rem; color: black">YouTube</span> link below.
        </div>
    """,
    unsafe_allow_html=True
)

    st.video("https://youtu.be/1p_WerKwFJA?si=vevaZRIjFDUlIz_X")
