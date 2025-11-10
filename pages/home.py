import streamlit as st
import pandas as pd

def page():
    st.title("🏡 Intro: Bid & Price Analytics Tool")
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

    st.markdown("""
    :orange[***Bid & Price Analytics Tool***] is a [Streamlit](https://streamlit.io)-based analytical application 
    designed to assist the CTPO - IOH team in performing automatic commercial evaluations of bidders.
    """)
    
    # Main Features
    st.subheader("Main Features")
    # col1, col2 = st.columns([3, 1])

    # with col1:
    data = {
        "Feature": [
            "🏡 Home",
            "1️⃣ Single Layer TCO",
            "2️⃣ Double Layer TCO"
        ],
        "Description": [
            "Introduction and user guide for the Bid & Price Analytics Tool",
            "Analyze bidder-level pricing (single-sheet file)",
            "Analyze multi-year pricing (multi-sheet file)"
        ]
    }

    df = pd.DataFrame(data)
    st.dataframe(df, width="stretch", hide_index=True)

    # with col2:
    #     st.image("assets/dap_man.png", width=148)

    # st.divider()

    # How to Use
    st.subheader("How to Use")

    # 1️⃣ Single Layer TCO
    st.markdown("##### 1️⃣ Single Layer TCO")
    # st.markdown("""
    # This feature is designed to process pricing data where all bidders are presented within a single sheet.
    # The required format is:
    # """)

    st.markdown(
        """
        <div style="text-align: justify;">
            This feature is designed to process pricing data where all bidders are presented within a single sheet.  
            The required format is:
        </div>
        """,
        unsafe_allow_html=True
    )

    st.write("")
    st.image("assets/single-layer-tco.png", caption="Single Layer TCO Template Structure")

    # st.markdown("""
    # Each component's value for every bidder is summed to calculate the total TCO per bidder, and then ranked to identify the most cost-efficient option.
    # Additionaly, the tool sums each component across all bidders to provide a component-level cost comparison, visualized through a bar chart.
    # """)

    st.markdown(
        """
        <div style="text-align: justify;">
            The column structure must be followed, but the layout is flexible —
            you can have any number of <span style="color:#FF69B4;">components</span> 
            or <span style="color:#FF69B4;">bidders</span>, 
            and the column names are fully customizable.
        </div>
        """,
        unsafe_allow_html=True
    )
    
    st.write("")

    # st.markdown("""
    # The column structure must be followed, but the layout is flexible —
    # you can have any number of bidders or components, and the column names are fully customizable.
    # """)

    # 2️⃣ Double Layer TCO
    st.markdown("##### 2️⃣ Double Layer TCO")
    # st.markdown("""
    # This feature is designed for pricing evaluations where each bidder's data is stored in a separate sheet, 
    # and pricing is spread across multiple years.
    # The required format is: 
    # """)

    st.markdown(
        """
        <div style="text-align: justify;">
            This feature is designed for pricing evaluations where each bidder's data is stored in a separate sheet, 
            and pricing is spread across multiple years.
            The required format is: 
        </div>
        """,
        unsafe_allow_html=True
    )

    st.write("")

    st.image("assets/double-layer-tco.png")
    st.image("assets/all-bidders.png", caption="Double Layer TCO Template Structure")

    # st.markdown("""
    # The tool automatically reads all sheets, adds a 'Total' column (the sum of all year columns per component), 
    # and merges all bidders' total columns into a single consolidated table.
    # It then ranks the bidders based on total TCO and displays a bar chart comparing total and component-level costs.
    # """)

    st.markdown(
        """
        <div style="text-align: justify;">
            The structure and column order must remain consistent across all sheets —
            you can have any number of <span style="color:#FF69B4;">components</span>
            or <span style="color:#FF69B4;">years</span> (3-Year or 5-Year TCO, etc), 
            and the column names are fully customizable.
        </div>
        """,
        unsafe_allow_html=True
    )

    st.write("")

    # st.markdown("""
    # The structure and column order must remain consistent across all sheets —
    # you can have any number of components or years (3-Year or 5-Year TCO, etc), and the column names are fully customizable.
    # """)

    


    


