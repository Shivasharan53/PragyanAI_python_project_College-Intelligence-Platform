
import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="CampusPulse",
    page_icon="🎓",
    layout="wide"
)

st.markdown("""
<style>
.stApp {
    background: #f3f7fc;
}
.block-container {
    padding-top: 1.2rem;
}
[data-testid="stMetric"] {
    background: white;
    border: 1px solid #dce6f5;
    padding: 16px;
    border-radius: 12px;
}
</style>
""", unsafe_allow_html=True)

st.title("🎓 CampusPulse")
st.subheader("Live College Intelligence Platform")
st.caption("Powered by: Shivasharanappa")

st.sidebar.title("CampusPulse")
page = st.sidebar.radio(
    "Navigation",
    [
        "Dashboard",
        "Upload & Process",
        "Tier Results",
        "College Search",
        "College Map",
        "Reviews",
        "RAG Chatbot",
        "About"
    ]
)

if "colleges" not in st.session_state:
    st.session_state.colleges = pd.DataFrame()

if "results" not in st.session_state:
    st.session_state.results = pd.DataFrame()

st.header(page)

if page in ["Dashboard", "Upload & Process"]:
    st.markdown("### 1. Upload Excel File")

    uploaded_file = st.file_uploader(
        "Upload college information",
        type=["xlsx", "xls"]
    )

    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            df.columns = [
                str(c).strip() for c in df.columns
            ]

            st.session_state.colleges = df
            st.success(
                f"Loaded {len(df)} rows from Excel."
            )
            st.dataframe(
                df,
                use_container_width=True,
                hide_index=True
            )

        except Exception as error:
            st.error(f"Excel processing failed: {error}")

    colleges = st.session_state.colleges

    if not colleges.empty:
        name_column = next(
            (
                c for c in colleges.columns
                if c.strip().lower() in [
                    "college",
                    "college name",
                    "name"
                ]
            ),
            None
        )

        if name_column:
            names = colleges[name_column].dropna().astype(str).tolist()

            selected = st.multiselect(
                "Select colleges",
                options=names
            )

            if st.button("Process selected colleges"):
                if not selected:
                    st.warning("Select at least one college.")
                else:
                    chosen = colleges[
                        colleges[name_column].astype(str).isin(selected)
                    ].copy()

                    st.session_state.results = chosen
                    st.success(
                        f"Selected {len(chosen)} colleges."
                    )
        else:
            st.warning(
                "Please provide a column named College "
                "or College Name."
            )

if page in ["Dashboard", "Tier Results"]:
    st.markdown("### 2. Tier Classification Results")

    results = st.session_state.results

    if not results.empty:
        st.info(
            "The tier classification rules will be added next. "
            "These are extracted records, not verified tier results."
        )
        st.dataframe(
            results,
            use_container_width=True,
            hide_index=True
        )

        st.download_button(
            "Download extracted data",
            data=results.to_csv(index=False).encode("utf-8"),
            file_name="campuspulse_colleges.csv",
            mime="text/csv"
        )
    else:
        st.info(
            "Upload an Excel file and process selected colleges "
            "to display results."
        )

if page == "College Search":
    st.write("College search will be connected to the extracted dataset.")

if page == "College Map":
    st.write("Verified college locations and map markers will be added here.")

if page == "Reviews":
    st.write("Ratings and reviews will be connected to a supported provider.")

if page == "RAG Chatbot":
    st.write("The RAG chatbot will be connected to the college knowledge base.")

if page == "About":
    st.markdown("""
    **CampusPulse — Live College Intelligence Platform**

    Excel extraction, college classification, maps, reviews,
    and a retrieval-augmented college assistant.

    **Powered by: Shivasharanappa**
    """)
