import io
import re
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="CampusPulse",
    page_icon="🎓",
    layout="wide"
)

# -------------------- STYLE --------------------

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
    border-radius: 12px;
    padding: 14px;
}
h1, h2, h3 {
    color: #123b78;
}
</style>
""", unsafe_allow_html=True)


# -------------------- DATA PROCESSING --------------------

ALIASES = {
    "college": "College",
    "college name": "College",
    "name": "College",
    "institution": "College",
    "institution name": "College",
    "city": "City",
    "town": "City",
    "website": "Website",
    "web site": "Website",
    "url": "Website",
    "wikipedia": "Wikipedia",
    "wikipedia url": "Wikipedia",
    "address": "Address",
    "full address": "Address",
    "district": "District",
    "state": "State",
    "pin code": "PIN Code",
    "pincode": "PIN Code",
    "postal code": "PIN Code",
    "affiliation": "Affiliation",
    "affiliated to": "Affiliation",
    "university affiliation": "Affiliation",
    "status": "Status",
    "college status": "Status",
    "type": "Type",
    "college type": "Type",
    "institution type": "Type",
}


def normalize_columns(df):
    rename = {}

    for column in df.columns:
        key = re.sub(
            r"\s+", " ", str(column).strip().lower()
        )
        if key in ALIASES:
            rename[column] = ALIASES[key]

    return df.rename(columns=rename)


def load_colleges(uploaded_file):
    df = pd.read_excel(uploaded_file)
    df = normalize_columns(df)

    if "College" not in df.columns:
        raise ValueError(
            "A college-name column is required. "
            "Use College, College Name, Name, or Institution."
        )

    expected_columns = [
        "City", "Website", "Wikipedia", "Address",
        "District", "State", "PIN Code",
        "Affiliation", "Status", "Type"
    ]

    for column in expected_columns:
        if column not in df.columns:
            df[column] = ""

    df["College"] = df["College"].fillna("").astype(str).str.strip()
    df = df[df["College"] != ""]
    df = df.drop_duplicates(subset=["College"], keep="first")

    return df.reset_index(drop=True)


def classify_college(row):
    # Only classify from evidence contained in the spreadsheet.
    # These are project rules, not an official college ranking.

    status = str(row.get("Status", "")).lower()
    kind = str(row.get("Type", "")).lower()
    affiliation = str(row.get("Affiliation", "")).lower()

    evidence = f"{status} {kind} {affiliation}".strip()

    if re.search(r"\bautonomous\b", evidence):
        return (
            "Tier 1",
            "Autonomous status indicated in spreadsheet"
        )

    if re.search(r"\buniversit(?:y|ies)\b", evidence):
        return (
            "Tier 1",
            "University status indicated in spreadsheet"
        )

    if "vtu" in evidence and (
        "affiliat" in evidence or "affiliation" in evidence
    ):
        return (
            "Tier 1",
            "VTU affiliation indicated in spreadsheet"
        )

    if evidence and evidence not in {"nan", "none"}:
        return (
            "Tier 2",
            "Institution details available; verify independently"
        )

    return (
        "Tier 3",
        "Insufficient affiliation or institution-status information"
    )


def classify_dataframe(df):
    result = df.copy()

    classifications = result.apply(
        classify_college, axis=1
    )

    result["Tier"] = [item[0] for item in classifications]
    result["Reason"] = [item[1] for item in classifications]

    return result


def make_excel(df):
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(
            writer,
            index=False,
            sheet_name="College Results"
        )

    buffer.seek(0)
    return buffer.getvalue()


# -------------------- SESSION STATE --------------------

if "colleges" not in st.session_state:
    st.session_state.colleges = pd.DataFrame()

if "results" not in st.session_state:
    st.session_state.results = pd.DataFrame()


# -------------------- HEADER --------------------

st.title("🎓 CampusPulse")
st.subheader("Live College Intelligence Platform")
st.caption("Powered by: Shivasharanappa")

st.sidebar.title("🎓 CampusPulse")

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


# -------------------- UPLOAD & PROCESS --------------------

if page in ["Dashboard", "Upload & Process"]:
    st.header("1. Upload & Extract Colleges")

    uploaded_file = st.file_uploader(
        "Upload your college Excel file",
        type=["xlsx", "xls"],
        key="college_excel_upload"
    )

    if uploaded_file is not None:
        try:
            df = load_colleges(uploaded_file)
            st.session_state.colleges = df
            st.success(f"Extracted {len(df)} colleges.")

        except Exception as error:
            st.error(f"Excel processing failed: {error}")

    colleges = st.session_state.colleges

    if not colleges.empty:
        st.subheader("2. Extracted Colleges")

        st.metric("Total Colleges Found", len(colleges))

        st.dataframe(
            colleges,
            use_container_width=True,
            hide_index=True
        )

        st.subheader("3. Select Colleges & Process")

        names = colleges["College"].tolist()

        selected = st.multiselect(
            "Search and select one or more colleges",
            options=names,
            key="selected_colleges"
        )

        col1, col2 = st.columns(2)

        with col1:
            select_all = st.button(
                "Select All Colleges",
                use_container_width=True
            )

        with col2:
            process = st.button(
                "Process Selected Colleges",
                type="primary",
                use_container_width=True
            )

        if select_all:
            st.session_state.selected_colleges = names
            st.rerun()

        if process:
            chosen_names = st.session_state.get(
                "selected_colleges", []
            )

            if not chosen_names:
                st.warning("Select at least one college first.")
            else:
                chosen = colleges[
                    colleges["College"].isin(chosen_names)
                ].copy()

                st.session_state.results = classify_dataframe(
                    chosen
                )

                st.success(
                    f"Processed {len(chosen)} colleges."
                )

    else:
        st.info(
            "Upload an Excel file to start. "
            "The file must contain a college-name column."
        )


# -------------------- TIER RESULTS --------------------

if page in ["Dashboard", "Tier Results"]:
    st.header("4. Tier Classification Results")

    results = st.session_state.results

    if not results.empty:
        tier1, tier2, tier3 = st.columns(3)

        for column, tier, label in [
            (tier1, "Tier 1", "🏆 Tier 1 Colleges"),
            (tier2, "Tier 2", "🥈 Tier 2 Colleges"),
            (tier3, "Tier 3", "📋 Tier 3 Colleges"),
        ]:
            subset = results[results["Tier"] == tier]

            column.metric(label, len(subset))

            with column:
                st.dataframe(
                    subset[["College", "City", "Reason"]],
                    use_container_width=True,
                    hide_index=True
                )

        st.subheader("All Processed Results")

        st.dataframe(
            results,
            use_container_width=True,
            hide_index=True
        )

        st.download_button(
            "⬇️ Download Classified Excel",
            data=make_excel(results),
            file_name="campuspulse_classified_colleges.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            )
        )

    else:
        st.info(
            "Upload an Excel file and process selected colleges "
            "to see the tier results."
        )


# -------------------- COLLEGE SEARCH --------------------

if page == "College Search":
    st.header("College Search")

    df = st.session_state.colleges

    if not df.empty:
        query = st.text_input(
            "Search by college name, city, or district"
        )

        searchable = df.fillna("").astype(str)

        if query.strip():
            mask = searchable.apply(
                lambda column: column.str.contains(
                    query,
                    case=False,
                    regex=False
                )
            ).any(axis=1)

            searchable = searchable[mask]

        st.dataframe(
            searchable,
            use_container_width=True,
            hide_index=True
        )
    else:
        st.info("Upload an Excel file first.")


# -------------------- FUTURE MODULES --------------------

if page == "College Map":
    st.header("📍 College Map")
    st.info(
        "Next: verified addresses, coordinates, and map markers."
    )

if page == "Reviews":
    st.header("⭐ College Reviews")
    st.info(
        "Next: connect a supported provider for current ratings "
        "and reviews."
    )

if page == "RAG Chatbot":
    st.header("💬 RAG College Assistant")
    st.info(
        "Next: index college records and connect the retrieval "
        "and answer-generation pipeline."
    )

if page == "About":
    st.header("About CampusPulse")
    st.write(
        "CampusPulse is a college intelligence platform for "
        "Excel extraction, tier classification, college discovery, "
        "maps, reviews, and RAG-based question answering."
    )
    st.markdown("**Powered by: Shivasharanappa**")
