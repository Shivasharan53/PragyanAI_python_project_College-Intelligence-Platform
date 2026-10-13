import io
import re
import os
import pandas as pd
import requests
import streamlit as st
import folium

from streamlit_folium import st_folium
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity


# ==========================================================
# CampusPulse — Live College Intelligence Platform
# Powered by Shivasharanappa
# ==========================================================

st.set_page_config(
    page_title="CampusPulse | College Intelligence",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

# -------------------------- CSS ----------------------------

st.markdown("""
<style>
.stApp {
    background: #f3f7fc;
    color: #142746;
}
.block-container {
    max-width: 1800px;
    padding: 1rem 1rem 2rem;
}
header[data-testid="stHeader"] {
    background: #f3f7fc;
}
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg,#111d39,#172c49);
}
section[data-testid="stSidebar"] * {
    color: #ffffff;
}
.cp-header {
    background: linear-gradient(110deg,#151a48,#0755a1);
    color: white;
    padding: 20px 24px;
    border: 1px solid #285a9b;
    border-radius: 14px;
    margin-bottom: 14px;
}
.cp-header h1 {
    color: white;
    font-size: clamp(25px,3vw,38px);
    margin: 0;
    line-height: 1.2;
}
.cp-header p {
    margin: 5px 0 0;
    color: #e4edff;
}
.cp-brand {
    color: #ffd84d;
    font-weight: 800;
}
.cp-card {
    background: #ffffff;
    border: 1px solid #d7e4f5;
    border-radius: 12px;
    padding: 14px;
    margin-bottom: 10px;
}
.cp-heading {
    font-size: 18px;
    color: #073f91;
    font-weight: 800;
    margin-bottom: 4px;
}
.cp-muted {
    font-size: 12px;
    color: #62748e;
}
.cp-tier1 {
    background: #e3fff0;
    border: 1px solid #71d9a3;
    border-radius: 10px;
    padding: 10px;
}
.cp-tier2 {
    background: #fff4d4;
    border: 1px solid #edcd70;
    border-radius: 10px;
    padding: 10px;
}
.cp-tier3 {
    background: #ffe7e7;
    border: 1px solid #f2a0a0;
    border-radius: 10px;
    padding: 10px;
}
.cp-tier1 h4 {color:#075d3a;}
.cp-tier2 h4 {color:#7d5100;}
.cp-tier3 h4 {color:#a21f2b;}
[data-testid="stMetric"] {
    background: #fff;
    border: 1px solid #d8e4f5;
    border-radius: 10px;
    padding: 10px;
}
.stButton > button {
    border-radius: 8px;
    min-height: 40px;
    font-weight: 650;
}
a {overflow-wrap:anywhere;}
</style>
""", unsafe_allow_html=True)


# ------------------------ SECRETS --------------------------

def secret_value(name):
    try:
        return str(st.secrets.get(name, "")).strip()
    except Exception:
        return os.getenv(name, "").strip()


GROQ_API_KEY = secret_value("GROQ_API_KEY")


# --------------------- EXCEL HELPERS -----------------------

ALIASES = {
    "college": "College",
    "college name": "College",
    "name": "College",
    "institution": "College",
    "institution name": "College",
    "city": "City",
    "town": "City",
    "district": "District",
    "state": "State",
    "website": "Website",
    "web site": "Website",
    "url": "Website",
    "wikipedia": "Wikipedia",
    "wikipedia url": "Wikipedia",
    "address": "Address",
    "full address": "Address",
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
    "latitude": "Latitude",
    "longitude": "Longitude",
}


def text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_columns(df):
    rename = {}

    for col in df.columns:
        key = re.sub(r"\s+", " ", str(col).strip().lower())
        if key in ALIASES:
            rename[col] = ALIASES[key]

    return df.rename(columns=rename)


def read_colleges(file):
    df = pd.read_excel(file)
    df = normalize_columns(df)

    if "College" not in df.columns:
        raise ValueError(
            "Add a college-name column such as College, "
            "College Name, Name, or Institution."
        )

    expected = [
        "City", "District", "State", "Website", "Wikipedia",
        "Address", "PIN Code", "Affiliation", "Status", "Type",
        "Latitude", "Longitude",
    ]

    for col in expected:
        if col not in df.columns:
            df[col] = ""

    df["College"] = df["College"].fillna("").astype(str).str.strip()
    df = df[df["College"] != ""]
    df = df.drop_duplicates(subset=["College"], keep="first")

    return df.reset_index(drop=True)


# -------------------- TIER CLASSIFICATION -----------------

def classify_one(row):
    """
    Project-defined rules only.
    The result is not an official ranking or accreditation.
    """
    status = text(row.get("Status", "")).lower()
    kind = text(row.get("Type", "")).lower()
    affiliation = text(row.get("Affiliation", "")).lower()

    evidence = f"{status} {kind} {affiliation}"

    if re.search(r"\bautonomous\b", evidence):
        return (
            "Tier 1",
            "Autonomous status indicated in the spreadsheet",
        )

    if re.search(r"\buniversit(?:y|ies)\b", evidence):
        return (
            "Tier 1",
            "University status indicated in the spreadsheet",
        )

    if "vtu" in evidence and "affiliat" in evidence:
        return (
            "Tier 1",
            "VTU affiliation indicated in the spreadsheet",
        )

    if evidence.strip() and evidence.strip() not in {"nan", "none"}:
        return (
            "Tier 2",
            "Institution details available; verify independently",
        )

    return (
        "Tier 3",
        "Insufficient institution or affiliation information",
    )


def classify_colleges(df):
    result = df.copy()
    pairs = result.apply(classify_one, axis=1)
    result["Tier"] = [p[0] for p in pairs]
    result["Reason"] = [p[1] for p in pairs]
    return result


def to_excel(df):
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="College Results")

    return buffer.getvalue()


# ---------------------- FREE GEOCODING --------------------

@st.cache_data(ttl=86400, show_spinner=False)
def lookup_location(query):
    """Best-effort OpenStreetMap Nominatim lookup."""
    if not query:
        return None

    try:
        response = requests.get(
            "https://nominatim.openstreetmap.org/search",
            params={
                "q": query,
                "format": "jsonv2",
                "limit": 1,
            },
            headers={
                "User-Agent": "CampusPulseCollegeExplorer/1.0"
            },
            timeout=12,
        )
        response.raise_for_status()
        results = response.json()

        if not results:
            return None

        item = results[0]
        return {
            "lat": float(item["lat"]),
            "lon": float(item["lon"]),
            "address": item.get("display_name", query),
        }
    except Exception:
        return None


def college_location(row):
    try:
        lat = float(row.get("Latitude", ""))
        lon = float(row.get("Longitude", ""))

        if -90 <= lat <= 90 and -180 <= lon <= 180:
            if lat != 0 and lon != 0:
                return {
                    "lat": lat,
                    "lon": lon,
                    "address": text(row.get("Address", "")),
                }
    except (ValueError, TypeError):
        pass

    parts = [
        text(row.get("College", "")),
        text(row.get("Address", "")),
        text(row.get("City", "")),
        text(row.get("District", "")),
        text(row.get("State", "")),
        text(row.get("PIN Code", "")),
    ]
    return lookup_location(", ".join(p for p in parts if p))


# ----------------------- RAG RETRIEVAL ---------------------

def row_to_document(row):
    fields = [
        "College", "City", "District", "State", "Address",
        "PIN Code", "Website", "Wikipedia", "Affiliation",
        "Status", "Type", "Tier", "Reason",
    ]

    return "\n".join(
        f"{key}: {text(row.get(key, ''))}"
        for key in fields
        if text(row.get(key, ""))
    )


def retrieve(question, df, limit=5):
    if df.empty:
        return []

    docs = [row_to_document(row) for _, row in df.iterrows()]

    if not docs or not any(docs):
        return []

    try:
        vectorizer = TfidfVectorizer(
            stop_words="english",
            ngram_range=(1, 2),
        )
        matrix = vectorizer.fit_transform(docs + [question])
        scores = cosine_similarity(
            matrix[-1], matrix[:-1]
        ).flatten()

        order = scores.argsort()[::-1][:limit]

        return [
            docs[i]
            for i in order
            if scores[i] > 0
        ]
    except ValueError:
        return []


def answer_question(question, df):
    documents = retrieve(question, df)

    if not documents:
        return (
            "I couldn't find a matching college record. "
            "Try a college name, city, affiliation, or website."
        )

    context = "\n\n---\n\n".join(documents)

    if not GROQ_API_KEY:
        return (
            "**Matching records from your Excel data**\n\n"
            "AI generation is not configured, so here are the retrieved "
            "records instead:\n\n" + context
        )

    try:
        from langchain_groq import ChatGroq
        from langchain_core.messages import SystemMessage, HumanMessage

        model = ChatGroq(
            model="openai/gpt-oss-120b",
            api_key=GROQ_API_KEY,
            temperature=0.2,
            max_tokens=1200,
        )

        response = model.invoke([
            SystemMessage(content=(
                "You are the CampusPulse college assistant. Answer using "
                "only the supplied college records. If the answer is "
                "missing, say so. Never invent addresses, affiliations, "
                "ratings, or reviews.\n\n"
                f"COLLEGE RECORDS:\n{context}"
            )),
            HumanMessage(content=question),
        ])

        return response.content

    except Exception as exc:
        return (
            "The records were retrieved, but AI generation failed. "
            f"Check your optional Groq configuration. Details: {exc}\n\n"
            + context
        )


# ---------------------- SESSION STATE ---------------------

for key, initial in {
    "colleges": pd.DataFrame(),
    "results": pd.DataFrame(),
    "upload_name": "",
    "chat": [],
}.items():
    if key not in st.session_state:
        st.session_state[key] = initial


# --------------------------- HEADER ------------------------

st.markdown("""
<div class="cp-header">
  <h1>🎓 CampusPulse</h1>
  <p>Live College Intelligence Platform</p>
  <p>Excel Extractor &nbsp; | &nbsp; Tier Classification &nbsp; | &nbsp;
  Live Location &nbsp; | &nbsp; Reviews &nbsp; | &nbsp; RAG Assistant</p>
  <p class="cp-brand">👑 Powered by: Shivasharanappa</p>
</div>
""", unsafe_allow_html=True)


# ------------------------- SIDEBAR -------------------------

with st.sidebar:
    st.markdown("## 🎓 CampusPulse")
    st.caption("College Intelligence Platform")

    page = st.radio(
        "Navigation",
        [
            "Dashboard",
            "Upload & Process",
            "Tier Results",
            "College Search",
            "College Map",
            "Reviews",
            "RAG Chatbot",
            "About",
        ],
    )

    st.divider()
    st.caption("Discover · Analyze · Compare · Choose Smarter")


# --------------------- UPLOAD & PREVIEW --------------------

if page in ["Dashboard", "Upload & Process"]:
    upload_col, preview_col = st.columns([0.9, 1.25], gap="medium")

    with upload_col:
        st.markdown(
            '<div class="cp-card"><div class="cp-heading">'
            '📗 1. Upload Excel File</div>'
            '<div class="cp-muted">Extract college names, cities, websites, etc.</div></div>',
            unsafe_allow_html=True,
        )

        uploaded = st.file_uploader(
            "Drag and drop your Excel file",
            type=["xlsx", "xls"],
            key="college_file",
        )

        if uploaded is not None:
            try:
                if st.session_state.upload_name != uploaded.name:
                    st.session_state.colleges = read_colleges(uploaded)
                    st.session_state.results = pd.DataFrame()
                    st.session_state.upload_name = uploaded.name
                    st.session_state.chat = []

                st.success(
                    f"{uploaded.name} · "
                    f"{len(st.session_state.colleges)} colleges loaded"
                )

            except Exception as exc:
                st.error(f"Excel error: {exc}")

        if st.button("Clear uploaded data", use_container_width=True):
            st.session_state.colleges = pd.DataFrame()
            st.session_state.results = pd.DataFrame()
            st.session_state.upload_name = ""
            st.session_state.chat = []
            st.rerun()

    with preview_col:
        df = st.session_state.colleges

        st.markdown(
            '<div class="cp-card"><div class="cp-heading">'
            '📋 2. Extracted Colleges (Preview)</div></div>',
            unsafe_allow_html=True,
        )

        st.metric("Total Colleges Found", len(df))

        if not df.empty:
            query = st.text_input(
                "Search college name, city or keyword",
                key="preview_query",
            )

            preview = df
            if query:
                mask = df.fillna("").astype(str).apply(
                    lambda col: col.str.contains(
                        query, case=False, regex=False
                    )
                ).any(axis=1)
                preview = df[mask]

            cols = [
                c for c in [
                    "College", "City", "District", "State",
                    "Website", "Wikipedia",
                ]
                if c in preview.columns
            ]

            st.dataframe(
                preview[cols],
                use_container_width=True,
                hide_index=True,
                height=260,
            )

            st.download_button(
                "⬇ Download Extracted Excel",
                data=to_excel(df),
                file_name="campuspulse_extracted.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        else:
            st.info("Upload an Excel file to display college records.")


# ---------------------- SELECT & PROCESS -------------------

if page in ["Dashboard", "Upload & Process"]:
    st.markdown(
        '<div class="cp-heading">✅ 3. Select Colleges & Process</div>',
        unsafe_allow_html=True,
    )

    df = st.session_state.colleges

    if not df.empty:
        names = df["College"].astype(str).tolist()

        selected = st.multiselect(
            "Search and select colleges",
            options=names,
            key="selected_colleges",
        )

        select_col, process_col = st.columns(2)

        with select_col:
            if st.button("✅ Select All", use_container_width=True):
                st.session_state.selected_colleges = names
                st.rerun()

        with process_col:
            if st.button(
                "▶ Process",
                type="primary",
                use_container_width=True,
            ):
                chosen_names = st.session_state.get(
                    "selected_colleges", []
                )

                if not chosen_names:
                    st.warning("Select one or more colleges first.")
                else:
                    chosen = df[
                        df["College"].astype(str).isin(chosen_names)
                    ].copy()

                    st.session_state.results = classify_colleges(chosen)
                    st.success(f"Processed {len(chosen)} colleges.")

        st.caption(
            "Classification follows configurable project rules; "
            "it is not an official academic ranking."
        )
    else:
        st.info("Upload Excel to enable selection and processing.")


# --------------------- TIER RESULT TABLES ------------------

if page in ["Dashboard", "Tier Results"]:
    st.markdown("## 📊 Tier Classification Results")

    results = st.session_state.results

    if not results.empty:
        columns = st.columns(3, gap="small")

        tier_configs = [
            (columns[0], "Tier 1", "🏅 Tier 1 Colleges", "cp-tier1"),
            (columns[1], "Tier 2", "🏆 Tier 2 Colleges", "cp-tier2"),
            (columns[2], "Tier 3", "🛡 Tier 3 Colleges", "cp-tier3"),
        ]

        for col, tier, label, css in tier_configs:
            subset = results[results["Tier"] == tier]

            with col:
                st.markdown(
                    f'<div class="{css}"><h4>{label}</h4>'
                    f'<b>{len(subset)} colleges</b></div>',
                    unsafe_allow_html=True,
                )

                visible = [
                    c for c in ["College", "City", "Type", "Reason"]
                    if c in subset.columns
                ]

                st.dataframe(
                    subset[visible],
                    use_container_width=True,
                    hide_index=True,
                    height=240,
                )

        st.download_button(
            "⬇ Download Classified Excel",
            data=to_excel(results),
            file_name="campuspulse_classified.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.info("Upload Excel and click Process to see tier tables.")


# ----------------------- COLLEGE DETAILS -------------------

if page in ["Dashboard", "College Search"]:
    if page == "College Search":
        st.header("🔎 College Search")
    else:
        st.markdown("## ℹ️ College Details")

    df = st.session_state.colleges

    if df.empty:
        st.info("Upload a college spreadsheet first.")
    else:
        query = st.text_input(
            "Find a college by name or city",
            key="detail_query",
        )

        matches = df

        if query:
            mask = df.fillna("").astype(str).apply(
                lambda col: col.str.contains(
                    query, case=False, regex=False
                )
            ).any(axis=1)
            matches = df[mask]

        if not matches.empty:
            college_name = st.selectbox(
                "Selected college",
                matches["College"].astype(str).tolist(),
                key="detail_college",
            )

            row = matches[
                matches["College"].astype(str) == college_name
            ].iloc[0]

            detail_col, links_col = st.columns([1.4, 1])

            with detail_col:
                fields = [
                    "College", "City", "District", "State", "Address",
                    "PIN Code", "Website", "Wikipedia", "Affiliation",
                    "Status", "Type",
                ]

                detail_rows = [
                    {"Field": field, "Value": text(row.get(field, ""))}
                    for field in fields
                    if text(row.get(field, ""))
                ]

                st.dataframe(
                    pd.DataFrame(detail_rows),
                    use_container_width=True,
                    hide_index=True,
                )

            with links_col:
                website = text(row.get("Website", ""))
                wiki = text(row.get("Wikipedia", ""))

                if website.startswith(("https://", "http://")):
                    st.link_button(
                        "🌐 Open Official Website",
                        website,
                        use_container_width=True,
                    )

                if wiki.startswith(("https://", "http://")):
                    st.link_button(
                        "📖 Open Wikipedia",
                        wiki,
                        use_container_width=True,
                    )

                maps_query = requests.utils.quote(
                    ", ".join(
                        x for x in [
                            college_name,
                            text(row.get("City", "")),
                            text(row.get("State", "")),
                        ] if x
                    )
                )

                st.link_button(
                    "📍 Search on OpenStreetMap",
                    f"https://www.openstreetmap.org/search?query={maps_query}",
                    use_container_width=True,
                )
        else:
            st.warning("No colleges match your search.")


# ------------------------- LOCATION MAP --------------------

if page in ["Dashboard", "College Map"]:
    if page == "College Map":
        st.header("📍 College Location Map")
    else:
        st.markdown("## 📍 College Location Map")

    df = st.session_state.colleges

    if df.empty:
        st.info("Upload college data to explore locations.")
    else:
        map_name = st.selectbox(
            "Choose a college to locate",
            df["College"].astype(str).tolist(),
            key="map_college",
        )

        row = df[df["College"].astype(str) == map_name].iloc[0]

        parts = [
            text(row.get("Address", "")),
            text(row.get("City", "")),
            text(row.get("District", "")),
            text(row.get("State", "")),
            text(row.get("PIN Code", "")),
        ]
        address = ", ".join(p for p in parts if p)

        st.write(f"**College:** {map_name}")
        st.write(f"**Address in your data:** {address or 'Not provided'}")

        if st.button("Find college location"):
            with st.spinner("Searching OpenStreetMap..."):
                location = college_location(row)

            if location:
                map_obj = folium.Map(
                    location=[location["lat"], location["lon"]],
                    zoom_start=15,
                )

                folium.Marker(
                    [location["lat"], location["lon"]],
                    tooltip=map_name,
                    popup=location.get("address", map_name),
                ).add_to(map_obj)

                st_folium(
                    map_obj,
                    use_container_width=True,
                    height=420,
                    key="campuspulse_map",
                )

                st.caption(
                    "Geocoding is approximate. Verify the pin against "
                    "the college's official address."
                )
            else:
                st.warning(
                    "Location not found. Add a complete address or "
                    "latitude and longitude to the Excel sheet."
                )


# ----------------------- REVIEWS PANEL ---------------------

if page in ["Dashboard", "Reviews"]:
    if page == "Reviews":
        st.header("⭐ Reviews & Ratings")
    else:
        st.markdown("## ⭐ Reviews & Ratings")

    df = st.session_state.colleges

    if df.empty:
        st.info("Upload college data first.")
    else:
        review_name = st.selectbox(
            "Choose a college",
            df["College"].astype(str).tolist(),
            key="review_college",
        )

        review_row = df[
            df["College"].astype(str) == review_name
        ].iloc[0]

        review_city = text(review_row.get("City", ""))
        review_query = requests.utils.quote(
            f"{review_name} {review_city} reviews"
        )

        st.markdown(
            "**Live review links**\n\n"
            "This version does not invent ratings or student reviews."
        )

        st.link_button(
            "⭐ Search Google reviews",
            f"https://www.google.com/search?q={review_query}",
            use_container_width=True,
        )

        st.link_button(
            "🗺 Search OpenStreetMap",
            f"https://www.openstreetmap.org/search?query={review_query}",
            use_container_width=True,
        )

        st.caption(
            "Google search opens public listings when available. "
            "This link does not retrieve or verify review scores inside "
            "CampusPulse. Direct in-dashboard reviews require a supported "
            "data provider."
        )


# ------------------------- RAG CHATBOT ---------------------

if page in ["Dashboard", "RAG Chatbot"]:
    if page == "RAG Chatbot":
        st.header("💬 RAG College Assistant")
    else:
        st.markdown("## 💬 RAG Chatbot — Ask about Colleges")

    df = st.session_state.colleges

    if df.empty:
        st.info("Upload an Excel sheet to create the searchable knowledge base.")
    else:
        st.caption(
            "Answers use records retrieved from your uploaded Excel data. "
            "Groq is optional."
        )

        examples = [
            "Show Tier 1 colleges in Bengaluru",
            "Which colleges are autonomous?",
            "Find VTU-affiliated colleges",
            "Compare two colleges",
        ]

        prompt_choice = st.selectbox(
            "Try a sample question",
            ["Choose a sample..."] + examples,
            key="rag_example",
        )

        with st.form("rag_form", clear_on_submit=True):
            question = st.text_input(
                "Ask your college question",
                value="" if prompt_choice == "Choose a sample..." else prompt_choice,
                placeholder="Example: Which colleges are in Bengaluru?",
            )

            submitted = st.form_submit_button("➤ Ask CampusPulse")

        if submitted and question.strip():
            with st.spinner("Searching college records..."):
                answer = answer_question(
                    question.strip(),
                    st.session_state.results
                    if not st.session_state.results.empty
                    else df,
                )

            st.session_state.chat.append({
                "question": question.strip(),
                "answer": answer,
            })

        if st.session_state.chat:
            for item in reversed(st.session_state.chat):
                with st.expander(
                    f"Q: {item['question']}",
                    expanded=True,
                ):
                    st.markdown(item["answer"])

            if st.button("Clear chat history"):
                st.session_state.chat = []
                st.rerun()


# --------------------------- ABOUT -------------------------

if page == "About":
    st.header("About CampusPulse")

    st.markdown("""
    **CampusPulse — Live College Intelligence Platform**

    - Excel college extraction
    - Configurable Tier 1 / Tier 2 / Tier 3 classification
    - College search and details
    - OpenStreetMap location lookup
    - Public review search links
    - RAG-style retrieval over uploaded college records
    - Optional Groq-powered answers

    **Powered by: Shivasharanappa**
    """)

    st.warning(
        "Tier labels are project-defined rules, not official rankings. "
        "Verify autonomy, affiliation, and institutional status with "
        "official sources before relying on a classification."
    )
