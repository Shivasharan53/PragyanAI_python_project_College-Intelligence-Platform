
import os
import re
from urllib.parse import quote_plus

import pandas as pd
import streamlit as st

# ============================================================
# CAMPUSPULSE | COLLEGE INTELLIGENCE PLATFORM
# ============================================================

st.set_page_config(
    page_title="CampusPulse",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

# -------------------- VISUAL DESIGN --------------------

st.markdown("""
<style>
.stApp {
    background: #f4f7fc;
    color: #172b4d;
}
.block-container {
    max-width: 1500px;
    padding: 1.3rem 1.6rem 2.5rem;
}
h1, h2, h3, h4, p, label, .stMarkdown {
    color: #172b4d;
}
h1 { font-size: 2rem !important; }
h2 { font-size: 1.55rem !important; }
h3 { font-size: 1.2rem !important; }

.cp-hero {
    background: linear-gradient(115deg, #102344, #174b91, #0878e8);
    border-radius: 16px;
    padding: 24px;
    margin-bottom: 20px;
}
.cp-hero h1 {
    color: #ffffff !important;
    font-size: 2.25rem !important;
    margin: 0;
}
.cp-hero p {
    color: #eaf3ff !important;
    margin: 6px 0 0;
}
.cp-hero .brand {
    color: #ffe08a !important;
    margin-top: 12px;
    font-weight: 700;
}
[data-testid="stMetric"] {
    background: #ffffff;
    border: 1px solid #d9e4f2;
    border-radius: 13px;
    padding: 16px;
    box-shadow: 0 2px 7px rgba(20, 55, 100, .04);
}
[data-testid="stMetricLabel"],
[data-testid="stMetricValue"],
[data-testid="stMetricDelta"] {
    color: #172b4d !important;
}
[data-testid="stSidebar"] {
    background: #122442;
}
[data-testid="stSidebar"] * {
    color: #f4f8ff;
}
.stButton > button,
.stDownloadButton > button,
.stLinkButton > a,
button[kind="primary"] {
    border-radius: 9px !important;
    font-weight: 650 !important;
    min-height: 42px;
}
div[data-testid="stDataFrame"] {
    border: 1px solid #d9e4f2;
    border-radius: 10px;
}
div[data-testid="stAlert"] {
    border-radius: 10px;
}
.cp-section {
    background: #ffffff;
    border: 1px solid #d9e4f2;
    border-radius: 13px;
    padding: 15px;
    margin: 8px 0 14px;
}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="cp-hero">
  <h1>🎓 Campus<span style="color:#54c6ff">Pulse</span></h1>
  <p>Live College Intelligence Platform</p>
  <p>Discover · Analyze · Compare · Choose Smarter</p>
  <p class="brand">Powered by: Shivasharanappa</p>
</div>
""", unsafe_allow_html=True)


# ============================================================
# SESSION STATE
# ============================================================

DEFAULTS = {
    "raw_df": None,
    "processed_df": None,
    "source_name": "",
    "source_sheet": "",
    "column_mapping": {},
    "chat_history": [],
    "last_processed_names": [],
}

for key, value in DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = value


# ============================================================
# EXCEL COLUMN DETECTION
# ============================================================

ALIASES = {
    "College Name": [
        "college name", "clg name", "college", "institution",
        "institution name", "institute", "institute name",
        "name of college", "college/institution"
    ],
    "City": ["city", "town", "place", "city name"],
    "District": ["district", "dist", "district name"],
    "Taluk": ["taluk", "tq", "tehsil", "taluka"],
    "State": ["state", "state name"],
    "Website": [
        "website", "official website", "college website",
        "web site", "homepage", "url"
    ],
    "Address": ["address", "full address", "postal address"],
    "Affiliation": [
        "affiliation", "affiliated university",
        "university", "affiliated to"
    ],
    "Status": [
        "status", "college status", "autonomous status",
        "recognition status"
    ],
    "Type": [
        "type", "college type", "institution type",
        "category", "ownership"
    ],
    "Wikipedia": ["wikipedia", "wiki", "wikipedia url"],
    "Rating": ["rating", "google rating", "star rating"],
    "Review Count": ["review count", "number of reviews", "reviews count"],
    "Latitude": ["latitude", "lat"],
    "Longitude": ["longitude", "lon", "lng", "long"],
    "Principal Name": ["principal name", "principal"],
    "Principal Email": ["principal email", "official email", "email"],
}


def normalize(value):
    return re.sub(r"[^a-z0-9]+", " ", str(value).lower()).strip()


def detect_columns(columns):
    detected = {}
    normalized_columns = {normalize(c): c for c in columns}

    for field, aliases in ALIASES.items():
        # Exact match first
        for alias in aliases:
            key = normalize(alias)
            if key in normalized_columns:
                detected[field] = normalized_columns[key]
                break

        # Partial match only if no exact match was found
        if field not in detected:
            for column in columns:
                for alias in aliases:
                    alias_key = normalize(alias)
                    if len(alias_key) >= 4 and alias_key in normalize(column):
                        detected[field] = column
                        break
                if field in detected:
                    break

    return detected


def read_excel_workbook(uploaded):
    """Read every worksheet and prefer one with a college-name column."""
    try:
        excel = pd.ExcelFile(uploaded)
        readable = []

        for sheet_name in excel.sheet_names:
            try:
                frame = pd.read_excel(excel, sheet_name=sheet_name)
                frame = frame.dropna(how="all").dropna(axis=1, how="all")
                frame.columns = [str(c).strip() for c in frame.columns]
                if not frame.empty and len(frame.columns):
                    readable.append((sheet_name, frame))
            except Exception:
                continue

        if not readable:
            return None, None, "No readable worksheet was found."

        for sheet_name, frame in readable:
            if "College Name" in detect_columns(frame.columns):
                return frame, sheet_name, "College-name column detected."

        sheet_name, frame = max(readable, key=lambda item: len(item[1]))
        return (
            frame,
            sheet_name,
            "Could not detect the college-name column automatically. "
            "Please select it manually."
        )

    except Exception as exc:
        return None, None, f"Could not read this Excel workbook: {exc}"


def is_blank(value):
    if pd.isna(value):
        return True
    return str(value).strip().lower() in ("", "nan", "none", "nat")


def get_value(row, column):
    if column not in row.index or is_blank(row[column]):
        return ""
    return str(row[column]).strip()


def prepare_colleges(raw, mapping):
    """Create standard field names while preserving the uploaded columns."""
    result = raw.copy()

    for standard_name, source_column in mapping.items():
        if source_column in raw.columns:
            result[standard_name] = raw[source_column]

    if "College Name" not in result.columns:
        return pd.DataFrame()

    result["College Name"] = (
        result["College Name"].fillna("").astype(str).str.strip()
    )
    result = result[
        (result["College Name"] != "")
        & (result["College Name"].str.lower() != "nan")
    ].copy()

    result = result.drop_duplicates(
        subset=["College Name"]
    ).reset_index(drop=True)

    for standard_name in ALIASES:
        if standard_name not in result.columns:
            result[standard_name] = ""

    return result


# ============================================================
# CLASSIFICATION RULES
# ============================================================

def classify_college(row):
    """
    Preliminary indicator-based classification.
    This is not an official college ranking.
    """
    name = get_value(row, "College Name").lower()
    status = get_value(row, "Status").lower()
    college_type = get_value(row, "Type").lower()
    affiliation = get_value(row, "Affiliation").lower()

    combined = f"{name} {status} {college_type} {affiliation}"

    # Tier 1: explicit autonomous/university/VTU indicators
    tier1_terms = [
        "autonomous",
        "university",
        "vtu",
        "visvesvaraya technological university",
    ]

    if any(term in combined for term in tier1_terms):
        return (
            "Tier 1",
            "Autonomous, university or VTU indicator found; verify official status"
        )

    # Tier 2: affiliation or recognition indicators
    tier2_terms = [
        "affiliated", "recognised", "recognized",
        "approved", "accredited", "affiliation"
    ]

    if any(term in combined for term in tier2_terms):
        return (
            "Tier 2",
            "Affiliation or recognition indicator found; verify official details"
        )

    return (
        "Tier 3",
        "Needs verification: no matching classification indicator found"
    )


def apply_classification(frame):
    result = frame.copy()
    classifications = result.apply(classify_college, axis=1)
    result["Tier"] = [item[0] for item in classifications]
    result["Classification Reason"] = [item[1] for item in classifications]
    return result


# ============================================================
# COLLEGE LINKS
# ============================================================

def official_website(row):
    website = get_value(row, "Website")
    if not website:
        return ""
    if not website.startswith(("https://", "http://")):
        website = "https://" + website
    return website


def maps_link(row):
    latitude = get_value(row, "Latitude")
    longitude = get_value(row, "Longitude")

    if latitude and longitude:
        try:
            return (
                "https://www.google.com/maps?q="
                f"{float(latitude)},{float(longitude)}"
            )
        except ValueError:
            pass

    query = " ".join(
        value for value in [
            get_value(row, "College Name"),
            get_value(row, "City"),
            get_value(row, "District"),
            get_value(row, "State"),
        ] if value
    )
    return (
        "https://www.google.com/maps/search/?api=1&query="
        + quote_plus(query)
    )


def reviews_link(row):
    query = " ".join(
        value for value in [
            get_value(row, "College Name"),
            get_value(row, "City"),
            "college reviews ratings",
        ] if value
    )
    return "https://www.google.com/search?q=" + quote_plus(query)


def dataframe_csv(frame):
    return frame.to_csv(index=False).encode("utf-8-sig")


# ============================================================
# RAG RETRIEVAL + OPTIONAL GROQ
# ============================================================

def retrieve_records(question, frame, limit=8):
    if frame is None or frame.empty:
        return pd.DataFrame()

    records = frame.fillna("").astype(str)
    documents = records.apply(
        lambda row: " | ".join(
            f"{column}: {row[column]}" for column in records.columns
        ),
        axis=1,
    ).tolist()

    # First try TF-IDF retrieval. Fallback requires no extra package.
    try:
        from sklearn.feature_extraction.text import TfidfVectorizer
        from sklearn.metrics.pairwise import cosine_similarity

        vectorizer = TfidfVectorizer(
            stop_words="english",
            ngram_range=(1, 2),
        )
        matrix = vectorizer.fit_transform(documents + [question])
        scores = cosine_similarity(matrix[-1], matrix[:-1]).ravel()
        indices = scores.argsort()[::-1]
        indices = [i for i in indices if scores[i] > 0][:limit]
        return frame.iloc[indices].copy()

    except (ImportError, ValueError):
        terms = [
            word for word in normalize(question).split()
            if len(word) > 2
        ]
        ranked = []

        for index, document in enumerate(documents):
            score = sum(term in document.lower() for term in terms)
            if score > 0:
                ranked.append((score, index))

        ranked.sort(reverse=True)
        return frame.iloc[
            [index for _, index in ranked[:limit]]
        ].copy()


def answer_question(question, frame):
    matches = retrieve_records(question, frame)

    if matches.empty:
        return (
            "I couldn't find matching records in the uploaded workbook. "
            "Try a college name, city, district, website or affiliation."
        ), matches

    # JSON-like record text avoids a dependency on tabulate.
    evidence = matches.fillna("").astype(str).to_dict(orient="records")

    try:
        api_key = st.secrets.get(
            "GROQ_API_JNTU_KEY",
            os.getenv("GROQ_API_JNTU_KEY", ""),
        ).strip()

        if api_key:
            from groq import Groq

            client = Groq(api_key=api_key)
            response = client.chat.completions.create(
                model="llama-3.3-70b-versatile",
                temperature=0.2,
                messages=[
                    {
                        "role": "system",
                        "content": (
                            "You are CampusPulse, a college information "
                            "assistant. Answer using only the retrieved "
                            "Excel records. If information is missing, say "
                            "it is not available. Never invent official "
                            "rankings, ratings, reviews, fees or facts."
                        ),
                    },
                    {
                        "role": "user",
                        "content": (
                            f"Question: {question}\n\n"
                            f"Retrieved records:\n{evidence}"
                        ),
                    },
                ],
            )
            return response.choices[0].message.content, matches

    except Exception:
        # Retrieval remains useful even if Groq is unavailable.
        pass

    return (
        "I found matching records in your Excel workbook. "
        "Review the evidence table below. To enable generated AI answers, "
        "check your Groq secret and the Groq dependency."
    ), matches


# ============================================================
# SIDEBAR: ONLY TWO PAGES
# ============================================================

with st.sidebar:
    st.markdown("## 🎓 CampusPulse")
    st.caption("COLLEGE INTELLIGENCE")

    page = st.radio(
        "Navigation",
        ["Dashboard", "Upload & Process"],
        label_visibility="collapsed",
    )

    current_df = st.session_state.processed_df
    processed_count = len(current_df) if current_df is not None else 0

    st.divider()
    st.metric("Processed Colleges", processed_count)
    st.markdown("**Powered by: Shivasharanappa**")


# ============================================================
# DASHBOARD: SIMPLE OVERVIEW ONLY
# ============================================================

if page == "Dashboard":
    st.title("Dashboard")
    st.caption("A simple overview of your processed college data.")

    df = st.session_state.processed_df

    if df is None or df.empty:
        st.info(
            "No processed colleges yet. Open **Upload & Process**, upload "
            "your Excel workbook, map the columns, select colleges and "
            "click Process. Your overview will update here."
        )
        total = 0
        cities = 0
        websites = 0
    else:
        total = len(df)

        city_values = df["City"].fillna("").astype(str).str.strip()
        cities = city_values[city_values != ""].nunique()

        website_values = df["Website"].fillna("").astype(str).str.strip()
        websites = website_values[
            (website_values != "")
            & (website_values.str.lower() != "nan")
        ].nunique()

        st.success(f"Current dataset: {st.session_state.source_name or 'Uploaded workbook'}")

    col1, col2, col3 = st.columns(3)
    col1.metric("Total Processed Colleges", total)
    col2.metric("Cities Covered", cities)
    col3.metric("Websites Available", websites)

    st.divider()
    st.subheader("Quick Start")

    quick1, quick2 = st.columns(2)

    with quick1:
        st.markdown("""
        <div class="cp-section">
            <h3>📤 Upload & Process</h3>
            <p>Upload a workbook, map columns, select colleges and process
            the results.</p>
        </div>
        """, unsafe_allow_html=True)

        if st.button(
            "Go to Upload & Process",
            type="primary",
            use_container_width=True,
        ):
            st.session_state.go_upload = True
            st.rerun()

    with quick2:
        st.markdown("""
        <div class="cp-section">
            <h3>📚 Current Dataset</h3>
            <p>Review the status of your processed workbook and the number
            of records available to explore.</p>
        </div>
        """, unsafe_allow_html=True)

        if df is not None and not df.empty:
            st.write(f"**Workbook:** {st.session_state.source_name}")
            st.write(f"**Records:** {len(df)}")
        else:
            st.write("Upload and process a workbook to populate this section.")

    # Use the quick-start action to switch the radio selection.
    if st.session_state.pop("go_upload", False):
        st.info("Choose **Upload & Process** in the sidebar to continue.")


# ============================================================
# UPLOAD & PROCESS: ALL WORKFLOW AND RESULTS
# ============================================================

elif page == "Upload & Process":
    st.title("Upload & Process")
    st.caption(
        "Upload Excel → map columns → select colleges → process → "
        "view results, college details and RAG answers."
    )

    uploaded = st.file_uploader(
        "1. Upload Excel workbook",
        type=["xlsx", "xls"],
        help="Works with different worksheet names and common column-name variations.",
        key="excel_upload",
    )

    raw = None
    sheet_name = None
    mapping = {}

    if uploaded is not None:
        raw, sheet_name, message = read_excel_workbook(uploaded)

        if raw is None:
            st.error(message)
        else:
            st.session_state.raw_df = raw
            st.session_state.source_name = uploaded.name
            st.session_state.source_sheet = sheet_name

            detected = detect_columns(raw.columns)
            st.success(
                f"{message} Worksheet: {sheet_name}. "
                f"{len(raw)} rows and {len(raw.columns)} columns detected."
            )

            st.subheader("2. Detect or map columns")

            default_name = detected.get("College Name", raw.columns[0])
            name_column = st.selectbox(
                "Which column contains college names?",
                options=list(raw.columns),
                index=list(raw.columns).index(default_name),
                key="college_name_column",
            )
            mapping["College Name"] = name_column

            with st.expander("Optional: map additional columns", expanded=True):
                map_fields = [
                    "City", "District", "Taluk", "State", "Website",
                    "Address", "Affiliation", "Status", "Type", "Wikipedia",
                    "Rating", "Review Count", "Latitude", "Longitude",
                    "Principal Name", "Principal Email",
                ]

                map_col1, map_col2 = st.columns(2)

                for index, field in enumerate(map_fields):
                    options = ["— Not available —"] + list(raw.columns)
                    detected_col = detected.get(field)
                    default_index = (
                        options.index(detected_col)
                        if detected_col in options else 0
                    )

                    target = map_col1 if index % 2 == 0 else map_col2
                    with target:
                        selected_column = st.selectbox(
                            field,
                            options=options,
                            index=default_index,
                            key=f"map_{field}",
                        )

                    if selected_column != "— Not available —":
                        mapping[field] = selected_column

            preview = prepare_colleges(raw, mapping)

            st.subheader("Excel Preview")
            st.dataframe(
                preview.head(8),
                use_container_width=True,
                hide_index=True,
            )

            st.subheader("3. Search and select colleges")

            search_colleges = st.text_input(
                "Search colleges by name",
                placeholder="Type a college name to filter the list...",
                key="select_college_search",
            )

            all_names = preview["College Name"].tolist()
            if search_colleges.strip():
                available_names = [
                    name for name in all_names
                    if search_colleges.lower() in name.lower()
                ]
            else:
                available_names = all_names

            st.caption(
                f"{len(available_names)} colleges available in the selection list."
            )

            select_col, clear_col = st.columns(2)

            with select_col:
                if st.button(
                    "Select all visible colleges",
                    use_container_width=True,
                ):
                    st.session_state.selected_college_names = available_names

            with clear_col:
                if st.button(
                    "Clear selection",
                    use_container_width=True,
                ):
                    st.session_state.selected_college_names = []

            # Initialize selection after upload or when the workbook changes.
            selection_key = "selected_college_names"
            if (
                selection_key not in st.session_state
                or st.session_state.get("selection_source") != uploaded.name
            ):
                st.session_state[selection_key] = available_names.copy()
                st.session_state.selection_source = uploaded.name

            valid_defaults = [
                name for name in st.session_state[selection_key]
                if name in available_names
            ]

            selected_names = st.multiselect(
                "Choose one or more colleges to process",
                options=available_names,
                default=valid_defaults,
                key="college_multiselect",
            )

            st.session_state[selection_key] = selected_names

            st.subheader("4. Process selected colleges")

            if st.button(
                "⚙️ Process Colleges",
                type="primary",
                use_container_width=True,
            ):
                if not selected_names:
                    st.warning("Select at least one college before processing.")
                else:
                    selected_df = preview[
                        preview["College Name"].isin(selected_names)
                    ].copy()

                    st.session_state.processed_df = apply_classification(
                        selected_df
                    )
                    st.session_state.last_processed_names = selected_names
                    st.session_state.column_mapping = mapping
                    st.session_state.chat_history = []

                    st.success(
                        f"Successfully processed {len(selected_df)} colleges."
                    )

    # --------------------------------------------------------
    # RESULTS: ONLY DISPLAY AFTER PROCESSING
    # --------------------------------------------------------

    result = st.session_state.processed_df

    if result is not None and not result.empty:
        st.divider()
        st.header("5. Processing Results")

        st.caption(
            "Tier rules are preliminary indicators based on available "
            "data, not official rankings. Tier 3 means needs verification."
        )

        tier1 = result[result["Tier"] == "Tier 1"].copy()
        tier2 = result[result["Tier"] == "Tier 2"].copy()
        tier3 = result[result["Tier"] == "Tier 3"].copy()

        m1, m2, m3 = st.columns(3)
        m1.metric("Tier 1 Colleges", len(tier1))
        m2.metric("Tier 2 Colleges", len(tier2))
        m3.metric("Tier 3 / Needs Verification", len(tier3))

        # Classification tables
        st.subheader("Classification Tables")

        tier_tab1, tier_tab2, tier_tab3 = st.tabs([
            f"Tier 1 ({len(tier1)})",
            f"Tier 2 ({len(tier2)})",
            f"Tier 3 ({len(tier3)})",
        ])

        display_columns = [
            column for column in [
                "College Name", "City", "District", "Taluk", "State",
                "Type", "Status", "Affiliation", "Tier",
                "Classification Reason",
            ] if column in result.columns
        ]

        with tier_tab1:
            if tier1.empty:
                st.info("No records matched the Tier 1 indicators.")
            else:
                st.dataframe(
                    tier1[display_columns],
                    use_container_width=True,
                    hide_index=True,
                )
            st.download_button(
                "Download Tier 1 CSV",
                dataframe_csv(tier1),
                "tier1_colleges.csv",
                "text/csv",
                key="download_tier1",
            )

        with tier_tab2:
            if tier2.empty:
                st.info("No records matched the Tier 2 indicators.")
            else:
                st.dataframe(
                    tier2[display_columns],
                    use_container_width=True,
                    hide_index=True,
                )
            st.download_button(
                "Download Tier 2 CSV",
                dataframe_csv(tier2),
                "tier2_colleges.csv",
                "text/csv",
                key="download_tier2",
            )

        with tier_tab3:
            if tier3.empty:
                st.info("No records need verification under these rules.")
            else:
                st.dataframe(
                    tier3[display_columns],
                    use_container_width=True,
                    hide_index=True,
                )
            st.download_button(
                "Download Tier 3 CSV",
                dataframe_csv(tier3),
                "tier3_colleges.csv",
                "text/csv",
                key="download_tier3",
            )

        st.download_button(
            "⬇️ Download all processed colleges",
            dataframe_csv(result),
            "campuspulse_processed_colleges.csv",
            "text/csv",
            use_container_width=True,
        )

        # Charts
        st.divider()
        st.subheader("📊 College Analytics")

        chart_col1, chart_col2 = st.columns(2)

        with chart_col1:
            st.markdown("**Colleges by classification**")
            counts = result["Tier"].value_counts().reindex(
                ["Tier 1", "Tier 2", "Tier 3"], fill_value=0
            )
            st.bar_chart(counts)

        with chart_col2:
            st.markdown("**Colleges by city**")
            if "City" in result.columns:
                cities = result["City"].fillna("").astype(str).str.strip()
                cities = cities[cities != ""]
                city_counts = cities.value_counts().head(10)
                if city_counts.empty:
                    st.info("No city information was supplied or mapped.")
                else:
                    st.bar_chart(city_counts)
            else:
                st.info("Map a City column to display this chart.")

        # Selected college details
        st.divider()
        st.subheader("🏫 Selected College Details")

        names = result["College Name"].astype(str).tolist()
        selected_detail = st.selectbox(
            "Choose a processed college",
            names,
            key="processed_detail_college",
        )
        row = result[result["College Name"] == selected_detail].iloc[0]

        detail_left, detail_right = st.columns([1.5, 1])

        with detail_left:
            detail_items = []
            for field in [
                "College Name", "City", "District", "Taluk", "State",
                "Address", "Type", "Status", "Affiliation", "Website",
                "Principal Name", "Principal Email", "Tier",
                "Classification Reason",
            ]:
                value = get_value(row, field)
                if value:
                    detail_items.append({"Field": field, "Value": value})

            st.dataframe(
                pd.DataFrame(detail_items),
                use_container_width=True,
                hide_index=True,
            )

        with detail_right:
            website = official_website(row)

            if website:
                st.link_button(
                    "🌐 Official Website",
                    website,
                    use_container_width=True,
                )
            else:
                st.info("No website supplied for this college.")

            st.link_button(
                "📍 Open Google Maps",
                maps_link(row),
                use_container_width=True,
            )

            st.link_button(
                "⭐ Search Public Reviews",
                reviews_link(row),
                use_container_width=True,
            )

            rating = get_value(row, "Rating")
            review_count = get_value(row, "Review Count")

            if rating:
                st.metric("Rating in uploaded data", rating)
            else:
                st.caption("No rating supplied in the workbook.")

            if review_count:
                st.caption(f"Review count in uploaded data: {review_count}")

            st.caption(
                "Public review links open external search results. "
                "Live reviews are not fetched automatically."
            )

        # ----------------------------------------------------
        # OPEN-ENDED RAG CHATBOT
        # ----------------------------------------------------

        st.divider()
        st.header("🤖 CampusPulse RAG Chatbot")
        st.write(
            "Type any question about the processed colleges. "
            "Suggested topics are optional; you can ask your own question."
        )

        with st.expander("Example questions (optional)"):
            st.write(
                "- Tell me about a particular college.\n"
                "- Find colleges in a city or district.\n"
                "- Which colleges have websites?\n"
                "- Compare two colleges using the available fields.\n"
                "- Which records are missing affiliation information?"
            )

        # Display previous conversation
        for message in st.session_state.chat_history:
            with st.chat_message(message["role"]):
                st.markdown(message["content"])

        # Use a form so the question input and button remain visible.
        with st.form("campuspulse_rag_form", clear_on_submit=True):
            user_question = st.text_input(
                "Type your question",
                placeholder="Ask anything about the processed college data...",
            )
            ask = st.form_submit_button(
                "🔎 Ask CampusPulse",
                type="primary",
                use_container_width=True,
            )

        if ask:
            if not user_question.strip():
                st.warning("Please type a question first.")
            else:
                st.session_state.chat_history.append({
                    "role": "user",
                    "content": user_question,
                })

                answer, evidence = answer_question(user_question, result)

                st.session_state.chat_history.append({
                    "role": "assistant",
                    "content": answer,
                })

                st.rerun()

        # Show the latest retrieved evidence after submission.
        if (
            st.session_state.chat_history
            and st.session_state.chat_history[-1]["role"] == "assistant"
        ):
            latest_answer = st.session_state.chat_history[-1]["content"]
            st.markdown("### Latest answer")
            st.markdown(latest_answer)

            # Retrieve matching rows again to display the evidence.
            latest_question = None
            for message in reversed(st.session_state.chat_history[:-1]):
                if message["role"] == "user":
                    latest_question = message["content"]
                    break

            if latest_question:
                latest_evidence = retrieve_records(latest_question, result)
                if not latest_evidence.empty:
                    st.markdown("**Retrieved Excel evidence**")
                    st.dataframe(
                        latest_evidence,
                        use_container_width=True,
                        hide_index=True,
                    )

        if st.button("🧹 Clear chat", key="clear_campus_pulse_chat"):
            st.session_state.chat_history = []
            st.rerun()

    else:
        st.info(
            "Upload an Excel workbook, map the columns, select colleges "
            "and click **Process Colleges**. The tier results, charts, "
            "college details and RAG chatbot will appear here."
        )
