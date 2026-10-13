
import os
import re
from urllib.parse import quote_plus

import pandas as pd
import streamlit as st

# ============================================================
# CAMPUSPULSE — LIVE COLLEGE INTELLIGENCE PLATFORM
# Powered by Shivasharanappa
# ============================================================

st.set_page_config(
    page_title="CampusPulse | College Intelligence",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --------------------------- DESIGN ---------------------------

st.markdown("""
<style>
.stApp {
    background: #f4f8ff;
}
.block-container {
    max-width: 1600px;
    padding-top: 1.1rem;
    padding-bottom: 2rem;
}
.cp-hero {
    background: linear-gradient(115deg,#101d3c,#174b98,#0878e8);
    border-radius: 18px;
    padding: 24px;
    color: white;
    margin-bottom: 18px;
}
.cp-hero h1 {
    color: white;
    margin: 0;
    font-size: 2.4rem;
    font-weight: 850;
}
.cp-hero p {
    color: #e6f1ff;
    margin: 5px 0 0;
}
.cp-brand {
    background: #102c58;
    border: 1px solid #69b6ff;
    border-radius: 12px;
    padding: 13px;
    color: white;
    text-align: center;
}
.cp-brand b {
    color: #ffd44d;
}
.cp-panel {
    border: 1px solid #d8e5f7;
    border-radius: 13px;
    padding: 15px;
    background: white;
}
div[data-testid="stMetric"] {
    background: white;
    border: 1px solid #dbe7f6;
    padding: 12px;
    border-radius: 12px;
}
.stButton > button, .stLinkButton > a {
    border-radius: 9px;
    font-weight: 600;
}
[data-testid="stSidebar"] {
    background: linear-gradient(180deg,#10203d,#172c4b);
}
[data-testid="stSidebar"] * {
    color: #f2f7ff;
}
</style>
""", unsafe_allow_html=True)


# --------------------------- COLUMN ALIASES ---------------------------

ALIASES = {
    "College Name": [
        "college name", "clg name", "college", "institution",
        "institution name", "institute", "institute name",
        "name of college", "college/institution"
    ],
    "City": ["city", "town", "place", "city name"],
    "District": ["district", "dist", "district name"],
    "Taluk": ["taluk", "tq", "tehsil", "taluka", "sub district"],
    "State": ["state", "state name"],
    "Website": [
        "website", "official website", "college website",
        "web site", "homepage", "url"
    ],
    "Wikipedia": ["wikipedia", "wikipedia url", "wiki"],
    "Address": ["address", "full address", "postal address"],
    "Affiliation": [
        "affiliation", "affiliated university", "university",
        "affiliated to"
    ],
    "Status": ["status", "college status", "autonomous status"],
    "Type": [
        "type", "college type", "institution type",
        "category", "ownership"
    ],
    "Principal Name": ["principal name", "principal"],
    "Principal Email": ["principal email", "official email", "email"],
    "Rating": ["rating", "google rating", "review rating", "star rating"],
    "Review Count": ["review count", "number of reviews", "reviews count"],
    "Latitude": ["latitude", "lat"],
    "Longitude": ["longitude", "lon", "lng", "long"],
}


# --------------------------- HELPERS ---------------------------

def norm(value):
    return re.sub(
        r"[^a-z0-9]+", " ", str(value).strip().lower()
    ).strip()


def blank(value):
    if pd.isna(value):
        return True
    return str(value).strip().lower() in ("", "nan", "none", "nat")


def value_of(row, column):
    if column not in row.index or blank(row[column]):
        return ""
    return str(row[column]).strip()


def detect_mapping(columns):
    mapping = {}
    normalized = {norm(c): c for c in columns}

    for field, aliases in ALIASES.items():
        for alias in aliases:
            if norm(alias) in normalized:
                mapping[field] = normalized[norm(alias)]
                break

        if field not in mapping:
            for column in columns:
                cn = norm(column)
                for alias in aliases:
                    an = norm(alias)
                    if an and len(an) >= 4 and an in cn:
                        mapping[field] = column
                        break
                if field in mapping:
                    break

    return mapping


def read_workbook(uploaded):
    if uploaded is None:
        return None, None, "Upload an Excel workbook to begin."

    try:
        excel = pd.ExcelFile(uploaded)
        candidates = []

        for sheet in excel.sheet_names:
            try:
                frame = pd.read_excel(excel, sheet_name=sheet)
                frame = frame.dropna(how="all").dropna(axis=1, how="all")
                frame.columns = [str(c).strip() for c in frame.columns]

                mapping = detect_mapping(frame.columns)
                if "College Name" in mapping:
                    candidates.append(
                        (len(frame), sheet, frame, mapping)
                    )
            except Exception:
                continue

        if candidates:
            _, sheet, frame, mapping = max(
                candidates, key=lambda item: item[0]
            )
            return frame, sheet, (
                f"Detected worksheet '{sheet}'. "
                f"Found {len(frame)} rows."
            )

        # If automatic detection fails, let the user map columns manually.
        for sheet in excel.sheet_names:
            try:
                frame = pd.read_excel(excel, sheet_name=sheet)
                frame = frame.dropna(how="all").dropna(axis=1, how="all")
                if not frame.empty:
                    frame.columns = [str(c).strip() for c in frame.columns]
                    return frame, sheet, (
                        "Could not automatically identify the college-name "
                        "column. Select it manually below."
                    )
            except Exception:
                continue

        return None, None, "No readable worksheet was found."

    except Exception as exc:
        return None, None, f"Could not read the workbook: {exc}"


def clean_colleges(frame, manual_mapping):
    result = frame.copy()

    name_column = manual_mapping.get("College Name")
    if name_column not in result.columns:
        return pd.DataFrame()

    for logical, source in manual_mapping.items():
        if source in result.columns:
            result[logical] = result[source]

    result["College Name"] = (
        result[name_column].fillna("").astype(str).str.strip()
    )
    result = result[
        (result["College Name"] != "")
        & (result["College Name"].str.lower() != "nan")
    ].copy()

    result = result.drop_duplicates(
        subset=["College Name"]
    ).reset_index(drop=True)

    # Add missing standard columns so the app can handle different files.
    for column in ALIASES:
        if column not in result.columns:
            result[column] = ""

    return result


def classify_row(row):
    """
    Preliminary heuristic classification, not an official ranking.
    Missing information is treated as uncertain, not as proof of low quality.
    """
    status = value_of(row, "Status").lower()
    typ = value_of(row, "Type").lower()
    affiliation = value_of(row, "Affiliation").lower()

    if any(word in status for word in [
        "autonomous", "deemed university", "deemed-to-be university"
    ]) or "university" in typ:
        return "Tier 1", "Autonomous / university indicator"

    if any(word in status for word in [
        "affiliated", "approved", "recognized", "recognised"
    ]) or any(word in affiliation for word in [
        "vtu", "visvesvaraya", "affiliated", "university"
    ]):
        return "Tier 2", "Affiliation / approval information found"

    return "Tier 3", "Needs verification / limited classification data"


def prepare_tiers(frame):
    result = frame.copy()
    pairs = result.apply(classify_row, axis=1)
    result["Tier"] = [p[0] for p in pairs]
    result["Reason"] = [p[1] for p in pairs]
    return result


def website_url(row):
    website = value_of(row, "Website")
    if not website:
        return ""
    if not website.startswith(("http://", "https://")):
        website = "https://" + website
    return website


def maps_url(row):
    lat = value_of(row, "Latitude")
    lon = value_of(row, "Longitude")

    if lat and lon:
        try:
            return (
                "https://www.google.com/maps?q="
                f"{float(lat)},{float(lon)}"
            )
        except ValueError:
            pass

    query = " ".join(
        x for x in [
            value_of(row, "College Name"),
            value_of(row, "City"),
            value_of(row, "State"),
        ] if x
    )
    return (
        "https://www.google.com/maps/search/?api=1&query="
        + quote_plus(query)
    )


def reviews_url(row):
    query = " ".join(
        x for x in [
            value_of(row, "College Name"),
            value_of(row, "City"),
            "reviews ratings",
        ] if x
    )
    return "https://www.google.com/search?q=" + quote_plus(query)


def wikipedia_url(row):
    supplied = value_of(row, "Wikipedia")
    if supplied.startswith(("http://", "https://")):
        return supplied
    return (
        "https://en.wikipedia.org/w/index.php?search="
        + quote_plus(value_of(row, "College Name"))
    )


def csv_bytes(frame):
    return frame.to_csv(index=False).encode("utf-8-sig")


# --------------------------- RAG RETRIEVAL ---------------------------

def retrieve_rows(question, frame, top_k=8):
    if frame is None or frame.empty:
        return pd.DataFrame()

    documents = frame.fillna("").astype(str).apply(
        lambda row: " | ".join(
            f"{column}: {row[column]}"
            for column in frame.columns
        ),
        axis=1,
    ).tolist()

    try:
        from sklearn.feature_extraction.text import TfidfVectorizer
        from sklearn.metrics.pairwise import cosine_similarity

        vectorizer = TfidfVectorizer(
            ngram_range=(1, 2),
            stop_words="english",
        )
        matrix = vectorizer.fit_transform(documents + [question])
        scores = cosine_similarity(matrix[-1], matrix[:-1]).ravel()
        indices = scores.argsort()[::-1]
        indices = [i for i in indices if scores[i] > 0][:top_k]

        return frame.iloc[indices].copy()

    except Exception:
        terms = [t for t in norm(question).split() if len(t) > 2]
        scores = []

        for index, text in enumerate(documents):
            score = sum(term in text.lower() for term in terms)
            if score:
                scores.append((score, index))

        scores.sort(reverse=True)
        return frame.iloc[
            [index for _, index in scores[:top_k]]
        ].copy()


def answer_question(question, frame):
    question_lower = question.lower().strip()

    if frame is None or frame.empty:
        return (
            "Please upload and process a college Excel workbook first."
        )

    # Handle direct tier questions.
    requested_tier = None
    if re.search(r"\btier[\s-]*1\b|\btier one\b", question_lower):
        requested_tier = "Tier 1"
    elif re.search(r"\btier[\s-]*2\b|\btier two\b", question_lower):
        requested_tier = "Tier 2"
    elif re.search(r"\btier[\s-]*3\b|\btier three\b", question_lower):
        requested_tier = "Tier 3"

    if requested_tier:
        tiered = prepare_tiers(frame)
        matches = tiered[tiered["Tier"] == requested_tier]
        if matches.empty:
            return f"No records are currently classified as {requested_tier}."
        return (
            f"Found {len(matches)} records classified as {requested_tier}.\n\n"
            + matches[
                ["College Name", "City", "District", "Tier", "Reason"]
            ].head(20).to_markdown(index=False)
            + "\n\n*Classification is preliminary, not an official ranking.*"
        )

    # Handle city questions using the City field when available.
    city_terms = {
        "bengaluru": ["bengaluru", "bangalore"],
        "bangalore": ["bengaluru", "bangalore"],
        "mysuru": ["mysuru", "mysore"],
        "mysore": ["mysuru", "mysore"],
    }

    for city, variants in city_terms.items():
        if city in question_lower:
            city_values = frame["City"].fillna("").astype(str).str.lower()
            mask = city_values.apply(
                lambda value: any(variant in value for variant in variants)
            )
            matches = frame[mask]
            if not matches.empty:
                return (
                    f"Found {len(matches)} records matching {city.title()}.\n\n"
                    + matches[
                        ["College Name", "City", "District", "Website"]
                    ].head(20).to_markdown(index=False)
                )

    matches = retrieve_rows(question, frame)

    if matches.empty:
        return (
            "I couldn't find matching information in the uploaded data. "
            "Try a college name, city, district, website or affiliation."
        )

    context = matches.fillna("").astype(str).to_markdown(index=False)

    # Groq is optional; use it only when the secret is configured.
    try:
        api_key = st.secrets.get(
            "GROQ_API_JNTU_KEY",
            os.getenv("GROQ_API_JNTU_KEY", ""),
        ).strip()

        if api_key:
            from groq import Groq

            client = Groq(api_key=api_key)
            response = client.chat.completions.create(
                model=os.getenv(
                    "GROQ_MODEL", "llama-3.3-70b-versatile"
                ),
                messages=[
                    {
                        "role": "system",
                        "content": (
                            "You are CampusPulse, a college information "
                            "assistant. Answer only from the supplied "
                            "spreadsheet evidence. If information is missing, "
                            "say so. Never invent ratings, reviews, fees, "
                            "official rankings, accreditation or facts."
                        ),
                    },
                    {
                        "role": "user",
                        "content": (
                            f"Question: {question}\n\n"
                            f"Spreadsheet evidence:\n{context}"
                        ),
                    },
                ],
                temperature=0.2,
            )

            return (
                response.choices[0].message.content
                + "\n\n### Retrieved evidence\n\n"
                + context
            )

    except Exception:
        # Still return evidence if the AI service is unavailable.
        pass

    return (
        "### Matching records from your workbook\n\n"
        + context
        + "\n\n*Configure GROQ_API_JNTU_KEY to enable generated AI answers.*"
    )


# --------------------------- SESSION STATE ---------------------------

DEFAULTS = {
    "raw_df": None,
    "processed_df": None,
    "source_name": "",
    "sheet_name": "",
    "rag_messages": [],
    "manual_mapping": {},
    "dashboard_search": "",
}

for key, default in DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = default


# --------------------------- SIDEBAR ---------------------------

with st.sidebar:
    st.markdown("## 🎓 CampusPulse")
    st.caption("NAVIGATION")

    page = st.radio(
        "Navigation",
        ["Dashboard", "Upload & Process"],
        label_visibility="collapsed",
    )

    st.divider()

    count = (
        len(st.session_state.processed_df)
        if st.session_state.processed_df is not None
        else 0
    )
    st.success(f"**{count} colleges loaded**")
    st.markdown(
        '<div class="cp-brand">👑 <b>Powered by: '
        'Shivasharanappa</b><br>College Intelligence · RAG</div>',
        unsafe_allow_html=True,
    )


# --------------------------- HEADER ---------------------------

st.markdown("""
<div class="cp-hero">
    <h1>🎓 Campus<span style="color:#28bfff">Pulse</span></h1>
    <p style="font-size:1.12rem">Live College Intelligence Platform</p>
    <p>Discover · Analyze · Compare · Choose Smarter</p>
</div>
""", unsafe_allow_html=True)


# ============================================================
# UPLOAD & PROCESS PAGE
# ============================================================

if page == "Upload & Process":

    st.title("📤 Upload & Process")
    st.write(
        "Upload an Excel workbook. CampusPulse scans all worksheets "
        "and detects common college columns where possible."
    )

    uploaded = st.file_uploader(
        "Upload Excel workbook",
        type=["xlsx", "xls"],
        key="college_excel_upload",
    )

    if uploaded is not None:
        raw, sheet, message = read_workbook(uploaded)

        if raw is None:
            st.error(message)

        else:
            st.session_state.raw_df = raw
            st.session_state.source_name = uploaded.name
            st.session_state.sheet_name = sheet

            st.info(message)

            detected = detect_mapping(raw.columns)
            default_name = detected.get("College Name", raw.columns[0])

            st.subheader("1. Confirm the college-name column")

            name_col = st.selectbox(
                "Which column contains the college/institution name?",
                options=list(raw.columns),
                index=list(raw.columns).index(default_name),
                key="manual_college_name",
            )

            manual = {"College Name": name_col}

            with st.expander("Optional: map other columns manually"):
                for logical in [
                    "City", "District", "Taluk", "State", "Website",
                    "Wikipedia", "Address", "Affiliation", "Status",
                    "Type", "Principal Name", "Principal Email",
                    "Rating", "Review Count", "Latitude", "Longitude",
                ]:
                    detected_col = detected.get(logical)
                    options = ["— Not available —"] + list(raw.columns)
                    default_index = (
                        options.index(detected_col)
                        if detected_col in options
                        else 0
                    )

                    selected_col = st.selectbox(
                        f"{logical} column",
                        options=options,
                        index=default_index,
                        key=f"map_{norm(logical)}",
                    )

                    if selected_col != "— Not available —":
                        manual[logical] = selected_col

            st.subheader("2. Preview uploaded data")
            st.dataframe(
                raw.head(10),
                use_container_width=True,
                hide_index=True,
            )

            if st.button(
                "✅ Process Colleges",
                type="primary",
                use_container_width=True,
            ):
                cleaned = clean_colleges(raw, manual)

                if cleaned.empty:
                    st.error(
                        "No college records found. Please select the "
                        "correct college-name column."
                    )
                else:
                    st.session_state.processed_df = prepare_tiers(cleaned)
                    st.session_state.manual_mapping = manual
                    st.session_state.rag_messages = []

                    st.success(
                        f"Successfully processed {len(cleaned)} college "
                        f"records from worksheet '{sheet}'."
                    )

                    st.dataframe(
                        st.session_state.processed_df.head(20),
                        use_container_width=True,
                        hide_index=True,
                    )

                    st.download_button(
                        "⬇️ Download processed CSV",
                        data=csv_bytes(st.session_state.processed_df),
                        file_name="campuspulse_processed_colleges.csv",
                        mime="text/csv",
                        use_container_width=True,
                    )

    elif st.session_state.processed_df is not None:
        st.success(
            f"Current workbook: {st.session_state.source_name}"
        )
        st.info(
            "Your processed data is available. Upload another workbook "
            "above if you want to replace it."
        )
        st.dataframe(
            st.session_state.processed_df.head(10),
            use_container_width=True,
            hide_index=True,
        )

    else:
        st.info("No workbook loaded yet. Upload an Excel file to begin.")


# ============================================================
# DASHBOARD PAGE — ALL FEATURES ON THIS PAGE
# ============================================================

elif page == "Dashboard":

    df = st.session_state.processed_df

    if df is None or df.empty:
        st.title("🏠 CampusPulse Dashboard")
        st.info(
            "Start by opening **Upload & Process** in the sidebar. "
            "Upload your Excel workbook and click **Process Colleges**. "
            "Then return here to see the dashboard and use the RAG chatbot."
        )

        st.markdown("### CampusPulse features")
        c1, c2, c3 = st.columns(3)

        c1.markdown(
            '<div class="cp-panel"><h3>📊 Analytics</h3>'
            '<p>College totals, tier summaries and city distribution.</p></div>',
            unsafe_allow_html=True,
        )
        c2.markdown(
            '<div class="cp-panel"><h3>🔎 Search</h3>'
            '<p>Search college names, locations, websites and affiliations.</p></div>',
            unsafe_allow_html=True,
        )
        c3.markdown(
            '<div class="cp-panel"><h3>🤖 RAG Assistant</h3>'
            '<p>Ask questions grounded in your uploaded workbook.</p></div>',
            unsafe_allow_html=True,
        )

        st.markdown("## 🤖 CampusPulse RAG Assistant")
        st.caption(
            "The chat interface is visible here. Upload a workbook first "
            "to enable college-data answers."
        )

        for message in st.session_state.rag_messages:
            with st.chat_message(message["role"]):
                st.markdown(message["content"])

        question = st.chat_input(
            "Ask about college names, cities, tiers or affiliation...",
            key="rag_input_empty",
        )

        if question:
            st.session_state.rag_messages.extend([
                {"role": "user", "content": question},
                {
                    "role": "assistant",
                    "content": (
                        "Please upload and process a college Excel workbook "
                        "first. Then I can retrieve matching college records."
                    ),
                },
            ])
            st.rerun()

    else:
        df = prepare_tiers(df)
        st.session_state.processed_df = df

        st.title("📊 College Intelligence Dashboard")

        # Metrics
        tier_counts = df["Tier"].value_counts()
        total = len(df)

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Colleges", total)
        m2.metric("Tier 1", int(tier_counts.get("Tier 1", 0)))
        m3.metric("Tier 2", int(tier_counts.get("Tier 2", 0)))
        m4.metric("Tier 3", int(tier_counts.get("Tier 3", 0)))

        # Charts
        left, right = st.columns(2)

        with left:
            st.subheader("🏅 Colleges by Tier")
            tier_chart = pd.DataFrame({
                "Tier": ["Tier 1", "Tier 2", "Tier 3"],
                "Colleges": [
                    int(tier_counts.get("Tier 1", 0)),
                    int(tier_counts.get("Tier 2", 0)),
                    int(tier_counts.get("Tier 3", 0)),
                ],
            }).set_index("Tier")
            st.bar_chart(tier_chart)

        with right:
            st.subheader("📍 Colleges by City")
            city_counts = (
                df["City"].fillna("").astype(str).str.strip()
            )
            city_counts = city_counts[city_counts != ""].value_counts().head(8)

            if not city_counts.empty:
                st.bar_chart(city_counts, horizontal=True)
            else:
                st.info(
                    "City information was not detected. Map the City column "
                    "on the Upload & Process page."
                )

        st.divider()

        # College search
        st.subheader("🔎 College Search & Details")

        search_text = st.text_input(
            "Search by college name, city, district, type, or affiliation",
            key="dashboard_search",
        )

        shown = df.copy()

        if search_text.strip():
            search_mask = shown.fillna("").astype(str).apply(
                lambda column: column.str.contains(
                    re.escape(search_text),
                    case=False,
                    na=False,
                )
            ).any(axis=1)
            shown = shown[search_mask]

        st.caption(f"Showing {len(shown)} matching colleges")

        display_columns = [
            c for c in [
                "College Name", "City", "District", "Taluk", "State",
                "Type", "Status", "Affiliation", "Website", "Wikipedia",
                "Tier", "Reason",
            ] if c in shown.columns
        ]

        st.dataframe(
            shown[display_columns],
            use_container_width=True,
            hide_index=True,
            height=240,
        )

        st.download_button(
            "⬇️ Download filtered colleges",
            data=csv_bytes(shown),
            file_name="campuspulse_filtered_colleges.csv",
            mime="text/csv",
        )

        # Details, map, reviews
        names = df["College Name"].dropna().astype(str).tolist()

        st.subheader("🏫 College Details, Location & Reviews")

        selected = st.selectbox(
            "Choose a college",
            names,
            key="detail_college",
        )

        row = df[df["College Name"].astype(str) == selected].iloc[0]

        details_col, map_col, reviews_col = st.columns([1.3, 1, 1])

        with details_col:
            st.markdown("#### College Details")

            details = []
            for column in [
                "College Name", "City", "District", "Taluk", "State",
                "Address", "Type", "Status", "Affiliation", "Website",
                "Wikipedia", "Tier", "Reason", "Principal Name",
                "Principal Email",
            ]:
                val = value_of(row, column)
                if val:
                    details.append({"Field": column, "Value": val})

            st.dataframe(
                pd.DataFrame(details),
                use_container_width=True,
                hide_index=True,
                height=300,
            )

            web = website_url(row)
            if web:
                st.link_button(
                    "🌐 Official Website",
                    web,
                    use_container_width=True,
                )

            st.link_button(
                "📚 Wikipedia Search",
                wikipedia_url(row),
                use_container_width=True,
            )

        with map_col:
            st.markdown("#### 📍 Location Map")
            st.link_button(
                "🗺️ Open in Google Maps",
                maps_url(row),
                use_container_width=True,
            )

            if value_of(row, "Latitude") and value_of(row, "Longitude"):
                try:
                    latitude = float(value_of(row, "Latitude"))
                    longitude = float(value_of(row, "Longitude"))
                    st.map(pd.DataFrame({
                        "lat": [latitude],
                        "lon": [longitude],
                    }))
                except ValueError:
                    st.info(
                        "The supplied coordinates are invalid. "
                        "Use Google Maps to locate this college."
                    )
            else:
                st.info(
                    "No coordinates were supplied for this college. "
                    "Use Google Maps to view its location."
                )

        with reviews_col:
            st.markdown("#### ⭐ Reviews & Ratings")

            rating = value_of(row, "Rating")
            review_count = value_of(row, "Review Count")

            if rating:
                st.metric("Rating in uploaded data", rating)
            else:
                st.info(
                    "No verified rating was supplied in the workbook."
                )

            if review_count:
                st.write(f"Review count in uploaded data: {review_count}")

            st.caption(
                "Current public reviews are not automatically scraped. "
                "Open a public source to view available reviews."
            )

            st.link_button(
                "⭐ Search Public Reviews",
                reviews_url(row),
                use_container_width=True,
            )

            st.link_button(
                "🗺️ Google Maps Listing",
                maps_url(row),
                use_container_width=True,
            )

        st.divider()

        # Tier results
        st.subheader("🏅 Tier Classification Results")
        st.caption(
            "Preliminary rules based on available fields, not an official "
            "college ranking or accreditation assessment."
        )

        tabs = st.tabs(["Tier 1", "Tier 2", "Tier 3", "All Colleges"])

        for tab, tier in zip(tabs[:3], ["Tier 1", "Tier 2", "Tier 3"]):
            with tab:
                subset = df[df["Tier"] == tier]
                st.write(f"**{len(subset)} colleges**")

                cols = [
                    c for c in [
                        "College Name", "City", "District", "Type",
                        "Status", "Affiliation", "Reason",
                    ] if c in subset.columns
                ]

                st.dataframe(
                    subset[cols],
                    use_container_width=True,
                    hide_index=True,
                    height=250,
                )

                st.download_button(
                    f"Download {tier} CSV",
                    data=csv_bytes(subset),
                    file_name=f"{tier.lower().replace(' ', '_')}_colleges.csv",
                    mime="text/csv",
                    key=f"download_{tier}",
                )

        with tabs[3]:
            st.dataframe(
                df,
                use_container_width=True,
                hide_index=True,
                height=300,
            )

        st.divider()

        # RAG Chatbot
        st.subheader("🤖 CampusPulse RAG Assistant")
        st.write(
            "Ask questions about the colleges in your uploaded workbook. "
            "CampusPulse retrieves relevant records and uses Groq to "
            "generate a grounded answer when configured."
        )

        examples = [
            "Show colleges in Bengaluru",
            "Find Tier 1 colleges",
            "Which colleges have websites?",
            f"Tell me about {selected}",
        ]

        example_columns = st.columns(4)

        for index, example in enumerate(examples):
            if example_columns[index].button(
                example,
                key=f"rag_example_{index}",
                use_container_width=True,
            ):
                st.session_state["rag_pending_question"] = example

        # Render the complete conversation before the input.
        for message in st.session_state.rag_messages:
            with st.chat_message(message["role"]):
                st.markdown(message["content"])

        typed_question = st.chat_input(
            "Ask about college names, cities, tiers, websites or affiliations...",
            key="rag_input",
        )

        pending_question = st.session_state.pop(
            "rag_pending_question", None
        )

        question = typed_question or pending_question

        if question:
            st.session_state.rag_messages.append({
                "role": "user",
                "content": question,
            })

            with st.chat_message("user"):
                st.markdown(question)

            with st.chat_message("assistant"):
                with st.spinner("Searching your college data..."):
                    answer = answer_question(question, df)
                st.markdown(answer)

            st.session_state.rag_messages.append({
                "role": "assistant",
                "content": answer,
            })

        if st.button(
            "🧹 Clear RAG Conversation",
            key="clear_rag_conversation",
        ):
            st.session_state.rag_messages = []
            st.rerun()

        st.caption(
            "Powered by Shivasharanappa · Excel retrieval + optional Groq AI"
        )
