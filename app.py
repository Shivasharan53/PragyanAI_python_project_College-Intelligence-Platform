
import re
import os
from urllib.parse import quote_plus

import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="CampusPulse",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ==========================================================
# CAMPUSPULSE DESIGN
# ==========================================================

st.markdown("""
<style>
.stApp {
    background: #f3f7ff;
}
.block-container {
    max-width: 1550px;
    padding-top: 1rem;
    padding-bottom: 2rem;
}
.hero {
    background: linear-gradient(115deg,#10244a,#1555a0,#0879ed);
    padding: 24px;
    border-radius: 17px;
    color: white;
    margin-bottom: 20px;
}
.hero h1, .hero p {
    color: white !important;
    margin: 0;
}
.hero p {
    margin-top: 7px;
}
.panel {
    background: white;
    border: 1px solid #d8e5f7;
    border-radius: 13px;
    padding: 16px;
    margin-bottom: 12px;
}
[data-testid="stMetric"] {
    background: white;
    border: 1px solid #d8e5f7;
    border-radius: 12px;
    padding: 14px;
}
[data-testid="stSidebar"] {
    background: #122442;
}
[data-testid="stSidebar"] * {
    color: #f3f7ff;
}
.stButton > button,
.stDownloadButton > button,
.stLinkButton > a {
    border-radius: 9px;
    font-weight: 600;
}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
    <h1>🎓 Campus<span style="color:#35c4ff">Pulse</span></h1>
    <p>Live College Intelligence Platform</p>
    <p>Discover · Analyze · Compare · Choose Smarter</p>
    <p style="margin-top:12px;color:#ffe078">
        👑 Powered by: Shivasharanappa
    </p>
</div>
""", unsafe_allow_html=True)


# ==========================================================
# SESSION STATE
# ==========================================================

defaults = {
    "raw_df": None,
    "college_df": None,
    "source_name": "",
    "name_column": None,
    "chat_history": [],
    "last_processed": [],
}

for key, value in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = value


# ==========================================================
# COLUMN DETECTION — SUPPORT DIFFERENT EXCEL FORMATS
# ==========================================================

ALIASES = {
    "College Name": [
        "college name", "clg name", "college", "institution",
        "institution name", "institute", "institute name",
        "name of college", "college/institution"
    ],
    "City": ["city", "town", "place"],
    "District": ["district", "dist"],
    "Taluk": ["taluk", "tq", "tehsil", "taluka"],
    "State": ["state"],
    "Website": [
        "website", "official website", "college website", "url",
        "web site", "homepage"
    ],
    "Address": ["address", "full address"],
    "Affiliation": [
        "affiliation", "affiliated university", "university",
        "affiliated to"
    ],
    "Status": [
        "status", "college status", "autonomous status",
        "recognition status"
    ],
    "Type": ["type", "college type", "institution type", "category"],
    "Wikipedia": ["wikipedia", "wiki", "wikipedia url"],
    "Rating": ["rating", "google rating", "star rating"],
    "Review Count": ["review count", "number of reviews", "reviews count"],
    "Latitude": ["latitude", "lat"],
    "Longitude": ["longitude", "lon", "lng", "long"],
}


def normalize(text):
    return re.sub(r"[^a-z0-9]+", " ", str(text).lower()).strip()


def detect_columns(columns):
    result = {}
    normalized = {normalize(c): c for c in columns}

    for field, aliases in ALIASES.items():
        for alias in aliases:
            if normalize(alias) in normalized:
                result[field] = normalized[normalize(alias)]
                break

        if field not in result:
            for column in columns:
                for alias in aliases:
                    if len(normalize(alias)) >= 4 and (
                        normalize(alias) in normalize(column)
                    ):
                        result[field] = column
                        break
                if field in result:
                    break

    return result


def read_excel_file(uploaded):
    try:
        excel = pd.ExcelFile(uploaded)
        frames = []

        for sheet in excel.sheet_names:
            frame = pd.read_excel(excel, sheet_name=sheet)
            frame = frame.dropna(how="all").dropna(axis=1, how="all")
            if not frame.empty:
                frame.columns = [str(c).strip() for c in frame.columns]
                frames.append((sheet, frame))

        if not frames:
            return None, None

        # Prefer the worksheet with a recognizable college-name column.
        for sheet, frame in frames:
            mapping = detect_columns(frame.columns)
            if "College Name" in mapping:
                return frame, sheet

        # If detection fails, let the user choose the name column.
        return frames[0][1], frames[0][0]

    except Exception as exc:
        st.error(f"Could not read this Excel file: {exc}")
        return None, None


def is_blank(value):
    return pd.isna(value) or str(value).strip().lower() in (
        "", "nan", "none", "nat"
    )


def get_value(row, column):
    if column not in row.index or is_blank(row[column]):
        return ""
    return str(row[column]).strip()


def prepare_data(raw, mapping):
    result = raw.copy()

    for logical_name, source_column in mapping.items():
        if source_column in raw.columns:
            result[logical_name] = raw[source_column]

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

    for column in ALIASES:
        if column not in result.columns:
            result[column] = ""

    return result


# ==========================================================
# CLASSIFICATION
# ==========================================================

def classify_college(row):
    status = get_value(row, "Status").lower()
    typ = get_value(row, "Type").lower()
    affiliation = get_value(row, "Affiliation").lower()

    combined = f"{status} {typ} {affiliation}"

    # Tier 1: indicators requested by the user.
    if any(word in combined for word in [
        "autonomous",
        "university",
        "vtu",
        "visvesvaraya technological university",
    ]):
        return (
            "Tier 1",
            "Autonomous / university / VTU affiliation indicator"
        )

    # Tier 2: some affiliation or recognition information is present.
    if any(word in combined for word in [
        "affiliated", "recognized", "recognised",
        "approved", "accredited"
    ]):
        return (
            "Tier 2",
            "Affiliation or recognition indicator; verify details"
        )

    # Tier 3 means needs review, not necessarily a bad college.
    return (
        "Tier 3",
        "Needs verification / insufficient classification information"
    )


def add_tiers(frame):
    result = frame.copy()
    classifications = result.apply(classify_college, axis=1)

    result["Tier"] = [x[0] for x in classifications]
    result["Reason"] = [x[1] for x in classifications]

    return result


# ==========================================================
# LINKS
# ==========================================================

def official_website(row):
    website = get_value(row, "Website")
    if not website:
        return ""

    if not website.startswith(("https://", "http://")):
        website = "https://" + website

    return website


def google_maps_link(row):
    lat = get_value(row, "Latitude")
    lon = get_value(row, "Longitude")

    if lat and lon:
        try:
            return f"https://www.google.com/maps?q={float(lat)},{float(lon)}"
        except ValueError:
            pass

    query = " ".join(
        x for x in [
            get_value(row, "College Name"),
            get_value(row, "City"),
            get_value(row, "District"),
            get_value(row, "State"),
        ] if x
    )

    return (
        "https://www.google.com/maps/search/?api=1&query="
        + quote_plus(query)
    )


def public_reviews_link(row):
    query = (
        get_value(row, "College Name") + " "
        + get_value(row, "City") + " college reviews"
    )
    return "https://www.google.com/search?q=" + quote_plus(query)


# ==========================================================
# RAG — RETRIEVE RELEVANT EXCEL ROWS, THEN GROQ ANSWER
# ==========================================================

def retrieve_records(question, frame, limit=8):
    if frame is None or frame.empty:
        return pd.DataFrame()

    text_rows = frame.fillna("").astype(str).apply(
        lambda row: " | ".join(
            f"{column}: {row[column]}" for column in frame.columns
        ),
        axis=1,
    ).tolist()

    try:
        from sklearn.feature_extraction.text import TfidfVectorizer
        from sklearn.metrics.pairwise import cosine_similarity

        vectorizer = TfidfVectorizer(
            stop_words="english",
            ngram_range=(1, 2)
        )

        matrix = vectorizer.fit_transform(text_rows + [question])
        scores = cosine_similarity(matrix[-1], matrix[:-1]).ravel()
        indexes = scores.argsort()[::-1]
        indexes = [i for i in indexes if scores[i] > 0][:limit]

        return frame.iloc[indexes].copy()

    except Exception:
        words = [
            word for word in normalize(question).split()
            if len(word) > 2
        ]

        scores = []
        for i, text in enumerate(text_rows):
            score = sum(word in text.lower() for word in words)
            if score:
                scores.append((score, i))

        scores.sort(reverse=True)
        return frame.iloc[
            [index for _, index in scores[:limit]]
        ].copy()


def get_rag_answer(question, frame):
    matches = retrieve_records(question, frame)

    if matches.empty:
        return (
            "I couldn't find matching information in your uploaded "
            "workbook. Try a college name, city, district, website "
            "or affiliation."
        )

    evidence = matches.fillna("").astype(str).to_markdown(index=False)

    try:
        api_key = st.secrets.get(
            "GROQ_API_JNTU_KEY",
            os.getenv("GROQ_API_JNTU_KEY", "")
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
                            "You are CampusPulse, a college-data assistant. "
                            "Answer using only the spreadsheet evidence. "
                            "Never invent ratings, reviews, fees, rankings, "
                            "or facts. Say when information is unavailable."
                        )
                    },
                    {
                        "role": "user",
                        "content": (
                            f"Question: {question}\n\n"
                            f"Retrieved college records:\n{evidence}"
                        )
                    }
                ]
            )

            answer = response.choices[0].message.content
            return answer + "\n\n**Retrieved Excel records**\n\n" + evidence

    except Exception as exc:
        st.warning(
            "Groq AI could not respond. Showing retrieved Excel records "
            "instead. Check your Groq secret if this continues."
        )

    return (
        "### Relevant college records\n\n"
        + evidence
        + "\n\n*AI-generated answers require a valid Groq API secret.*"
    )


# ==========================================================
# SIDEBAR — ONLY TWO PAGES
# ==========================================================

with st.sidebar:
    st.markdown("## 🎓 CampusPulse")

    page = st.radio(
        "Navigation",
        ["Dashboard", "Upload & Process"],
        label_visibility="collapsed"
    )

    count = (
        len(st.session_state.college_df)
        if st.session_state.college_df is not None
        else 0
    )

    st.divider()
    st.info(f"**{count} colleges loaded**")
    st.markdown("👑 **Powered by: Shivasharanappa**")


# ==========================================================
# UPLOAD PAGE
# ==========================================================

if page == "Upload & Process":

    st.title("📤 Upload & Process Colleges")

    st.write(
        "Upload your Excel workbook. Different worksheet names and "
        "common column-name variations are supported."
    )

    uploaded = st.file_uploader(
        "1. Upload Excel File",
        type=["xlsx", "xls"],
        key="college_upload"
    )

    if uploaded is not None:
        raw, sheet = read_excel_file(uploaded)

        if raw is not None:
            st.session_state.raw_df = raw
            st.session_state.source_name = uploaded.name

            detected = detect_columns(raw.columns)
            suggested = detected.get("College Name", raw.columns[0])

            st.success(
                f"Workbook loaded. Worksheet selected: {sheet}"
            )

            st.subheader("2. Select the college-name column")

            name_column = st.selectbox(
                "Which column contains college names?",
                options=list(raw.columns),
                index=list(raw.columns).index(suggested),
                key="college_name_mapping"
            )

            mapping = {"College Name": name_column}

            with st.expander("Optional: map additional columns"):
                for field in [
                    "City", "District", "Taluk", "State",
                    "Website", "Address", "Affiliation", "Status",
                    "Type", "Wikipedia", "Rating", "Review Count",
                    "Latitude", "Longitude"
                ]:
                    options = ["— Not available —"] + list(raw.columns)
                    detected_col = detected.get(field)

                    default_index = (
                        options.index(detected_col)
                        if detected_col in options
                        else 0
                    )

                    selected_col = st.selectbox(
                        field,
                        options,
                        index=default_index,
                        key=f"mapping_{field}"
                    )

                    if selected_col != "— Not available —":
                        mapping[field] = selected_col

            st.subheader("Excel Preview")
            st.dataframe(
                raw.head(8),
                use_container_width=True,
                hide_index=True
            )

            # Selection dropdown.
            preview_data = prepare_data(raw, mapping)

            st.subheader("3. Select Colleges")

            selected_colleges = st.multiselect(
                "Choose one or more colleges to process",
                options=preview_data["College Name"].tolist(),
                default=preview_data["College Name"].tolist(),
                key="selected_colleges"
            )

            # Process button.
            if st.button(
                "⚙️ Process Selected Colleges",
                type="primary",
                use_container_width=True
            ):
                if not selected_colleges:
                    st.warning("Please select at least one college.")
                else:
                    processed = preview_data[
                        preview_data["College Name"].isin(selected_colleges)
                    ].copy()

                    st.session_state.college_df = add_tiers(processed)
                    st.session_state.last_processed = selected_colleges
                    st.session_state.chat_history = []

                    st.success(
                        f"Successfully processed {len(processed)} colleges."
                    )

            # Results appear only after processing.
            if st.session_state.college_df is not None:
                st.divider()
                st.subheader("🏅 Processing Results")

                result = st.session_state.college_df
                tier1 = result[result["Tier"] == "Tier 1"]
                tier2 = result[result["Tier"] == "Tier 2"]
                tier3 = result[result["Tier"] == "Tier 3"]

                c1, c2, c3 = st.columns(3)
                c1.metric("Tier 1 Colleges", len(tier1))
                c2.metric("Tier 2 Colleges", len(tier2))
                c3.metric("Tier 3 Colleges", len(tier3))

                with st.expander("🟢 Tier 1 — Autonomous / University / VTU indicators", expanded=True):
                    st.dataframe(
                        tier1,
                        use_container_width=True,
                        hide_index=True
                    )
                    st.download_button(
                        "Download Tier 1 CSV",
                        tier1.to_csv(index=False).encode("utf-8-sig"),
                        "tier1_colleges.csv",
                        "text/csv"
                    )

                with st.expander("🟡 Tier 2 — Affiliation / Recognition indicators", expanded=True):
                    st.dataframe(
                        tier2,
                        use_container_width=True,
                        hide_index=True
                    )
                    st.download_button(
                        "Download Tier 2 CSV",
                        tier2.to_csv(index=False).encode("utf-8-sig"),
                        "tier2_colleges.csv",
                        "text/csv"
                    )

                with st.expander("🔴 Tier 3 — Needs Verification", expanded=True):
                    st.dataframe(
                        tier3,
                        use_container_width=True,
                        hide_index=True
                    )
                    st.download_button(
                        "Download Tier 3 CSV",
                        tier3.to_csv(index=False).encode("utf-8-sig"),
                        "tier3_colleges.csv",
                        "text/csv"
                    )

                st.download_button(
                    "⬇️ Download All Processed Colleges",
                    result.to_csv(index=False).encode("utf-8-sig"),
                    "campuspulse_processed.csv",
                    "text/csv",
                    use_container_width=True
                )

        else:
            st.error("No readable worksheet was found in this workbook.")

    elif st.session_state.college_df is not None:
        st.success(
            f"{len(st.session_state.college_df)} processed colleges "
            "are available in this session."
        )


# ==========================================================
# DASHBOARD — COLLEGE DROPDOWN + DETAILS + CHATBOX
# ==========================================================

elif page == "Dashboard":

    df = st.session_state.college_df

    if df is None or df.empty:
        st.title("🏠 Dashboard")

        st.info(
            "Start on Upload & Process. Upload your Excel file, select "
            "colleges, and click Process Selected Colleges. Your results "
            "and college chatbox will appear here."
        )

        st.markdown("### CampusPulse workflow")
        st.markdown(
            "1. Upload Excel file\n"
            "2. Select colleges from the dropdown\n"
            "3. Click Process Selected Colleges\n"
            "4. View Tier 1, Tier 2 and Tier 3 results\n"
            "5. Return to Dashboard to inspect a college and ask questions"
        )

    else:
        st.title("🏠 College Dashboard")

        # Main dashboard metrics are data-driven, not tier-focused.
        city_count = (
            df["City"].replace("", pd.NA).nunique()
            if "City" in df.columns else 0
        )
        website_count = (
            df["Website"].apply(lambda x: not is_blank(x)).sum()
            if "Website" in df.columns else 0
        )

        m1, m2, m3 = st.columns(3)
        m1.metric("Total Processed Colleges", len(df))
        m2.metric("Cities Covered", int(city_count))
        m3.metric("Websites Available", int(website_count))

        st.divider()

        st.subheader("📋 Processed College Records")

        search = st.text_input(
            "Search college name, city, district, affiliation or website",
            key="dashboard_search"
        )

        visible = df.copy()

        if search.strip():
            mask = visible.fillna("").astype(str).apply(
                lambda column: column.str.contains(
                    re.escape(search),
                    case=False,
                    na=False
                )
            ).any(axis=1)
            visible = visible[mask]

        st.dataframe(
            visible,
            use_container_width=True,
            hide_index=True,
            height=230
        )

        st.download_button(
            "⬇️ Download Dashboard Records",
            visible.to_csv(index=False).encode("utf-8-sig"),
            "campuspulse_dashboard.csv",
            "text/csv"
        )

        st.divider()

        # College dropdown.
        st.subheader("🏫 College Details")

        college_names = df["College Name"].astype(str).tolist()

        selected_name = st.selectbox(
            "Select a college to view details",
            college_names,
            key="dashboard_college"
        )

        row = df[df["College Name"] == selected_name].iloc[0]

        details = []
        for field in [
            "College Name", "City", "District", "Taluk", "State",
            "Address", "Type", "Status", "Affiliation", "Website",
            "Tier", "Reason"
        ]:
            value = get_value(row, field)
            if value:
                details.append({"Field": field, "Value": value})

        st.dataframe(
            pd.DataFrame(details),
            use_container_width=True,
            hide_index=True
        )

        link1, link2, link3 = st.columns(3)

        website = official_website(row)

        with link1:
            if website:
                st.link_button(
                    "🌐 Official Website",
                    website,
                    use_container_width=True
                )
            else:
                st.caption("Website not supplied")

        with link2:
            st.link_button(
                "📍 Google Maps",
                google_maps_link(row),
                use_container_width=True
            )

        with link3:
            st.link_button(
                "⭐ Public Reviews",
                public_reviews_link(row),
                use_container_width=True
            )

        rating = get_value(row, "Rating")
        review_count = get_value(row, "Review Count")

        if rating:
            st.info(f"Rating provided by workbook: {rating}")
        else:
            st.caption(
                "No rating supplied. Use Public Reviews to check current "
                "public information."
            )

        if review_count:
            st.caption(f"Review count supplied by workbook: {review_count}")

        st.divider()

        # RAG CHATBOX: visible on Dashboard.
        st.subheader("🤖 CampusPulse RAG Chatbot")
        st.write(
            "Ask questions about your processed colleges. Answers are "
            "grounded in relevant records retrieved from your Excel data."
        )

        suggestions = [
            "Show colleges in Bengaluru",
            "Find Tier 1 colleges",
            "Which colleges have websites?",
            f"Tell me about {selected_name}"
        ]

        suggestion_cols = st.columns(4)

        for i, suggestion in enumerate(suggestions):
            if suggestion_cols[i].button(
                suggestion,
                key=f"suggestion_{i}",
                use_container_width=True
            ):
                st.session_state.pending_question = suggestion

        # Display conversation.
        for message in st.session_state.chat_history:
            with st.chat_message(message["role"]):
                st.markdown(message["content"])

        question = st.chat_input(
            "Ask about college names, cities, affiliation or websites...",
            key="campuspulse_chat_input"
        )

        pending = st.session_state.pop("pending_question", None)
        question = question or pending

        if question:
            st.session_state.chat_history.append({
                "role": "user",
                "content": question
            })

            with st.chat_message("user"):
                st.markdown(question)

            with st.chat_message("assistant"):
                with st.spinner("Retrieving college information..."):
                    answer = get_rag_answer(question, df)
                st.markdown(answer)

            st.session_state.chat_history.append({
                "role": "assistant",
                "content": answer
            })

        if st.button("🧹 Clear Chat", key="clear_chat"):
            st.session_state.chat_history = []
            st.rerun()

        st.caption("Powered by Shivasharanappa · CampusPulse College RAG")
