import re
from urllib.parse import quote_plus

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(
    page_title="CampusPulse | College Intelligence",
    page_icon="🎓",
    layout="wide",
)

# -------------------- DESIGN --------------------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
html, body, [class*="css"] {font-family:Inter,sans-serif;}
.stApp {background:linear-gradient(135deg,#f1f6ff,#ffffff);}
.block-container {padding-top:1rem;max-width:1600px;}
.hero {background:linear-gradient(110deg,#111d49,#10488e,#0878d1);
color:white;padding:22px;border-radius:16px;margin-bottom:14px;}
.brand {font-size:2.1rem;font-weight:800;}
.brand span {color:#35c5ff;}
.power {text-align:right;font-weight:700;color:#ffda57;}
.panel {background:white;border:1px solid #d8e5f8;border-radius:13px;
padding:15px;margin-bottom:12px;}
.panel-title {color:#10449b;font-weight:800;font-size:1.08rem;}
.kpi {background:white;border:1px solid #d8e5f8;border-radius:13px;padding:16px;}
.kpi-label {color:#627594;font-size:.82rem;font-weight:700;}
.kpi-value {font-size:1.8rem;color:#1156b8;font-weight:800;}
.stButton>button {border-radius:9px;font-weight:700;min-height:40px;}
[data-testid="stSidebar"] {background:#122447;}
[data-testid="stSidebar"] * {color:white;}
</style>
""", unsafe_allow_html=True)

h1, h2 = st.columns([3, 1.5])
with h1:
    st.markdown("""
    <div class="hero">
      <div class="brand">🎓 Campus<span>Pulse</span></div>
      <div>Live College Intelligence Platform</div>
      <small>Discover · Analyze · Compare · Choose Smarter</small>
    </div>
    """, unsafe_allow_html=True)
with h2:
    st.markdown("""
    <div class="hero power">
      👑 Powered by: Shivasharanappa<br>
      <small style="color:white">College Intelligence · RAG</small>
    </div>
    """, unsafe_allow_html=True)

# -------------------- COLUMN DETECTION --------------------
ALIASES = {
    "name": ["college name", "clg name", "college", "institution",
             "institution name", "institute name", "institute", "name"],
    "city": ["city", "town", "place"],
    "district": ["district", "dist", "district name"],
    "taluk": ["taluk", "tq", "tehsil"],
    "state": ["state", "state name"],
    "website": ["website", "official website", "college website", "site"],
    "wiki": ["wikipedia", "wiki", "wikipedia url"],
    "address": ["address", "full address", "postal address"],
    "pin": ["pin code", "pincode", "postal code", "zip code"],
    "principal": ["principal name", "principal"],
    "email": ["principal email", "email", "official email", "e-mail"],
    "affiliation": ["affiliation", "university", "affiliated university",
                    "affiliated to"],
    "status": ["status", "college status", "autonomous status"],
    "type": ["type", "college type", "institution type", "category"],
    "latitude": ["latitude", "lat"],
    "longitude": ["longitude", "long", "lng"],
    "rating": ["rating", "google rating", "review rating"],
    "reviews": ["review count", "number of reviews", "reviews"],
}

LABELS = {
    "name": "College Name (required)", "city": "City",
    "district": "District", "taluk": "Taluk", "state": "State",
    "website": "Official Website", "wiki": "Wikipedia",
    "address": "Address", "pin": "PIN Code",
    "principal": "Principal Name", "email": "Principal Email",
    "affiliation": "Affiliation / University", "status": "Status",
    "type": "College Type", "latitude": "Latitude",
    "longitude": "Longitude", "rating": "Rating",
    "reviews": "Review Count",
}

def norm(value):
    return re.sub(r"[^a-z0-9]+", " ", str(value).lower()).strip()

def detect_columns(columns):
    result = {}
    for field, aliases in ALIASES.items():
        # Prefer exact normalized heading matches
        for col in columns:
            if norm(col) in [norm(a) for a in aliases]:
                result[field] = col
                break
        # Then allow partial matches
        if field not in result:
            for col in columns:
                if any(norm(a) in norm(col) for a in aliases):
                    result[field] = col
                    break
    return result

def clean_df(df):
    df = df.copy()
    df = df.dropna(how="all").dropna(axis=1, how="all")
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, [c for c in df.columns if c and not c.startswith("Unnamed:")]]
    for c in df.select_dtypes(include="object").columns:
        df[c] = df[c].fillna("").astype(str).str.strip()
        df[c] = df[c].replace({"nan": "", "None": ""})
    return df

def read_workbook(upload, sheet):
    # Reads the chosen worksheet, regardless of its name
    return clean_df(pd.read_excel(upload, sheet_name=sheet))

def classify(row, mapping):
    values = " ".join(
        str(row.get(mapping[k], ""))
        for k in ["affiliation", "status", "type"]
        if k in mapping
    ).lower()

    if any(x in values for x in [
        "autonomous", "deemed", "constituent", "university"
    ]):
        return pd.Series(["Tier 1", "Autonomous / university indicator"])

    if any(x in values for x in [
        "vtu", "affiliated", "recognized", "accredited", "naac", "nba"
    ]):
        return pd.Series(["Tier 2", "Affiliation / recognition indicator"])

    return pd.Series(["Tier 3", "Needs verification / limited information"])

def process_data(raw, mapping):
    name_col = mapping.get("name")
    if not name_col:
        return pd.DataFrame()

    df = raw.copy()
    df[name_col] = df[name_col].fillna("").astype(str).str.strip()
    df = df[(df[name_col] != "") & (df[name_col].str.lower() != "nan")]
    df = df.drop_duplicates(subset=[name_col]).reset_index(drop=True)

    df[["CampusPulse Tier", "Classification Reason"]] = df.apply(
        lambda row: classify(row, mapping), axis=1
    )
    return df

def value(row, mapping, field):
    col = mapping.get(field)
    if not col or col not in row.index:
        return ""
    v = row[col]
    return "" if pd.isna(v) else str(v).strip()

def google_maps(name, city=""):
    return "https://www.google.com/maps/search/" + quote_plus(name + " " + city)

def review_search(name, city=""):
    return "https://www.google.com/search?q=" + quote_plus(
        name + " " + city + " college reviews"
    )

def display_table(df, mapping):
    rename = {
        "name": "College Name", "city": "City", "district": "District",
        "taluk": "Taluk", "state": "State", "website": "Website",
        "wiki": "Wikipedia", "affiliation": "Affiliation",
        "status": "Status", "type": "Type",
    }
    out = pd.DataFrame(index=df.index)
    for field, label in rename.items():
        col = mapping.get(field)
        if col and col in df.columns:
            out[label] = df[col]
    out["Tier"] = df["CampusPulse Tier"]
    out["Reason"] = df["Classification Reason"]
    st.dataframe(out, use_container_width=True, hide_index=True)

# -------------------- SESSION STATE --------------------
if "data" not in st.session_state:
    st.session_state.data = None
if "mapping" not in st.session_state:
    st.session_state.mapping = {}
if "source" not in st.session_state:
    st.session_state.source = ""

with st.sidebar:
    st.markdown("## 🎓 CampusPulse")
    page = st.radio(
        "NAVIGATION",
        ["Dashboard", "Upload & Process", "Tier Results",
         "College Search", "College Map", "Reviews", "RAG Chatbot", "About"]
    )
    st.divider()
    if st.session_state.data is not None:
        st.success(f"{len(st.session_state.data)} colleges loaded")
    else:
        st.info("Upload an Excel workbook to begin.")

def ready():
    if st.session_state.data is None or st.session_state.data.empty:
        st.warning("Please upload and process a workbook in Upload & Process.")
        return False
    return True

# -------------------- UPLOAD PAGE --------------------
if page == "Upload & Process":
    st.markdown("## 📥 1. Upload Excel File")
    st.write(
        "CampusPulse checks every worksheet. It supports different column "
        "names and includes manual mapping when automatic detection fails."
    )
    uploaded = st.file_uploader(
        "Upload an Excel workbook", type=["xlsx", "xls"]
    )

    if uploaded:
        try:
            workbook = pd.ExcelFile(uploaded)
            sheet = st.selectbox("Choose worksheet", workbook.sheet_names)
            raw = read_workbook(uploaded, sheet)
            detected = detect_columns(raw.columns)

            st.success(
                f"Worksheet: {sheet} · {len(raw)} rows · "
                f"{len(raw.columns)} columns"
            )
            st.markdown("### Map your columns")
            st.caption(
                "Automatic matches are selected. Only College Name is required."
            )
            options = ["— Not available —"] + list(raw.columns)
            mapping = {}
            cols = st.columns(3)
            for i, field in enumerate(LABELS):
                with cols[i % 3]:
                    default = detected.get(field)
                    index = options.index(default) if default in options else 0
                    selected = st.selectbox(
                        LABELS[field], options, index=index,
                        key=f"mapping_{field}"
                    )
                    if selected != "— Not available —":
                        mapping[field] = selected

            st.markdown("### Preview")
            st.dataframe(raw.head(8), use_container_width=True)

            if not mapping.get("name"):
                st.warning(
                    "Choose the column that contains college names."
                )
            elif st.button(
                "✅ Process Colleges", type="primary",
                use_container_width=True
            ):
                result = process_data(raw, mapping)
                if result.empty:
                    st.error("No college names were found in that column.")
                else:
                    st.session_state.data = result
                    st.session_state.mapping = mapping
                    st.session_state.source = uploaded.name
                    st.success(f"Processed {len(result)} unique colleges.")
                    st.rerun()
        except Exception as exc:
            st.error(f"Could not read this Excel workbook: {exc}")
            st.info("Try saving your workbook as .xlsx and uploading again.")

# -------------------- DASHBOARD PAGE --------------------
elif page == "Dashboard":
    if not ready():
        st.markdown("""
        <div class="panel">
          <div class="panel-title">Welcome to CampusPulse</div>
          Upload your college Excel file to extract college records, classify
          institutions, compare details and ask questions through RAG.
        </div>
        """, unsafe_allow_html=True)
        a, b, c = st.columns(3)
        a.info("📥 Excel extractor\n\nAutomatic detection + manual mapping.")
        b.info("🏅 Tier classification\n\nTransparent rules and explanations.")
        c.info("🤖 RAG assistant\n\nRetrieve evidence from your workbook.")
    else:
        df = st.session_state.data
        mapping = st.session_state.mapping
        k1, k2, k3, k4 = st.columns(4)
        metrics = [
            ("TOTAL COLLEGES", len(df)),
            ("TIER 1", int((df["CampusPulse Tier"] == "Tier 1").sum())),
            ("TIER 2", int((df["CampusPulse Tier"] == "Tier 2").sum())),
            ("TIER 3", int((df["CampusPulse Tier"] == "Tier 3").sum())),
        ]
        for col, (label, number) in zip([k1, k2, k3, k4], metrics):
            with col:
                st.markdown(
                    f'<div class="kpi"><div class="kpi-label">{label}</div>'
                    f'<div class="kpi-value">{number:,}</div></div>',
                    unsafe_allow_html=True
                )

        left, right = st.columns(2)
        with left:
            st.markdown("### 🏅 Tier Classification Results")
            counts = df["CampusPulse Tier"].value_counts().reindex(
                ["Tier 1", "Tier 2", "Tier 3"], fill_value=0
            ).rename_axis("Tier").reset_index(name="Colleges")
            fig = px.bar(
                counts, x="Tier", y="Colleges", text="Colleges",
                color="Tier",
                color_discrete_map={
                    "Tier 1": "#16b86b", "Tier 2": "#f3ad16",
                    "Tier 3": "#ef4444"
                }
            )
            fig.update_layout(height=300, showlegend=False)
            st.plotly_chart(fig, use_container_width=True)

        with right:
            st.markdown("### 📍 Colleges by City")
            city_col = mapping.get("city")
            if city_col and city_col in df.columns:
                counts = df[city_col].replace("", pd.NA).dropna().value_counts()
                counts = counts.head(10).rename_axis("City").reset_index(name="Colleges")
                fig = px.bar(counts, x="Colleges", y="City", orientation="h")
                fig.update_layout(height=300, yaxis={"categoryorder": "total ascending"})
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("City data was not detected in the uploaded workbook.")

        st.markdown("### 📋 Extracted Colleges (Preview)")
        search = st.text_input("Search college name, city or keyword")
        view = df
        if search:
            mask = df.astype(str).apply(
                lambda col: col.str.contains(search, case=False, na=False)
            ).any(axis=1)
            view = df[mask]
        display_table(view.head(200), mapping)
        st.download_button(
            "⬇️ Download processed CSV",
            df.to_csv(index=False).encode("utf-8-sig"),
            "campuspulse_processed.csv", "text/csv",
            type="primary"
        )

# -------------------- TIER RESULTS --------------------
elif page == "Tier Results":
    st.markdown("## 🏅 Tier Classification Results")
    if ready():
        df = st.session_state.data
        mapping = st.session_state.mapping
        st.warning(
            "Tiers are keyword-based indicators, not official rankings or "
            "accreditation results. Verify college details independently."
        )
        tabs = st.tabs(["Tier 1", "Tier 2", "Tier 3", "All Colleges"])
        for tab, tier in zip(tabs, ["Tier 1", "Tier 2", "Tier 3", "ALL"]):
            with tab:
                sub = df if tier == "ALL" else df[df["CampusPulse Tier"] == tier]
                st.metric("Colleges", len(sub))
                display_table(sub, mapping)
                st.download_button(
                    f"Download {tier} CSV",
                    sub.to_csv(index=False).encode("utf-8-sig"),
                    f"{tier.lower().replace(' ', '_')}.csv",
                    "text/csv", key=f"download_{tier}"
                )

# -------------------- COLLEGE SEARCH --------------------
elif page == "College Search":
    st.markdown("## 🔎 College Search & Compare")
    if ready():
        df = st.session_state.data
        mapping = st.session_state.mapping
        name_col = mapping["name"]
        names = df[name_col].astype(str).tolist()
        selected = st.selectbox("Select a college", names)
        row = df[df[name_col].astype(str) == selected].iloc[0]
        left, right = st.columns([1.3, 1])

        with left:
            st.markdown("### College Profile")
            st.subheader(selected)
            for field, label in [
                ("city", "City"), ("district", "District"),
                ("taluk", "Taluk"), ("state", "State"),
                ("address", "Address"), ("pin", "PIN Code"),
                ("principal", "Principal"), ("email", "Principal Email"),
                ("affiliation", "Affiliation"), ("status", "Status"),
                ("type", "Type"),
            ]:
                v = value(row, mapping, field)
                if v:
                    st.write(f"**{label}:** {v}")

            for field, label in [("website", "Official Website"), ("wiki", "Wikipedia")]:
                v = value(row, mapping, field)
                if v:
                    url = v if v.startswith(("http://", "https://")) else "https://" + v
                    st.markdown(f"**{label}:** [{v}]({url})")

            st.write(
                f"**CampusPulse tier:** {row['CampusPulse Tier']} — "
                f"{row['Classification Reason']}"
            )

        with right:
            city = value(row, mapping, "city")
            st.markdown("### Quick Links")
            st.link_button(
                "🌐 Official Website",
                value(row, mapping, "website")
                if value(row, mapping, "website").startswith(("http://", "https://"))
                else "https://www.google.com/search?q=" + quote_plus(selected + " official website"),
                use_container_width=True
            )
            st.link_button("📍 Google Maps", google_maps(selected, city),
                           use_container_width=True)
            st.link_button("⭐ Public Reviews", review_search(selected, city),
                           use_container_width=True)
            st.link_button(
                "📚 Wikipedia Search",
                "https://en.wikipedia.org/w/index.php?search=" + quote_plus(selected),
                use_container_width=True
            )

        st.markdown("### Compare Colleges")
        compare = st.multiselect("Choose up to four colleges", names,
                                 default=[selected], max_selections=4)
        if compare:
            display_table(df[df[name_col].astype(str).isin(compare)], mapping)

# -------------------- COLLEGE MAP --------------------
elif page == "College Map":
    st.markdown("## 📍 College Map")
    if ready():
        df = st.session_state.data
        mapping = st.session_state.mapping
        lat_col, lon_col = mapping.get("latitude"), mapping.get("longitude")

        if lat_col and lon_col:
            import pydeck as pdk
            coords = df.copy()
            coords["_lat"] = pd.to_numeric(coords[lat_col], errors="coerce")
            coords["_lon"] = pd.to_numeric(coords[lon_col], errors="coerce")
            coords = coords[
                coords["_lat"].between(-90, 90) &
                coords["_lon"].between(-180, 180)
            ]
            if not coords.empty:
                st.pydeck_chart(pdk.Deck(
                    initial_view_state=pdk.ViewState(
                        latitude=float(coords["_lat"].mean()),
                        longitude=float(coords["_lon"].mean()),
                        zoom=5
                    ),
                    layers=[pdk.Layer(
                        "ScatterplotLayer", data=coords,
                        get_position="[_lon, _lat]", get_radius=2500,
                        get_fill_color=[20, 100, 240, 180], pickable=True
                    )]
                ), use_container_width=True)
            else:
                st.info("No valid latitude/longitude values were found.")
        else:
            st.info(
                "No coordinates were supplied. CampusPulse will not invent "
                "map locations; use the Google Maps search links below."
            )

        for _, row in df.head(100).iterrows():
            name = value(row, mapping, "name")
            city = value(row, mapping, "city")
            st.markdown(f"- [{name} on Google Maps]({google_maps(name, city)})")

# -------------------- REVIEWS --------------------
elif page == "Reviews":
    st.markdown("## ⭐ Reviews & Ratings")
    if ready():
        df = st.session_state.data
        mapping = st.session_state.mapping
        st.info(
            "Ratings and reviews are not fabricated. Supplied rating fields "
            "are shown when available; otherwise use public review links."
        )
        rating_col, reviews_col = mapping.get("rating"), mapping.get("reviews")
        if rating_col or reviews_col:
            cols = [mapping["name"]]
            cols += [c for c in [rating_col, reviews_col] if c and c not in cols]
            st.dataframe(df[cols], use_container_width=True, hide_index=True)

        for _, row in df.head(100).iterrows():
            name = value(row, mapping, "name")
            city = value(row, mapping, "city")
            with st.expander(f"⭐ {name}"):
                if rating_col:
                    st.write("Workbook rating:", row[rating_col])
                if reviews_col:
                    st.write("Workbook review count:", row[reviews_col])
                st.link_button("Search Public Reviews", review_search(name, city))
                st.link_button("Google Maps", google_maps(name, city))

# -------------------- RAG CHATBOT --------------------
elif page == "RAG Chatbot":
    st.markdown("## 🤖 CampusPulse RAG Assistant")
    st.write("Ask questions about the college records in your uploaded workbook.")

    if ready():
        df = st.session_state.data
        question = st.chat_input("Ask about college names, cities, tiers or affiliation")

        if "chat_history" not in st.session_state:
            st.session_state.chat_history = []
        if "rag_evidence" not in st.session_state:
            st.session_state.rag_evidence = []

        if question:
            documents = df.astype(str).agg(" | ".join, axis=1).tolist()
            results = []
            try:
                from sklearn.feature_extraction.text import TfidfVectorizer
                from sklearn.metrics.pairwise import cosine_similarity
                vectorizer = TfidfVectorizer(
                    ngram_range=(1, 2), stop_words="english"
                )
                matrix = vectorizer.fit_transform(documents + [question])
                scores = cosine_similarity(matrix[-1], matrix[:-1]).ravel()
                indices = scores.argsort()[::-1][:5]
                results = [
                    documents[i] for i in indices if scores[i] > 0
                ]
            except Exception:
                q = set(norm(question).split())
                results = sorted(
                    documents,
                    key=lambda doc: len(q & set(norm(doc).split())),
                    reverse=True
                )[:5]

            answer = ""
            context = "\n\n".join(results)

            # Add GROQ_API_KEY under Streamlit Secrets for generated answers.
            try:
                api_key = st.secrets.get("GROQ_API_KEY", "")
            except Exception:
                api_key = ""

            if api_key and results:
                try:
                    from groq import Groq
                    client = Groq(api_key=api_key)
                    response = client.chat.completions.create(
                        model="llama-3.3-70b-versatile",
                        messages=[
                            {
                                "role": "system",
                                "content": (
                                    "You are CampusPulse. Answer only from "
                                    "the spreadsheet context. If information "
                                    "is missing, say so. Never invent ratings, "
                                    "fees, accreditation or reviews."
                                )
                            },
                            {
                                "role": "user",
                                "content": f"Question: {question}\n\nContext:\n{context}"
                            }
                        ],
                        temperature=0.2
                    )
                    answer = response.choices[0].message.content
                except Exception:
                    answer = "The AI service was unavailable. Here are the matching spreadsheet records."

            if not answer:
                answer = (
                    "These are the closest matching spreadsheet records. "
                    "For generated AI answers, configure GROQ_API_KEY in "
                    "Streamlit Secrets.\n\n" + (context or "No matching records found.")
                )

            st.session_state.chat_history.append(("user", question))
            st.session_state.chat_history.append(("assistant", answer))
            st.session_state.rag_evidence = results

        for role, message in st.session_state.chat_history[-12:]:
            with st.chat_message(role):
                st.write(message)

        if st.session_state.rag_evidence:
            with st.expander("📚 Retrieved Excel evidence"):
                for item in st.session_state.rag_evidence:
                    st.write(item)

# -------------------- ABOUT --------------------
else:
    st.markdown("## ℹ️ About CampusPulse")
    st.markdown("""
    CampusPulse is a college-information dashboard built around uploaded
    Excel workbooks.

    - Flexible Excel upload and manual column mapping
    - Search and compare college records
    - Transparent heuristic tier classification
    - College map links and public review searches
    - RAG retrieval with optional Groq-generated answers

    **Important:** Tier labels are not official accreditation rankings.
    Only use real, verified ratings and location coordinates.
    """)
