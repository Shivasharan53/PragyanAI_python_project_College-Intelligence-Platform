import os
import re
import pandas as pd
import gradio as gr


# ============================================================
# CAMPUSPULSE — GRADIO + RAG COLLEGE ASSISTANT
# ============================================================

APP_TITLE = "CampusPulse RAG Assistant"
GROQ_MODEL = os.getenv(
    "GROQ_MODEL",
    "llama-3.3-70b-versatile"
)

# Flexible Excel column names
ALIASES = {
    "college_name": [
        "college name", "clg name", "college", "institution",
        "institution name", "institute", "institute name",
        "engineering college", "name"
    ],
    "city": ["city", "town", "place"],
    "district": ["district", "dist", "district name"],
    "taluk": ["taluk", "tq", "tehsil", "sub district"],
    "state": ["state", "state name"],
    "website": [
        "website", "official website", "college website",
        "web site", "homepage", "site"
    ],
    "wikipedia": ["wikipedia", "wiki", "wikipedia url"],
    "address": ["address", "full address", "postal address"],
    "affiliation": [
        "affiliation", "university", "affiliated university",
        "affiliated to"
    ],
    "status": ["status", "college status", "autonomous status"],
    "type": ["type", "college type", "institution type", "category"],
    "principal": ["principal name", "principal"],
    "email": ["principal email", "official email", "email", "e-mail"],
    "rating": ["rating", "google rating", "review rating"],
    "review_count": ["review count", "number of reviews", "reviews"],
}


# ============================================================
# NORMALIZE COLUMN NAMES
# ============================================================

def normalize(value):
    """Normalize a column name or search query."""
    return re.sub(
        r"[^a-z0-9]+",
        " ",
        str(value).lower()
    ).strip()


def detect_columns(columns):
    """Detect common column headings automatically."""
    mapping = {}

    normalized_columns = {
        normalize(column): column
        for column in columns
    }

    for field, aliases in ALIASES.items():

        # First, look for exact matches.
        for alias in aliases:
            key = normalize(alias)

            if key in normalized_columns:
                mapping[field] = normalized_columns[key]
                break

        # Then, look for partial matches.
        if field not in mapping:
            for column in columns:
                column_name = normalize(column)

                for alias in aliases:
                    alias_name = normalize(alias)

                    if (
                        alias_name
                        and alias_name in column_name
                    ):
                        mapping[field] = column
                        break

                if field in mapping:
                    break

    return mapping


# ============================================================
# LOAD EXCEL WORKBOOK
# ============================================================

def load_excel(file_path):
    """
    Read an uploaded Excel workbook.
    Scan all worksheets and select the largest worksheet
    containing a recognizable college-name column.
    """

    if not file_path:
        return None, "Please upload an Excel workbook first."

    try:
        workbook = pd.ExcelFile(file_path)
        candidates = []

        for sheet_name in workbook.sheet_names:
            try:
                df = pd.read_excel(
                    workbook,
                    sheet_name=sheet_name
                )

                df = df.dropna(
                    how="all"
                ).dropna(
                    axis=1,
                    how="all"
                )

                df.columns = [
                    str(column).strip()
                    for column in df.columns
                ]

                df = df.loc[
                    :,
                    [
                        column
                        for column in df.columns
                        if column
                        and not column.startswith("Unnamed:")
                    ]
                ]

                if df.empty or len(df.columns) == 0:
                    continue

                mapping = detect_columns(df.columns)

                if "college_name" in mapping:
                    candidates.append(
                        (
                            len(df),
                            sheet_name,
                            df,
                            mapping
                        )
                    )

            except Exception:
                # Ignore unreadable sheets and inspect the others.
                continue

        if not candidates:
            return None, (
                "No worksheet with a recognizable college-name "
                "column was found.\n\n"
                "Rename the college column to College Name, "
                "Clg name, Institution, College or Institute."
            )

        # Prefer the largest worksheet with a college-name column.
        _, sheet_name, df, mapping = max(
            candidates,
            key=lambda item: item[0]
        )

        name_column = mapping["college_name"]

        df = df.copy()
        df[name_column] = (
            df[name_column]
            .fillna("")
            .astype(str)
            .str.strip()
        )

        df = df[
            (df[name_column] != "")
            & (df[name_column].str.lower() != "nan")
        ]

        df = df.drop_duplicates(
            subset=[name_column]
        ).reset_index(drop=True)

        if df.empty:
            return None, (
                "The selected worksheet contains no college names."
            )

        # Create searchable documents from each spreadsheet row.
        documents = []

        for _, row in df.iterrows():
            fields = []

            for column in df.columns:
                value = row[column]

                if pd.isna(value):
                    continue

                value = str(value).strip()

                if value and value.lower() not in [
                    "nan", "none", "nat"
                ]:
                    fields.append(f"{column}: {value}")

            documents.append(" | ".join(fields))

        data = {
            "df": df,
            "mapping": mapping,
            "documents": documents,
            "sheet_name": sheet_name,
            "name_column": name_column,
        }

        message = (
            f"Successfully loaded {len(df)} college records "
            f"from worksheet '{sheet_name}'."
        )

        return data, message

    except Exception as error:
        return None, (
            f"Unable to read the Excel workbook: {error}\n\n"
            "Try uploading a valid .xlsx or .xls file."
        )


# ============================================================
# RAG RETRIEVER
# ============================================================

def retrieve_documents(question, documents, top_k=5):
    """
    Retrieval stage:
    1. Convert college records and question into TF-IDF vectors.
    2. Calculate cosine similarity.
    3. Return the most relevant spreadsheet rows.
    """

    if not documents:
        return []

    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.metrics.pairwise import cosine_similarity

    vectorizer = TfidfVectorizer(
        ngram_range=(1, 2),
        stop_words="english"
    )

    matrix = vectorizer.fit_transform(
        documents + [question]
    )

    scores = cosine_similarity(
        matrix[-1],
        matrix[:-1]
    ).flatten()

    ranked_indices = scores.argsort()[::-1][:top_k]

    results = []

    for index in ranked_indices:
        if scores[index] > 0:
            results.append({
                "text": documents[index],
                "score": float(scores[index]),
            })

    return results


# ============================================================
# GENERATE GROUNDED ANSWER
# ============================================================

def generate_answer(question, evidence):
    """
    Use Groq to generate a natural-language answer.
    Without a key, return the retrieved evidence instead.
    """

    if not evidence:
        return (
            "I couldn't find matching information in your workbook. "
            "Try searching for a college name, city, district, "
            "website or affiliation."
        )

    context = "\n\n".join(
        item["text"]
        for item in evidence
    )

    api_key = os.getenv("GROQ_API_KEY", "").strip()

    if not api_key:
        return (
            "### Retrieved College Records\n\n"
            "No GROQ_API_KEY is configured, so the chatbot is "
            "currently running in retrieval-only mode.\n\n"
            f"{context}\n\n"
            "Configure GROQ_API_KEY in your hosting environment "
            "to enable generated AI answers."
        )

    try:
        from groq import Groq

        client = Groq(api_key=api_key)

        response = client.chat.completions.create(
            model=GROQ_MODEL,
            messages=[
                {
                    "role": "system",
                    "content": (
                        "You are CampusPulse, a college information "
                        "assistant. Answer the user's question using "
                        "only the supplied Excel evidence. "
                        "If the evidence does not contain the answer, "
                        "clearly say that the uploaded data does not "
                        "provide that information. Never invent "
                        "college ratings, reviews, fees, accreditation, "
                        "addresses or official rankings. "
                        "Distinguish data from inference."
                    ),
                },
                {
                    "role": "user",
                    "content": (
                        f"QUESTION:\n{question}\n\n"
                        f"RETRIEVED EXCEL EVIDENCE:\n{context}"
                    ),
                },
            ],
            temperature=0.2,
        )

        return response.choices[0].message.content

    except Exception as error:
        return (
            "The AI generation service could not complete this "
            "request. Here are the retrieved records instead.\n\n"
            f"{context}\n\n"
            f"Technical detail: {error}"
        )


# ============================================================
# CHAT HANDLER
# ============================================================

def respond(question, file_path, history):
    """Process one user question and update Gradio chat history."""

    history = history or []

    if not question or not question.strip():
        return history

    data, status = load_excel(file_path)

    if data is None:
        return history + [
            {
                "role": "user",
                "content": question,
            },
            {
                "role": "assistant",
                "content": status,
            },
        ]

    documents = data["documents"]

    try:
        evidence = retrieve_documents(
            question,
            documents,
            top_k=5
        )

        answer = generate_answer(
            question,
            evidence
        )

        # Show supporting rows, so users can inspect the evidence.
        if evidence:
            evidence_text = "\n\n".join(
                f"**Record {index + 1}** "
                f"(similarity: {item['score']:.3f})\n"
                f"{item['text']}"
                for index, item in enumerate(evidence)
            )

            answer += (
                "\n\n---\n"
                "### 📚 Retrieved Evidence\n\n"
                f"{evidence_text}"
            )

    except Exception as error:
        answer = (
            "An error occurred while retrieving college information.\n\n"
            f"Details: {error}"
        )

    return history + [
        {
            "role": "user",
            "content": question,
        },
        {
            "role": "assistant",
            "content": answer,
        },
    ]


# ============================================================
# GRADIO INTERFACE
# ============================================================

CUSTOM_CSS = """
.gradio-container {
    max-width: 1150px !important;
    margin: auto !important;
}
#main-title {
    text-align: center;
    padding: 16px;
    border-radius: 16px;
    background: linear-gradient(110deg, #101e4c, #0878d1);
    color: white;
    margin-bottom: 14px;
}
#main-title h1 {
    color: white;
}
#main-title span {
    color: #43ccff;
}
"""

with gr.Blocks(
    theme=gr.themes.Soft(
        primary_hue="blue",
        secondary_hue="indigo",
    ),
    css=CUSTOM_CSS,
    title=APP_TITLE,
) as demo:

    gr.HTML("""
    <div id="main-title">
        <h1>🎓 Campus<span>Pulse</span></h1>
        <p>RAG College Intelligence Assistant</p>
        <small>
            Upload Excel · Retrieve Evidence · Ask Questions
        </small>
    </div>
    """)

    gr.Markdown(
        """
        Upload a college spreadsheet and ask questions about its
        college names, locations, affiliations and other available fields.
        """
    )

    with gr.Row():
        with gr.Column(scale=1):
            excel_file = gr.File(
                label="📥 Upload College Excel Workbook",
                file_types=[".xlsx", ".xls"],
                type="filepath",
            )

            load_button = gr.Button(
                "🔍 Check Excel File",
                variant="secondary",
            )

            file_status = gr.Markdown(
                "Upload a workbook to get started."
            )

            gr.Markdown(
                """
                **Example questions**
                - Find colleges in Bengaluru.
                - Compare two colleges.
                - Find colleges affiliated with VTU.
                - Show colleges with website information.
                - Which college records are missing an address?
                """
            )

        with gr.Column(scale=2):
            chatbot = gr.Chatbot(
                label="💬 Ask CampusPulse",
                type="messages",
                height=480,
            )

            question = gr.Textbox(
                label="Your question",
                placeholder=(
                    "Ask about college names, cities, districts, "
                    "websites or affiliations..."
                ),
                lines=2,
            )

            with gr.Row():
                ask_button = gr.Button(
                    "🚀 Ask CampusPulse",
                    variant="primary",
                )

                clear_button = gr.Button(
                    "🧹 Clear Chat",
                    variant="secondary",
                )

    def check_file(file_path):
        data, status = load_excel(file_path)

        if data is None:
            return status

        return (
            f"✅ {status}\n\n"
            f"Detected columns: "
            f"{', '.join(data['mapping'].keys())}"
        )

    load_button.click(
        fn=check_file,
        inputs=[excel_file],
        outputs=[file_status],
    )

    ask_button.click(
        fn=respond,
        inputs=[question, excel_file, chatbot],
        outputs=[chatbot],
    ).then(
        fn=lambda: "",
        inputs=None,
        outputs=[question],
    )

    question.submit(
        fn=respond,
        inputs=[question, excel_file, chatbot],
        outputs=[chatbot],
    ).then(
        fn=lambda: "",
        inputs=None,
        outputs=[question],
    )

    clear_button.click(
        fn=lambda: [],
        inputs=None,
        outputs=[chatbot],
    )

    gr.Markdown(
        """
        ---
        **CampusPulse | Powered by Shivasharanappa**

        **RAG pipeline:** Excel workbook → text documents →
        TF-IDF retrieval → relevant evidence → optional Groq LLM.

        The chatbot only has access to the workbook uploaded in this
        Gradio interface. It does not automatically share the current
        Streamlit session's uploaded data.
        """
    )


# ============================================================
# LAUNCH
# ============================================================

if __name__ == "__main__":
    demo.launch(
        server_name="0.0.0.0",
        server_port=int(os.getenv("PORT", "7860")),
    )
