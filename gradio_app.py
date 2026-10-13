import os
import re

import gradio as gr
import pandas as pd


def normalize(text):
    return re.sub(r"[^a-z0-9]+", " ", str(text).lower()).strip()


ALIASES = {
    "college": ["college name", "clg name", "college", "institution",
                "institution name", "institute", "institute name", "name"],
    "city": ["city", "town", "place"],
    "district": ["district", "dist"],
    "affiliation": ["affiliation", "university", "affiliated university"],
    "website": ["website", "official website", "college website"],
}


def load_excel(file_path):
    if not file_path:
        return None, "Please upload an Excel workbook."

    try:
        workbook = pd.ExcelFile(file_path)
        candidates = []

        for sheet in workbook.sheet_names:
            df = pd.read_excel(workbook, sheet_name=sheet)
            df = df.dropna(how="all").dropna(axis=1, how="all")

            mapping = {}
            for field, aliases in ALIASES.items():
                for column in df.columns:
                    if normalize(column) in [normalize(a) for a in aliases]:
                        mapping[field] = column
                        break

            if "college" in mapping:
                candidates.append((len(df), df, mapping, sheet))

        if not candidates:
            return None, (
                "Could not identify the college-name column. Rename a column "
                "to College Name, Clg name, College, or Institution."
            )

        _, df, mapping, sheet = max(candidates, key=lambda item: item[0])
        df = df.dropna(subset=[mapping["college"]]).copy()
        df[mapping["college"]] = df[mapping["college"]].astype(str).str.strip()
        df = df[df[mapping["college"]] != ""]
        df = df.drop_duplicates(subset=[mapping["college"]])

        documents = df.astype(str).agg(" | ".join, axis=1).tolist()

        return (df, mapping, documents), (
            f"Loaded {len(df)} college records from worksheet '{sheet}'."
        )

    except Exception as exc:
        return None, f"Could not read workbook: {exc}"


def retrieve(question, documents, top_k=5):
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.metrics.pairwise import cosine_similarity

    if not documents:
        return []

    vectorizer = TfidfVectorizer(
        ngram_range=(1, 2),
        stop_words="english"
    )
    matrix = vectorizer.fit_transform(documents + [question])
    scores = cosine_similarity(matrix[-1], matrix[:-1]).ravel()
    indices = scores.argsort()[::-1][:top_k]

    return [documents[i] for i in indices if scores[i] > 0]


def respond(question, file_path, history):
    if not question or not question.strip():
        return "Please enter a question.", history or []

    data, status = load_excel(file_path)
    if data is None:
        return status, (history or [])

    df, mapping, documents = data

    try:
        evidence = retrieve(question, documents)
    except Exception as exc:
        return f"Retrieval error: {exc}", (history or [])

    if not evidence:
        answer = (
            "I could not find a matching record in this workbook. "
            "Try a college name, city, district, website, or affiliation."
        )
        return answer, (history or []) + [
            {"role": "user", "content": question},
            {"role": "assistant", "content": answer},
        ]

    context = "\n\n".join(evidence)
    api_key = os.getenv("GROQ_API_KEY", "")
    answer = ""

    if api_key:
        try:
            from groq import Groq

            client = Groq(api_key=api_key)
            result = client.chat.completions.create(
                model=os.getenv(
                    "GROQ_MODEL", "llama-3.3-70b-versatile"
                ),
                messages=[
                    {
                        "role": "system",
                        "content": (
                            "You are CampusPulse, a college information "
                            "assistant. Answer only from the spreadsheet "
                            "context below. If the answer is missing, say "
                            "the uploaded data does not contain it. Never "
                            "invent ratings, fees, reviews, or accreditation."
                        ),
                    },
                    {
                        "role": "user",
                        "content": (
                            f"Question: {question}\n\n"
                            f"Retrieved spreadsheet context:\n{context}"
                        ),
                    },
                ],
                temperature=0.2,
            )
            answer = result.choices[0].message.content

        except Exception as exc:
            answer = (
                f"AI generation is unavailable ({exc}). "
                f"Retrieved records:\n\n{context}"
            )
    else:
        answer = (
            "Retrieved matching records (retrieval-only mode). Configure "
            "GROQ_API_KEY to enable generated answers.\n\n" + context
        )

    new_history = (history or []) + [
        {"role": "user", "content": question},
        {"role": "assistant", "content": answer},
    ]
    return answer, new_history


with gr.Blocks(
    theme=gr.themes.Soft(primary_hue="blue"),
    title="CampusPulse RAG Assistant",
) as demo:

    gr.Markdown(
        "# 🎓 CampusPulse RAG Assistant\n"
        "Ask questions grounded in college records from your Excel workbook."
    )

    excel_file = gr.File(
        label="Upload College Excel Workbook",
        file_types=[".xlsx", ".xls"],
        type="filepath",
    )

    chatbot = gr.Chatbot(
        label="College Q&A",
        type="messages",
        height=420,
    )

    question = gr.Textbox(
        label="Your question",
        placeholder="Compare colleges in Bengaluru...",
        lines=2,
    )

    with gr.Row():
        ask_button = gr.Button("Ask CampusPulse", variant="primary")
        clear_button = gr.Button("Clear Chat")

    ask_button.click(
        respond,
        inputs=[question, excel_file, chatbot],
        outputs=[gr.Textbox(visible=False), chatbot],
    )

    question.submit(
        respond,
        inputs=[question, excel_file, chatbot],
        outputs=[gr.Textbox(visible=False), chatbot],
    )

    clear_button.click(lambda: [], outputs=chatbot)

demo.launch(
    server_name="0.0.0.0",
    server_port=int(os.getenv("PORT", "7860")),
)
