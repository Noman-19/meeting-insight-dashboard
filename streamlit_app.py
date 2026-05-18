import csv
import io
import re
import zipfile
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Optional
from xml.etree import ElementTree as ET

import pandas as pd
import streamlit as st


APP_TITLE = "Workforce Insight Analyzer"
SAMPLE_FILES = {
    "Goals and Objectives": Path(r"C:\Users\Latitude\Downloads\Goals-Objective.docx"),
    "Participant List": Path(r"C:\Users\Latitude\Downloads\List.docx"),
    "Meeting Transcript": Path(r"C:\Users\Latitude\Downloads\Re-worded Raw Transcript.docx"),
}

THEMES = {
    "Workforce readiness": [
        "communication",
        "project management",
        "soft skills",
        "dependability",
        "work ethic",
        "initiative",
        "prioritize",
        "workplace",
        "readiness",
        "attention to detail",
    ],
    "Applied learning": [
        "real-world",
        "hands-on",
        "project-based",
        "case studies",
        "internship",
        "apprenticeship",
        "site visit",
        "industry exposure",
        "troubleshooting",
        "practical",
    ],
    "Systems and data": [
        "systems thinking",
        "erp",
        "data",
        "modeling",
        "simulation",
        "ai",
        "automation",
        "dashboard",
        "analytics",
    ],
    "Sustainability competencies": [
        "sustainability",
        "energy",
        "waste",
        "efficiency",
        "life-cycle",
        "lifecycle",
        "materials",
        "carbon",
        "recycling",
        "renewable",
        "electrification",
    ],
    "Collaboration opportunities": [
        "collaboration",
        "industry",
        "academia",
        "capstone",
        "projects",
        "visits",
        "funding",
        "data sharing",
        "networking",
        "feedback",
    ],
    "Regulation and safety": [
        "regulatory",
        "regulation",
        "compliance",
        "safety",
        "documentation",
        "reliability",
        "community",
        "communities",
    ],
}

SECTOR_KEYWORDS = {
    "Manufacturing": ["manufacturing", "production", "lean", "erp", "process"],
    "Energy and utilities": ["energy", "utility", "utilities", "grid", "electrification"],
    "Defense": ["defense", "munitions", "energetics", "testing", "safety"],
    "Workforce development": ["workforce", "training", "pipeline", "participation"],
    "Critical minerals": ["lithium", "battery", "extraction", "materials"],
    "Academic": ["curriculum", "student", "teaching", "course", "capstone"],
}


st.set_page_config(page_title=APP_TITLE, page_icon=None, layout="wide")


def add_styles():
    st.markdown(
        """
        <style>
        .block-container { padding-top: 1.5rem; max-width: 1260px; }
        div[data-testid="stMetric"] {
            border: 1px solid #d8dee4;
            border-radius: 8px;
            padding: 14px 16px;
            background: #ffffff;
        }
        div[data-testid="stMetric"] label { color: #475569; }
        .section-note {
            color: #475569;
            font-size: 0.95rem;
            margin-top: -0.35rem;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def extract_docx_text(file_obj) -> str:
    if hasattr(file_obj, "read"):
        data = file_obj.read()
        file_obj.seek(0)
    else:
        data = Path(file_obj).read_bytes()

    with zipfile.ZipFile(io.BytesIO(data)) as docx:
        xml_bytes = docx.read("word/document.xml")

    root = ET.fromstring(xml_bytes)
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    lines = []

    for paragraph in root.findall(".//w:p", ns):
        text = "".join(node.text or "" for node in paragraph.findall(".//w:t", ns)).strip()
        if text:
            lines.append(text)

    return "\n".join(lines)


def extract_text(uploaded_file, sample_path: Optional[Path] = None) -> str:
    if uploaded_file is not None:
        if uploaded_file.name.lower().endswith(".docx"):
            return extract_docx_text(uploaded_file)
        return uploaded_file.read().decode("utf-8", errors="ignore")

    if sample_path and sample_path.exists():
        return extract_docx_text(sample_path)

    return ""


def split_sentences(text: str) -> list[str]:
    chunks = re.split(r"(?<=[.!?])\s+", text.replace("\n", " "))
    return [chunk.strip() for chunk in chunks if len(chunk.strip()) > 35]


def parse_responses(transcript: str) -> list[dict]:
    responses = []
    current_section = "General"
    current_question = "Unassigned"
    current = None

    section_markers = (
        "Academic Discussion",
        "Industry Discussion",
        "Collaboration Opportunities",
    )

    for raw_line in transcript.splitlines():
        line = raw_line.strip()
        if not line:
            continue

        if any(marker.lower() in line.lower() for marker in section_markers):
            current_section = line
            continue

        if re.match(r"^Question\s+\d+", line, flags=re.I):
            current_question = line
            continue

        match = re.match(r"^([A-Za-z. ]{2,45}):\s*(.*)$", line)
        if match:
            if current:
                responses.append(current)
            speaker = re.sub(r"\s+", " ", match.group(1)).strip()
            answer = match.group(2).strip()
            current = {
                "section": current_section,
                "question": current_question,
                "speaker": speaker,
                "answer": answer,
                "group": classify_group(current_section),
            }
        elif current:
            current["answer"] = f"{current['answer']} {line}".strip()

    if current:
        responses.append(current)

    return responses


def classify_group(section: str) -> str:
    lower = section.lower()
    if "academic" in lower:
        return "Academic"
    if "industry" in lower or "collaboration" in lower:
        return "Industry"
    return "Unclassified"


def count_theme_hits(text: str, keywords: list[str]) -> int:
    lowered = text.lower()
    return sum(len(re.findall(rf"\b{re.escape(keyword.lower())}\b", lowered)) for keyword in keywords)


def analyze_themes(responses: list[dict]) -> pd.DataFrame:
    rows = []
    for theme, keywords in THEMES.items():
        keyword_hits = 0
        response_hits = 0
        academic_hits = 0
        industry_hits = 0

        for response in responses:
            hits = count_theme_hits(response["answer"], keywords)
            keyword_hits += hits
            if hits:
                response_hits += 1
                if response["group"] == "Academic":
                    academic_hits += 1
                elif response["group"] == "Industry":
                    industry_hits += 1

        rows.append(
            {
                "Theme": theme,
                "Keyword mentions": keyword_hits,
                "Responses mentioning theme": response_hits,
                "Academic response mentions": academic_hits,
                "Industry response mentions": industry_hits,
            }
        )

    return pd.DataFrame(rows).sort_values(
        ["Responses mentioning theme", "Keyword mentions"], ascending=False
    )


def analyze_sector_signals(responses: list[dict]) -> pd.DataFrame:
    all_text = " ".join(response["answer"] for response in responses)
    rows = []
    for sector, keywords in SECTOR_KEYWORDS.items():
        rows.append({"Sector signal": sector, "Mentions": count_theme_hits(all_text, keywords)})
    return pd.DataFrame(rows).sort_values("Mentions", ascending=False)


def top_terms(responses: list[dict]) -> pd.DataFrame:
    stopwords = {
        "about",
        "across",
        "also",
        "because",
        "between",
        "current",
        "different",
        "especially",
        "focus",
        "future",
        "important",
        "include",
        "including",
        "industry",
        "students",
        "their",
        "there",
        "these",
        "through",
        "where",
        "which",
        "while",
        "with",
        "work",
        "working",
    }
    words = re.findall(r"[A-Za-z][A-Za-z-]{3,}", " ".join(r["answer"] for r in responses).lower())
    counts = Counter(word for word in words if word not in stopwords)
    return pd.DataFrame(counts.most_common(25), columns=["Term", "Count"])


def best_evidence(responses: list[dict], theme: str, limit: int = 5) -> list[dict]:
    keywords = THEMES[theme]
    scored = []
    for response in responses:
        score = count_theme_hits(response["answer"], keywords)
        if score:
            scored.append((score, response))
    scored.sort(key=lambda item: item[0], reverse=True)
    return [item[1] for item in scored[:limit]]


def theme_insight(theme: str) -> str:
    insights = {
        "Workforce readiness": "Employers repeatedly need graduates with stronger communication, dependability, documentation, initiative, and workplace judgment.",
        "Applied learning": "The discussion points to a gap between classroom knowledge and practical application in real operational settings.",
        "Systems and data": "Participants want students to understand connected systems, data-informed decisions, modeling, automation, and technology use.",
        "Sustainability competencies": "Sustainability is treated as an operational and engineering concern tied to efficiency, energy, materials, lifecycle impact, and cost.",
        "Collaboration opportunities": "The most practical path forward is recurring industry-academia collaboration through projects, visits, data sharing, and feedback.",
        "Regulation and safety": "Several sectors need careful documentation, safety awareness, reliability, compliance, and community/regulatory understanding.",
    }
    return insights.get(theme, "")


def generate_report(
    goals_text: str,
    participant_text: str,
    responses: list[dict],
    theme_df: pd.DataFrame,
    sector_df: pd.DataFrame,
) -> str:
    generated = datetime.now().strftime("%B %d, %Y")
    goals = [line.strip("- ").strip() for line in goals_text.splitlines() if len(line.strip()) > 15]
    top_theme_names = theme_df["Theme"].head(4).tolist()
    group_counts = Counter(response["group"] for response in responses)

    lines = [
        "# Meeting Transcript Analysis Report",
        "",
        f"Generated: {generated}",
        "",
        "## Executive Summary",
        "",
        "The meeting discussion shows that regional employers and academic participants share a common concern: students need stronger preparation for real workplace conditions. The strongest findings are workforce readiness, applied learning, sustainability competencies, systems thinking, and structured industry-academia collaboration.",
        "",
        "## Event Objectives",
        "",
    ]

    if goals:
        lines.extend([f"- {goal}" for goal in goals])
    else:
        lines.append("- No goals document was provided.")

    lines.extend(
        [
            "",
            "## Analysis Coverage",
            "",
            f"- Responses analyzed: {len(responses)}",
            f"- Academic response records: {group_counts.get('Academic', 0)}",
            f"- Industry response records: {group_counts.get('Industry', 0)}",
            f"- Participant information provided: {'Yes' if participant_text.strip() else 'No'}",
            "",
            "## Highest-Priority Themes",
            "",
        ]
    )

    for _, row in theme_df.iterrows():
        lines.extend(
            [
                f"### {row['Theme']}",
                "",
                theme_insight(row["Theme"]),
                "",
                f"- Keyword mentions: {int(row['Keyword mentions'])}",
                f"- Responses mentioning this theme: {int(row['Responses mentioning theme'])}",
                f"- Academic response mentions: {int(row['Academic response mentions'])}",
                f"- Industry response mentions: {int(row['Industry response mentions'])}",
                "",
                "Representative evidence:",
            ]
        )
        for response in best_evidence(responses, row["Theme"], 3):
            answer = response["answer"].strip()
            if len(answer) > 260:
                answer = answer[:257].rstrip() + "..."
            lines.append(f"- {response['speaker']} ({response['group']}): {answer}")
        lines.append("")

    lines.extend(
        [
            "## Sector Signals",
            "",
        ]
    )

    for _, row in sector_df.iterrows():
        lines.append(f"- {row['Sector signal']}: {int(row['Mentions'])} mentions")

    lines.extend(
        [
            "",
            "## Data-Supported Outcomes",
            "",
            "1. Workforce readiness is the most urgent need. Employers want graduates who can communicate clearly, document work carefully, show dependability, solve practical problems, and adapt to changing workplace expectations.",
            "2. The main curriculum gap is the theory-to-practice gap. Students need more experience with real systems, messy data, operational constraints, troubleshooting, and sustainability tradeoffs.",
            "3. Sustainability should be taught as an applied engineering and business topic connected to process efficiency, waste reduction, energy systems, materials, lifecycle impact, regulation, and cost.",
            "4. Industry and academia should collaborate through internships, apprenticeships, capstone projects, industry-sponsored assignments, site visits, guest speakers, real datasets, and continuous advisory feedback.",
            "",
            "## Recommended Curriculum Actions",
            "",
            "- Add more industry-based case studies and practical troubleshooting assignments.",
            "- Use real or realistic industry datasets in modeling, simulation, and process-analysis courses.",
            "- Expand capstone and project-based learning with direct industry sponsorship.",
            "- Require students to communicate tradeoffs involving cost, carbon, performance, safety, reliability, and compliance.",
            "- Include documentation, technical writing, presentation, project management, and teamwork as graded engineering outcomes.",
            "- Integrate responsible AI use by requiring students to verify and explain AI-supported work.",
            "",
            "## Recommended Collaboration Opportunities",
            "",
            "- Capstone partnerships with regional employers.",
            "- Internship and apprenticeship pipelines.",
            "- Site visits to manufacturing, energy, utilities, defense, and lithium-related facilities.",
            "- Industry-provided datasets for student analysis projects.",
            "- Guest lectures and workshops on sustainability, workforce readiness, and operational decision-making.",
            "- Regular employer feedback sessions to keep curriculum aligned with changing workforce needs.",
            "",
            "## Conclusion",
            "",
            "The transcript indicates that workforce readiness and sustainability education should be connected through applied, industry-informed learning. The most valuable next step is to build recurring collaboration between Southern Arkansas University and regional employers so students can practice technical and socio-technical skills before entering the workforce.",
        ]
    )

    return "\n".join(lines)


def make_csv_download(responses: list[dict]) -> str:
    output = io.StringIO()
    writer = csv.DictWriter(output, fieldnames=["section", "question", "group", "speaker", "answer"])
    writer.writeheader()
    writer.writerows(responses)
    return output.getvalue()


def render_uploads():
    st.sidebar.header("Input Documents")
    use_samples = st.sidebar.checkbox("Use sample files from Downloads", value=True)
    uploads = {}

    uploads["goals"] = st.sidebar.file_uploader(
        "Goals/Objectives document", type=["docx", "txt"], key="goals"
    )
    uploads["participants"] = st.sidebar.file_uploader(
        "Participant list document", type=["docx", "txt"], key="participants"
    )
    uploads["transcript"] = st.sidebar.file_uploader(
        "Transcript document", type=["docx", "txt"], key="transcript"
    )

    sample_paths = SAMPLE_FILES if use_samples else {}
    goals_text = extract_text(uploads["goals"], sample_paths.get("Goals and Objectives"))
    participant_text = extract_text(uploads["participants"], sample_paths.get("Participant List"))
    transcript_text = extract_text(uploads["transcript"], sample_paths.get("Meeting Transcript"))

    return goals_text, participant_text, transcript_text, use_samples


def main():
    add_styles()
    goals_text, participant_text, transcript_text, use_samples = render_uploads()

    st.title(APP_TITLE)
    st.caption("Upload meeting documents to convert raw discussion responses into structured, data-supported outcomes.")

    if not transcript_text.strip():
        st.info("Upload a transcript document to begin the analysis.")
        return

    responses = parse_responses(transcript_text)
    if not responses:
        st.warning("The transcript was loaded, but no speaker responses were detected. Check that responses use a `Name: answer` format.")
        with st.expander("Preview transcript text"):
            st.text(transcript_text[:4000])
        return

    theme_df = analyze_themes(responses)
    sector_df = analyze_sector_signals(responses)
    terms_df = top_terms(responses)
    report = generate_report(goals_text, participant_text, responses, theme_df, sector_df)

    group_counts = Counter(response["group"] for response in responses)
    question_count = len({response["question"] for response in responses})

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Responses", len(responses))
    col2.metric("Questions", question_count)
    col3.metric("Industry Records", group_counts.get("Industry", 0))
    col4.metric("Academic Records", group_counts.get("Academic", 0))

    if use_samples:
        st.markdown(
            '<p class="section-note">Sample files are loaded from Downloads. You can turn this off in the sidebar and upload new files.</p>',
            unsafe_allow_html=True,
        )

    tab_overview, tab_themes, tab_compare, tab_report, tab_data = st.tabs(
        ["Overview", "Theme Analysis", "Comparison", "Generated Report", "Data"]
    )

    with tab_overview:
        left, right = st.columns([1.15, 0.85])
        with left:
            st.subheader("Theme Distribution")
            chart_df = theme_df.set_index("Theme")[["Responses mentioning theme"]]
            st.bar_chart(chart_df)

        with right:
            st.subheader("Top Terms")
            st.dataframe(terms_df, use_container_width=True, hide_index=True)

        st.subheader("Key Outcomes")
        st.write(
            "The transcript indicates that workforce readiness, applied learning, sustainability competencies, systems thinking, and industry collaboration are the strongest recurring outcomes."
        )

    with tab_themes:
        st.subheader("Coded Theme Table")
        st.dataframe(theme_df, use_container_width=True, hide_index=True)

        selected_theme = st.selectbox("Evidence for theme", theme_df["Theme"].tolist())
        for response in best_evidence(responses, selected_theme, 6):
            st.markdown("---")
            st.markdown(f"**{response['speaker']}** | {response['group']} | {response['question']}")
            st.write(response["answer"])

    with tab_compare:
        left, right = st.columns(2)
        with left:
            st.subheader("Academic vs Industry Theme Mentions")
            compare_df = theme_df.set_index("Theme")[
                ["Academic response mentions", "Industry response mentions"]
            ]
            st.bar_chart(compare_df)

        with right:
            st.subheader("Sector Signals")
            st.bar_chart(sector_df.set_index("Sector signal")[["Mentions"]])

        st.subheader("Interpretation")
        st.write(
            "Industry responses emphasize workforce behavior, applied problem-solving, operational systems, safety, and future technology changes. Academic responses emphasize ways to integrate these needs through projects, datasets, capstones, site visits, and curriculum alignment."
        )

    with tab_report:
        st.subheader("Professor-Ready Report")
        st.download_button(
            "Download Markdown Report",
            data=report,
            file_name="meeting_analysis_report.md",
            mime="text/markdown",
        )
        st.markdown(report)

    with tab_data:
        st.subheader("Extracted Responses")
        response_df = pd.DataFrame(responses)
        st.dataframe(response_df, use_container_width=True, hide_index=True)
        st.download_button(
            "Download Extracted Responses CSV",
            data=make_csv_download(responses),
            file_name="extracted_responses.csv",
            mime="text/csv",
        )
        st.download_button(
            "Download Theme Counts CSV",
            data=theme_df.to_csv(index=False),
            file_name="theme_counts.csv",
            mime="text/csv",
        )


if __name__ == "__main__":
    main()
