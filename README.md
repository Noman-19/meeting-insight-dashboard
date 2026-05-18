# Workforce Insight Analyzer

This Streamlit app analyzes meeting transcripts and supporting documents to generate data-supported workforce and curriculum insights.

## What the App Does

- Uploads three input documents:
  - Goals/Objectives document
  - Participant list document
  - Meeting transcript document
- Extracts speaker responses from the transcript.
- Detects recurring themes such as workforce readiness, applied learning, sustainability competencies, systems/data skills, regulation/safety, and collaboration opportunities.
- Compares academic and industry perspectives.
- Shows charts, tables, evidence excerpts, and sector signals.
- Generates a professor-ready report that can be downloaded as Markdown.
- Exports extracted responses and theme counts as CSV files.

## How to Run

```bash
conda activate fyp_env
cd /d "E:\professor_workforce_analyzer"
streamlit run streamlit_app.py
```

The sidebar can automatically load the sample documents from `C:\Users\Latitude\Downloads` if they are available. To analyze another meeting, turn off the sample option and upload new `.docx` or `.txt` files.

## Required Packages

The app uses:

- streamlit
- pandas

No API key is required.

