# AI Systematic Review Extraction App

This is a local Streamlit app for extracting systematic-review information from batches of research article PDFs.

The app is designed for personal research use:

- Upload multiple PDF files.
- Enter extraction fields and research questions in the browser.
- Use an OpenAI-compatible API endpoint.
- Export results as an Excel `.xlsx` workbook.
- Extract research-question evidence with an exhaustive strategy rather than a fixed excerpt limit.
- Run MMAT 2018 quality assessment as a separate step or together with extraction.
- View, edit, and restore separate AI prompt templates for extraction and MMAT appraisal.

## Setup

```bash
/opt/homebrew/bin/python3.12 -m venv .venv312
source .venv312/bin/activate
pip install -r requirements.txt
```

## Run

```bash
.venv312/bin/streamlit run app.py
```

Then open the local URL shown by Streamlit.

## API Settings

The app asks for:

- API key: entered in the browser and not saved by the app.
- Base URL: defaults to `https://api.openai.com/v1`.
- Model: defaults to `gpt-5.5`.

You can change the base URL later if you use another OpenAI-compatible provider.

## Notes

- Each PDF is processed as one article record.
- If one PDF fails, the batch continues.
- Missing information should be reported as `not found`, not guessed.
- Confidence cells marked `medium` or `low` are highlighted in the Excel export for review.
- MMAT response cells marked `No` or `Can't tell` are highlighted for review.
- If the extraction fields and research questions stay the same, the summary sheet keeps the same column structure.
- The Excel export includes extraction sheets, MMAT quality assessment sheets, and a `Methodology Prompt` sheet with the actual prompts used.

## Quality Assessment / MMAT

The MMAT workflow follows the 2018 tool:

- Every PDF gets the two MMAT screening questions.
- The app asks the AI to choose one MMAT study design category for suitable empirical primary studies.
- The app then asks only the five criteria for that chosen category.
- The app uses `Yes`, `No`, and `Can't tell`; it does not calculate an overall MMAT score.

Use:

- `Run extraction` to run only the extraction step.
- `Run quality assessment` to run only MMAT.
- `Run full workflow` to run extraction and MMAT for the same uploaded PDFs.
