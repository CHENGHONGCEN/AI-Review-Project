# AI Systematic Review Extraction App

This is a local Streamlit app for extracting systematic-review information from batches of research article PDFs.

The app is designed for personal research use:

- Upload multiple PDF files.
- Enter extraction fields and research questions in the browser.
- Use an OpenAI-compatible API endpoint.
- Export results as an Excel `.xlsx` workbook.
- Extract research-question evidence with an exhaustive strategy rather than a fixed excerpt limit.
- View, edit, and restore the AI prompt template for transparency and reproducibility.

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
- Low-confidence outputs are highlighted and included in the Excel export.
- The Excel export includes a `Methodology Prompt` sheet with the actual prompt used for the extraction.
