# AI Systematic Review Extraction App

This is a local Streamlit app for extracting systematic-review information from batches of research article PDFs.

The app is designed for personal research use:

- Upload multiple PDF files.
- Enter extraction fields and research questions in the browser.
- Use an OpenAI-compatible API endpoint.
- Export results as an Excel `.xlsx` workbook.
- Upload RIS and PubMed NBIB citation files for pre-screening deduplication.
- Mark citation records against user-provided inclusion criteria before full-text extraction.
- Export citation screening results as an Excel audit file, two standard RIS files, and a JSON backup that can be uploaded later to restore results.
- Extract research-question evidence with an exhaustive strategy rather than a fixed excerpt limit.
- Export highlighted copies of processed PDFs, with different research questions marked in different colors.
- Run MMAT 2018 quality assessment as a separate step or together with extraction.
- View, edit, and restore separate AI prompt templates for extraction and MMAT appraisal.
- Use a protected MMAT 2018 rubric inserted by the app, with the official manual available for download inside the interface.

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
- The highlighted PDF export creates a zip file with marked PDFs, a color legend, and a match report for excerpts that could not be located in the PDF text layer.
- Citation screening exports use timestamped file names in `YYYYMMDD_HHMM` format.
- Duplicate citation records are removed from the main screening result, but kept in the Excel duplicate log for traceability.
- AI inclusion marking is conservative and only flags records; it does not delete records from the screening Excel.
- The AI citation inclusion prompt is visible and editable in the sidebar.
- AI inclusion marking runs in small batches and sends the original title and original abstract only, without truncating either field.

## Citation Screening

Use the `Citation screening` section to upload `.ris` and PubMed `.nbib` files before PDF extraction.

The deduplication logic is:

- Matching DOI or PMID means the later record is removed as a duplicate.
- If DOI/PMID cannot identify a duplicate, title similarity must be at least 95%, and either abstract sequence similarity or abstract token overlap must be at least 95% before the later record is removed.
- PubMed NBIB records are split by each `PMID-` record start.
- The page and Excel export include an import log showing how many records were parsed from each uploaded citation file.

The export buttons create:

- `citation_screening_audit_YYYYMMDD_HHMM.xlsx`: screening results, duplicate log, and methodology details.
- `ai_suggested_inclusion_records_YYYYMMDD_HHMM.ris`: records marked as matching or potentially matching the inclusion criteria.
- `all_screening_records_YYYYMMDD_HHMM.ris`: all current screening records.
- `citation_screening_backup_YYYYMMDD_HHMM.json`: a restorable backup for the citation screening state.

For shared Streamlit Cloud deployments, use the JSON backup download and restore controls as the reliable recovery path. Server-side citation autosave is disabled by default so that multiple users do not recover or overwrite each other's screening state. To enable local-only autosave while running the app on your own computer, set `AQEREVIEW_ENABLE_LOCAL_CITATION_AUTOSAVE=1` before starting Streamlit.

Use `AI mark` when you want to mark the current uploaded records without deduplication first. Use `Deduplicate + AI mark` when you want the app to perform citation deduplication and AI inclusion marking in one step.

## Quality Assessment / MMAT

The MMAT workflow follows the 2018 tool:

- Every PDF gets the two MMAT screening questions.
- The app asks the AI to choose one MMAT study design category for suitable empirical primary studies.
- The app then asks only the five criteria for that chosen category.
- The app uses `Yes`, `No`, and `Can't tell`; it does not calculate an overall MMAT score.
- The editable MMAT prompt is only extra instruction text. The app always appends a protected MMAT 2018 rubric based on the bundled manual, so the official criteria are not removed by prompt edits.
- The bundled manual is stored at `assets/MMAT__criteria-manual_2018-08-01_ENG.pdf` and can be downloaded from the MMAT prompt area.

Use:

- `Run extraction` to run only the extraction step.
- `Run quality assessment` to run only MMAT.
- `Run full workflow` to run extraction and MMAT for the same uploaded PDFs.
