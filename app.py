import base64
import io
import json
from datetime import datetime
from html import escape
from typing import Any

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openai import OpenAI


DEFAULT_BASE_URL = "https://api.openai.com/v1"
DEFAULT_MODEL = "gpt-5.5"
CONFIDENCE_LEVELS = ["high", "medium", "low"]
DEFAULT_PROMPT_TEMPLATE = """
You are helping with systematic review data extraction.

Read the attached research article PDF and extract information for one article record.

Important rules:
- Do not guess. If something is absent or unclear, write "not found".
- Preserve original wording in evidence excerpts.
- Use an exhaustive extraction strategy for research-question evidence.
- For each research question, extract every relevant original-text excerpt you can find.
- Err on the side of including more potentially relevant excerpts rather than fewer.
- Prefer excerpts that are useful for later thematic analysis, but do not omit borderline relevant text merely to keep the list short.
- Include page numbers or section names in source_location when you can identify them.
- Mark confidence as "low" when the article is scanned poorly, evidence is indirect, pages are missing, or the answer is uncertain.
- Use low_confidence_reason to explain any medium or low confidence output in plain language.

Structured fields requested by the user:
{structured_fields}

Research questions requested by the user:
{research_questions}
""".strip()


EXTRACTION_SCHEMA: dict[str, Any] = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "article": {
            "type": "object",
            "additionalProperties": False,
            "properties": {
                "title": {"type": "string"},
                "authors": {"type": "string"},
                "year": {"type": "string"},
                "journal": {"type": "string"},
                "overall_confidence": {"type": "string", "enum": CONFIDENCE_LEVELS},
                "low_confidence_reason": {"type": "string"},
            },
            "required": [
                "title",
                "authors",
                "year",
                "journal",
                "overall_confidence",
                "low_confidence_reason",
            ],
        },
        "structured_fields": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "name": {"type": "string"},
                    "value": {"type": "string"},
                    "source_location": {"type": "string"},
                    "confidence": {"type": "string", "enum": CONFIDENCE_LEVELS},
                    "low_confidence_reason": {"type": "string"},
                },
                "required": [
                    "name",
                    "value",
                    "source_location",
                    "confidence",
                    "low_confidence_reason",
                ],
            },
        },
        "research_question_evidence": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "question": {"type": "string"},
                    "answer_summary": {"type": "string"},
                    "confidence": {"type": "string", "enum": CONFIDENCE_LEVELS},
                    "low_confidence_reason": {"type": "string"},
                    "excerpts": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "additionalProperties": False,
                            "properties": {
                                "text": {"type": "string"},
                                "source_location": {"type": "string"},
                                "relevance_note": {"type": "string"},
                            },
                            "required": ["text", "source_location", "relevance_note"],
                        },
                    },
                },
                "required": [
                    "question",
                    "answer_summary",
                    "confidence",
                    "low_confidence_reason",
                    "excerpts",
                ],
            },
        },
        "review_warnings": {
            "type": "array",
            "items": {"type": "string"},
        },
    },
    "required": [
        "article",
        "structured_fields",
        "research_question_evidence",
        "review_warnings",
    ],
}


def split_lines(text: str) -> list[str]:
    return [line.strip() for line in text.splitlines() if line.strip()]


def make_prompt(fields: list[str], questions: list[str], prompt_template: str) -> str:
    field_list = "\n".join(f"- {field}" for field in fields) or "- No extra structured fields"
    question_list = "\n".join(
        f"- RQ{index}: {question}" for index, question in enumerate(questions, start=1)
    ) or "- No research questions"
    template = prompt_template.strip() or DEFAULT_PROMPT_TEMPLATE
    if "{structured_fields}" not in template:
        template = f"{template}\n\nStructured fields requested by the user:\n{{structured_fields}}"
    if "{research_questions}" not in template:
        template = f"{template}\n\nResearch questions requested by the user:\n{{research_questions}}"
    return template.format(
        structured_fields=field_list,
        research_questions=question_list,
    ).strip()


def pdf_to_input_file(uploaded_file: Any) -> dict[str, str]:
    encoded = base64.b64encode(uploaded_file.getvalue()).decode("utf-8")
    return {
        "type": "input_file",
        "filename": uploaded_file.name,
        "file_data": f"data:application/pdf;base64,{encoded}",
    }


def extract_from_pdf(
    uploaded_file: Any,
    api_key: str,
    base_url: str,
    model: str,
    fields: list[str],
    questions: list[str],
    prompt_template: str,
) -> dict[str, Any]:
    client = OpenAI(api_key=api_key, base_url=base_url.rstrip("/"))
    prompt = make_prompt(fields, questions, prompt_template)

    response = client.responses.create(
        model=model,
        input=[
            {
                "role": "user",
                "content": [
                    pdf_to_input_file(uploaded_file),
                    {"type": "input_text", "text": prompt},
                ],
            }
        ],
        text={
            "format": {
                "type": "json_schema",
                "name": "systematic_review_extraction",
                "strict": True,
                "schema": EXTRACTION_SCHEMA,
            }
        },
    )

    raw_text = response.output_text
    data = json.loads(raw_text)
    data["source_file"] = uploaded_file.name
    data["prompt_used"] = prompt
    return data


def result_to_flat_row(result: dict[str, Any]) -> dict[str, str]:
    article = result.get("article", {})
    row = {
        "File names": result.get("source_file", ""),
        "title": article.get("title", ""),
        "authors": article.get("authors", ""),
        "year": article.get("year", ""),
        "journal": article.get("journal", ""),
        "overall_confidence": article.get("overall_confidence", ""),
        "low_confidence_reason": article.get("low_confidence_reason", ""),
        "review_warnings": "; ".join(result.get("review_warnings", [])),
    }

    for item in result.get("structured_fields", []):
        name = item.get("name", "field")
        row[name] = item.get("value", "")
        row[f"{name} confidence"] = item.get("confidence", "")

    for evidence in result.get("research_question_evidence", []):
        question = evidence.get("question", "research question")
        row[f"{question} summary"] = evidence.get("answer_summary", "")
        row[f"{question} confidence"] = evidence.get("confidence", "")

    return row


def confidence_needs_review(value: str) -> bool:
    return value.lower() in {"low", "medium"}


def style_results(df: pd.DataFrame) -> Any:
    def highlight(value: Any) -> str:
        if isinstance(value, str) and confidence_needs_review(value):
            return "background-color: #ffe5e5; color: #7a1f1f;"
        return ""

    return df.style.map(highlight)


def result_to_evidence_rows(result: dict[str, Any]) -> list[dict[str, str]]:
    article = result.get("article", {})
    rows = []
    for question_index, evidence in enumerate(result.get("research_question_evidence", []), start=1):
        excerpts = evidence.get("excerpts", []) or [{"text": "", "source_location": "", "relevance_note": ""}]
        for excerpt in excerpts:
            rows.append(
                {
                    "File names": result.get("source_file", ""),
                    "title": article.get("title", ""),
                    "research question": f"RQ{question_index}",
                    "answer_summary": evidence.get("answer_summary", ""),
                    "confidence": evidence.get("confidence", ""),
                    "excerpt": excerpt.get("text", ""),
                    "source_location": excerpt.get("source_location", ""),
                    "relevance_note": excerpt.get("relevance_note", ""),
                }
            )
    return rows


def add_rows_to_sheet(sheet: Any, rows: list[dict[str, str]]) -> None:
    if not rows:
        return
    headers = []
    for row in rows:
        for key in row:
            if key not in headers:
                headers.append(key)
    sheet.append(headers)
    for row in rows:
        sheet.append([row.get(header, "") for header in headers])


def tune_excel_sheet(sheet: Any) -> None:
    header_fill = PatternFill("solid", fgColor="1F2937")
    header_font = Font(color="FFFFFF", bold=True)
    review_fill = PatternFill("solid", fgColor="FCE4E4")
    border = Border(
        left=Side(style="thin", color="D1D5DB"),
        right=Side(style="thin", color="D1D5DB"),
        top=Side(style="thin", color="D1D5DB"),
        bottom=Side(style="thin", color="D1D5DB"),
    )

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions

    for cell in sheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    for row in sheet.iter_rows(min_row=2):
        row_needs_review = any(
            isinstance(cell.value, str) and confidence_needs_review(cell.value)
            for cell in row
        )
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = border
            if row_needs_review:
                cell.fill = review_fill

    for column_cells in sheet.columns:
        column_letter = get_column_letter(column_cells[0].column)
        longest_line = 0
        for cell in column_cells:
            value = "" if cell.value is None else str(cell.value)
            for line in value.splitlines() or [""]:
                longest_line = max(longest_line, len(line))
        sheet.column_dimensions[column_letter].width = min(max(longest_line + 2, 12), 60)

    for row in sheet.iter_rows():
        max_lines = 1
        for cell in row:
            value = "" if cell.value is None else str(cell.value)
            wrapped_lines = max(1, len(value) // 55 + 1)
            explicit_lines = value.count("\n") + 1
            max_lines = max(max_lines, wrapped_lines, explicit_lines)
        sheet.row_dimensions[row[0].row].height = min(max(18, max_lines * 15), 120)


def merge_repeated_evidence_cells(sheet: Any) -> None:
    if sheet.max_row < 3:
        return

    merge_columns = [1, 2, 3, 4, 5]
    group_start = 2
    previous_key = None

    for row_index in range(2, sheet.max_row + 2):
        if row_index <= sheet.max_row:
            current_key = (
                sheet.cell(row=row_index, column=1).value,
                sheet.cell(row=row_index, column=2).value,
                sheet.cell(row=row_index, column=3).value,
            )
        else:
            current_key = None

        if previous_key is None:
            previous_key = current_key
            continue

        if current_key != previous_key:
            group_end = row_index - 1
            if group_end > group_start:
                for column_index in merge_columns:
                    sheet.merge_cells(
                        start_row=group_start,
                        start_column=column_index,
                        end_row=group_end,
                        end_column=column_index,
                    )
                    sheet.cell(row=group_start, column=column_index).alignment = Alignment(
                        vertical="center",
                        wrap_text=True,
                    )
            group_start = row_index
            previous_key = current_key


def build_excel_export(results: list[dict[str, Any]]) -> bytes:
    workbook = Workbook()
    summary_sheet = workbook.active
    summary_sheet.title = "Article Summary"
    summary_sheet["A1"] = "No extraction results"

    summary_rows = [result_to_flat_row(result) for result in results]
    evidence_rows = []
    for result in results:
        evidence_rows.extend(result_to_evidence_rows(result))

    if summary_rows:
        summary_sheet.delete_rows(1, 1)
        add_rows_to_sheet(summary_sheet, summary_rows)
        tune_excel_sheet(summary_sheet)

    evidence_sheet = workbook.create_sheet("Evidence Excerpts")
    if evidence_rows:
        add_rows_to_sheet(evidence_sheet, evidence_rows)
    else:
        evidence_sheet.append(["File names", "title", "research question", "answer_summary", "confidence", "excerpt", "source_location", "relevance_note"])
    merge_repeated_evidence_cells(evidence_sheet)
    tune_excel_sheet(evidence_sheet)

    methodology_sheet = workbook.create_sheet("Methodology Prompt")
    methodology_sheet.append(["item", "value"])
    methodology_sheet.append(["Generated", datetime.now().strftime("%Y-%m-%d %H:%M")])
    methodology_sheet.append(["Prompt used", results[0].get("prompt_used", "not recorded") if results else "not recorded"])
    methodology_sheet.append(["Prompt note", "This is the actual prompt text sent to the AI model after inserting the current structured fields and research questions."])
    tune_excel_sheet(methodology_sheet)
    methodology_sheet.column_dimensions["A"].width = 22
    methodology_sheet.column_dimensions["B"].width = 100
    methodology_sheet.row_dimensions[3].height = 240

    output = io.BytesIO()
    workbook.save(output)
    return output.getvalue()


def initialise_state() -> None:
    st.session_state.setdefault("results", [])
    st.session_state.setdefault("errors", [])
    st.session_state.setdefault("research_questions", [""])
    st.session_state.setdefault("prompt_template", DEFAULT_PROMPT_TEMPLATE)


def apply_custom_style() -> None:
    st.markdown(
        """
        <style>
        :root {
            --app-bg: #ffffff;
            --panel-bg: #ffffff;
            --sidebar-bg: #f7f7f8;
            --text-main: #0d0d0d;
            --text-muted: #6b7280;
            --border: #e5e7eb;
            --border-strong: #d1d5db;
            --accent: #10a37f;
            --accent-dark: #0d8f70;
            --danger: #ef4444;
        }

        .stApp {
            background: var(--app-bg);
            color: var(--text-main);
        }

        [data-testid="stSidebar"] {
            background: var(--sidebar-bg);
            border-right: 1px solid var(--border);
        }

        [data-testid="stSidebar"] h2,
        [data-testid="stSidebar"] h3 {
            color: var(--text-main);
            letter-spacing: 0;
        }

        .main .block-container {
            max-width: 1040px;
            padding-top: 2.4rem;
            padding-bottom: 3rem;
        }

        h1 {
            color: var(--text-main);
            font-weight: 720;
            font-size: 2.35rem;
            letter-spacing: 0;
        }

        div[data-testid="stCaptionContainer"] {
            color: var(--text-muted);
        }

        .stTextInput input,
        .stTextArea textarea {
            border-radius: 8px;
            border: 1px solid var(--border-strong);
            background: #ffffff;
            color: var(--text-main);
        }

        .stTextInput input:focus,
        .stTextArea textarea:focus {
            border-color: var(--accent);
            box-shadow: 0 0 0 1px var(--accent);
        }

        .stButton button {
            border-radius: 8px;
            border: 1px solid var(--border-strong);
            background: #ffffff;
            color: var(--text-main);
            font-weight: 650;
        }

        .stButton button:hover {
            border-color: #9ca3af;
            color: var(--text-main);
            background: #f9fafb;
        }

        .stButton button[kind="primary"] {
            background: var(--text-main);
            border-color: var(--text-main);
            color: #ffffff;
        }

        .stButton button[kind="primary"]:hover {
            background: #2f2f2f;
            border-color: #2f2f2f;
            color: #ffffff;
        }

        div[data-testid="stFileUploader"] section {
            border-radius: 10px;
            border-color: var(--border);
            background: #fafafa;
        }

        div[data-testid="stDataFrame"] {
            border: 1px solid var(--border);
            border-radius: 8px;
            overflow: hidden;
        }

        .rq-label {
            font-weight: 700;
            color: #374151;
            padding-top: 0.45rem;
        }

        .section-note {
            color: var(--text-muted);
            font-size: 0.9rem;
            margin-top: -0.35rem;
            margin-bottom: 0.75rem;
        }

        .app-hero {
            margin-bottom: 1.4rem;
        }

        .app-kicker {
            display: inline-flex;
            align-items: center;
            gap: 0.45rem;
            color: var(--text-muted);
            font-size: 0.92rem;
            margin-bottom: 0.55rem;
        }

        .app-kicker svg,
        .panel-title svg,
        .metric-card svg {
            width: 18px;
            height: 18px;
            stroke: currentColor;
            stroke-width: 1.9;
            fill: none;
            stroke-linecap: round;
            stroke-linejoin: round;
        }

        .workspace-panel {
            border: 1px solid var(--border);
            background: var(--panel-bg);
            border-radius: 12px;
            padding: 1.1rem 1.1rem 0.25rem;
            box-shadow: 0 1px 2px rgba(0, 0, 0, 0.04);
            margin-bottom: 1.1rem;
        }

        .panel-title {
            display: flex;
            align-items: center;
            gap: 0.55rem;
            font-weight: 720;
            color: var(--text-main);
            margin-bottom: 0.2rem;
        }

        .panel-subtitle {
            color: var(--text-muted);
            font-size: 0.94rem;
            margin-bottom: 1rem;
        }

        .metric-grid {
            display: grid;
            grid-template-columns: repeat(4, minmax(0, 1fr));
            gap: 0.75rem;
            margin: 0.8rem 0 1rem;
        }

        .metric-card {
            border: 1px solid var(--border);
            border-radius: 10px;
            padding: 0.82rem;
            background: #fcfcfc;
            min-height: 82px;
        }

        .metric-card svg {
            color: var(--accent);
            margin-bottom: 0.35rem;
        }

        .metric-label {
            color: var(--text-muted);
            font-size: 0.78rem;
            line-height: 1.2;
        }

        .metric-value {
            color: var(--text-main);
            font-weight: 720;
            font-size: 1.05rem;
            margin-top: 0.16rem;
            overflow-wrap: anywhere;
        }

        .soft-divider {
            height: 1px;
            background: var(--border);
            margin: 0.35rem 0 1rem;
        }

        @media (max-width: 900px) {
            .metric-grid {
                grid-template-columns: repeat(2, minmax(0, 1fr));
            }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def svg_icon(name: str) -> str:
    icons = {
        "spark": '<path d="M12 3l1.8 4.2L18 9l-4.2 1.8L12 15l-1.8-4.2L6 9l4.2-1.8L12 3z"/><path d="M5 14l.9 2.1L8 17l-2.1.9L5 20l-.9-2.1L2 17l2.1-.9L5 14z"/>',
        "file": '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/><path d="M8 13h8"/><path d="M8 17h5"/>',
        "list": '<path d="M8 6h13"/><path d="M8 12h13"/><path d="M8 18h13"/><path d="M3 6h.01"/><path d="M3 12h.01"/><path d="M3 18h.01"/>',
        "search": '<circle cx="11" cy="11" r="8"/><path d="M21 21l-4.3-4.3"/>',
        "sheet": '<path d="M4 4h16v16H4z"/><path d="M4 10h16"/><path d="M10 4v16"/>',
        "settings": '<path d="M12 15.5A3.5 3.5 0 1 0 12 8a3.5 3.5 0 0 0 0 7.5z"/><path d="M19.4 15a1.7 1.7 0 0 0 .3 1.9l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.9-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1-1.5 1.7 1.7 0 0 0-1.9.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.9 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1 1.7 1.7 0 0 0-.3-1.9l-.1-.1A2 2 0 1 1 7 4.2l.1.1a1.7 1.7 0 0 0 1.9.3 1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.9-.3l.1-.1A2 2 0 1 1 19.8 7l-.1.1a1.7 1.7 0 0 0-.3 1.9 1.7 1.7 0 0 0 1.5 1h.1a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1z"/>',
    }
    return f'<svg viewBox="0 0 24 24" aria-hidden="true">{icons[name]}</svg>'


def render_header() -> None:
    st.markdown(
        f"""
        <div class="app-hero">
            <div class="app-kicker">{svg_icon("spark")} Systematic review extraction workspace</div>
            <h1>AI Systematic Review Extraction</h1>
            <div style="color:#6b7280; max-width:760px; line-height:1.55;">
                Batch extract article metadata, structured fields, and exhaustive research-question evidence from PDFs.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_workspace_panel(
    uploaded_count: int,
    field_count: int,
    question_count: int,
    model: str,
) -> None:
    st.markdown(
        f"""
        <div class="workspace-panel">
            <div class="panel-title">{svg_icon("search")} Extraction run</div>
            <div class="panel-subtitle">Current setup for the next batch extraction.</div>
            <div class="metric-grid">
                <div class="metric-card">
                    {svg_icon("file")}
                    <div class="metric-label">PDF files</div>
                    <div class="metric-value">{uploaded_count}</div>
                </div>
                <div class="metric-card">
                    {svg_icon("list")}
                    <div class="metric-label">Structured fields</div>
                    <div class="metric-value">{field_count}</div>
                </div>
                <div class="metric-card">
                    {svg_icon("search")}
                    <div class="metric-label">Research questions</div>
                    <div class="metric-value">{question_count}</div>
                </div>
                <div class="metric-card">
                    {svg_icon("settings")}
                    <div class="metric-label">Model</div>
                    <div class="metric-value">{escape(model)}</div>
                </div>
            </div>
            <div class="soft-divider"></div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_settings() -> tuple[str, str, str]:
    st.subheader("API settings")
    api_key = st.text_input("API key", type="password", help="The app uses this only for the current browser session.")
    base_url = st.text_input("Base URL", value=DEFAULT_BASE_URL)
    model = st.text_input("Model", value=DEFAULT_MODEL)
    return api_key, base_url, model


def restore_default_prompt() -> None:
    st.session_state.prompt_template = DEFAULT_PROMPT_TEMPLATE


def render_prompt_editor() -> str:
    with st.expander("Prompt transparency and editing", expanded=False):
        st.markdown(
            '<div class="section-note">This template is visible for transparency. The placeholders are filled automatically before each PDF is sent to the model.</div>',
            unsafe_allow_html=True,
        )
        st.button(
            "Restore default prompt",
            key="restore_prompt",
            help="Reset the prompt template to the built-in default.",
            on_click=restore_default_prompt,
        )
        prompt_template = st.text_area(
            "Prompt template",
            key="prompt_template",
            height=420,
            help="Keep {structured_fields} and {research_questions} if you want the app to insert the current template fields and RQs at those positions.",
        )
        missing = [
            placeholder
            for placeholder in ("{structured_fields}", "{research_questions}")
            if placeholder not in prompt_template
        ]
        if missing:
            st.info(
                "Missing placeholders will be appended automatically when extraction runs: "
                + ", ".join(missing)
            )
        return prompt_template


def add_research_question() -> None:
    st.session_state.research_questions.append("")


def delete_research_question(index: int) -> None:
    if len(st.session_state.research_questions) == 1:
        st.session_state.research_questions[0] = ""
    else:
        st.session_state.research_questions.pop(index)


def render_research_questions() -> list[str]:
    title_col, add_col = st.columns([0.78, 0.22])
    with title_col:
        st.markdown("**Research questions**")
        st.markdown(
            '<div class="section-note">Use one box for one RQ. The app will label them as RQ1, RQ2, RQ3 in the Excel export.</div>',
            unsafe_allow_html=True,
        )
    with add_col:
        st.button("+", key="add_rq", help="Add a research question", on_click=add_research_question)

    for index in range(len(st.session_state.research_questions)):
        label_col, input_col, delete_col = st.columns([0.16, 0.68, 0.16], vertical_alignment="top")
        with label_col:
            st.markdown(f'<div class="rq-label">RQ{index + 1}</div>', unsafe_allow_html=True)
        with input_col:
            st.session_state.research_questions[index] = st.text_area(
                f"RQ{index + 1}",
                value=st.session_state.research_questions[index],
                key=f"rq_text_{index}",
                height=88,
                label_visibility="collapsed",
                placeholder="Enter one research question",
            )
        with delete_col:
            st.button(
                "x",
                key=f"delete_rq_{index}",
                help=f"Delete RQ{index + 1}",
                on_click=delete_research_question,
                args=(index,),
            )

    return [question.strip() for question in st.session_state.research_questions if question.strip()]


def render_template() -> tuple[list[str], list[str]]:
    st.subheader("Extraction template")
    st.markdown(
        '<div class="section-note">The extraction now uses an exhaustive strategy: it asks the model to include all relevant evidence it can find.</div>',
        unsafe_allow_html=True,
    )
    field_text = st.text_area(
        "Structured fields, one per line",
        value="Study design\nPopulation / sample\nIntervention or exposure\nComparator\nMain findings\nLimitations",
        height=150,
    )
    questions = render_research_questions()
    return split_lines(field_text), questions


def render_results() -> None:
    if st.session_state.results:
        st.subheader("Results")
        rows = [result_to_flat_row(result) for result in st.session_state.results]
        df = pd.DataFrame(rows)
        st.dataframe(style_results(df), width="stretch")

        export_bytes = build_excel_export(st.session_state.results)
        st.download_button(
            "Download Excel export",
            data=export_bytes,
            file_name=f"systematic_review_extraction_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    if st.session_state.errors:
        st.subheader("Failed PDFs")
        for error in st.session_state.errors:
            st.error(f"{error['file']}: {error['message']}")


def main() -> None:
    st.set_page_config(page_title="AI Systematic Review Extraction", layout="wide")
    apply_custom_style()
    initialise_state()

    with st.sidebar:
        api_key, base_url, model = render_settings()
        st.divider()
        fields, questions = render_template()
        st.divider()
        prompt_template = render_prompt_editor()

    render_header()

    uploaded_files = st.file_uploader(
        "Upload PDF articles",
        type=["pdf"],
        accept_multiple_files=True,
        help="Select multiple PDFs. Browser folder upload is not required in this first version.",
    )

    render_workspace_panel(
        uploaded_count=len(uploaded_files or []),
        field_count=len(fields),
        question_count=len(questions),
        model=model,
    )

    col_run, col_clear = st.columns([1, 1])
    with col_clear:
        if st.button("Clear current results"):
            st.session_state.results = []
            st.session_state.errors = []
            st.rerun()

    with col_run:
        run = st.button("Run extraction", type="primary")

    if run:
        if not api_key:
            st.warning("Please enter an API key before running extraction.")
        elif not uploaded_files:
            st.warning("Please upload at least one PDF.")
        elif not fields and not questions:
            st.warning("Please enter at least one structured field or research question.")
        else:
            st.session_state.results = []
            st.session_state.errors = []
            progress = st.progress(0)
            status = st.empty()

            for index, uploaded_file in enumerate(uploaded_files, start=1):
                status.info(f"Processing {uploaded_file.name} ({index}/{len(uploaded_files)})")
                try:
                    result = extract_from_pdf(
                        uploaded_file=uploaded_file,
                        api_key=api_key,
                        base_url=base_url,
                        model=model,
                        fields=fields,
                        questions=questions,
                        prompt_template=prompt_template,
                    )
                    st.session_state.results.append(result)
                except Exception as exc:
                    st.session_state.errors.append(
                        {"file": uploaded_file.name, "message": str(exc)}
                    )
                progress.progress(index / len(uploaded_files))

            status.success("Batch finished.")

    render_results()


if __name__ == "__main__":
    main()
