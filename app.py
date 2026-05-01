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
You are an expert systematic review data extractor with experience in evidence synthesis, qualitative evidence extraction, health communication research, and thematic analysis.

Your task is to read the attached research article PDF and extract information for one article record. Work as a careful research assistant, not as a creative writer.

Core principles:
- Accuracy is more important than fluency.
- Transparency is more important than completeness of prose.
- Do not invent, infer beyond the article, or fill gaps from general knowledge.
- If the article does not clearly provide an answer, write "Not found".
- If the article is unrelated to a research question, write "Not relevant to this article" in the summary and explain why briefly when possible.
- Preserve original wording in evidence excerpts. Do not paraphrase inside excerpt text.
- Keep summaries concise and evidence-based. Summaries must be supported by the extracted excerpts or clearly identifiable article content.

Anti-hallucination rules:
- Use only information found in the attached PDF.
- Do not use outside knowledge, assumptions about the topic, or common patterns from similar papers.
- Do not make up page numbers, sections, participant characteristics, methods, outcomes, or conclusions.
- Do not treat the abstract alone as enough if the full text contains more specific evidence.
- If a requested field is ambiguous, choose the most conservative answer and mark confidence as "low".
- If evidence is indirect, partial, or only implied, say so in low_confidence_reason.
- If the PDF text is unreadable, incomplete, scanned poorly, or appears to omit pages, mark confidence as "low" and add a review warning.

Anti-miss rules:
- Use an exhaustive extraction strategy for research-question evidence.
- Search the whole article, including abstract, introduction/background, methods, results/findings, discussion, tables, figures, boxes, appendices, and limitations if available.
- For each research question, extract every relevant original-text excerpt you can find.
- Err on the side of including more potentially relevant excerpts rather than fewer.
- Include borderline relevant excerpts when they may help later thematic analysis, and explain the relevance_note.
- Do not omit repeated but meaningfully different evidence from different sections or participant groups.
- Do not collapse multiple distinct findings into one excerpt if separate excerpts would be useful for coding.

Source-location rules:
- Include page numbers when visible or inferable from the PDF.
- If page numbers are not visible, use section names such as "Abstract", "Results", "Discussion", "Table 2", or "Figure 1".
- If neither page nor section is clear, write "location unclear".

Confidence rules:
- Use "high" only when the answer is directly supported by clear article text.
- Use "medium" when the answer is supported but incomplete, scattered, or requires minor interpretation.
- Use "low" when the answer is uncertain, indirect, missing, contradictory, or affected by PDF quality.
- Use low_confidence_reason to explain every medium or low confidence answer in plain language.

Output content rules:
- Return one article record only.
- For every requested structured field, provide a value, source_location, confidence, and low_confidence_reason.
- For every requested research question, provide an answer_summary, confidence, low_confidence_reason, and all relevant excerpts.
- Never leave required values blank. Use "Not found" or "Not relevant to this article" where appropriate.
- Follow the required structured output schema exactly.

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
    return (
        template.replace("{structured_fields}", field_list)
        .replace("{research_questions}", question_list)
        .strip()
    )


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
    data = normalize_extraction_result(data, fields, questions)
    data["source_file"] = uploaded_file.name
    data["requested_fields"] = fields
    data["requested_questions"] = questions
    data["prompt_used"] = prompt
    return data


def clean_text(value: Any, default: str = "Not found") -> str:
    if value is None:
        return default
    text = str(value).strip()
    return text if text else default


def normalize_extraction_result(
    data: dict[str, Any],
    fields: list[str],
    questions: list[str],
) -> dict[str, Any]:
    article = data.setdefault("article", {})
    for key in ("title", "authors", "year", "journal"):
        article[key] = clean_text(article.get(key))
    article["overall_confidence"] = clean_text(article.get("overall_confidence"), "low").lower()
    if article["overall_confidence"] not in CONFIDENCE_LEVELS:
        article["overall_confidence"] = "low"
    article["low_confidence_reason"] = clean_text(article.get("low_confidence_reason"))

    returned_fields = {
        clean_text(item.get("name"), "").casefold(): item
        for item in data.get("structured_fields", [])
        if isinstance(item, dict)
    }
    normalized_fields = []
    for field in fields:
        item = returned_fields.get(field.casefold(), {})
        confidence = clean_text(item.get("confidence"), "low").lower()
        normalized_fields.append(
            {
                "name": field,
                "value": clean_text(item.get("value")),
                "source_location": clean_text(item.get("source_location")),
                "confidence": confidence if confidence in CONFIDENCE_LEVELS else "low",
                "low_confidence_reason": clean_text(item.get("low_confidence_reason")),
            }
        )
    data["structured_fields"] = normalized_fields

    returned_evidence = [
        item for item in data.get("research_question_evidence", []) if isinstance(item, dict)
    ]
    normalized_evidence = []
    for index, question in enumerate(questions, start=1):
        item = returned_evidence[index - 1] if index - 1 < len(returned_evidence) else {}
        confidence = clean_text(item.get("confidence"), "low").lower()
        excerpts = []
        for excerpt in item.get("excerpts", []) or []:
            if not isinstance(excerpt, dict):
                continue
            excerpts.append(
                {
                    "text": clean_text(excerpt.get("text")),
                    "source_location": clean_text(excerpt.get("source_location")),
                    "relevance_note": clean_text(excerpt.get("relevance_note")),
                }
            )
        normalized_evidence.append(
            {
                "question": f"RQ{index}: {question}",
                "answer_summary": clean_text(item.get("answer_summary")),
                "confidence": confidence if confidence in CONFIDENCE_LEVELS else "low",
                "low_confidence_reason": clean_text(item.get("low_confidence_reason")),
                "excerpts": excerpts,
            }
        )
    data["research_question_evidence"] = normalized_evidence
    data["review_warnings"] = [
        clean_text(warning, "") for warning in data.get("review_warnings", []) if clean_text(warning, "")
    ]
    return data


def requested_fields_for_result(result: dict[str, Any]) -> list[str]:
    fields = result.get("requested_fields")
    if fields:
        return list(fields)
    return [item.get("name", "") for item in result.get("structured_fields", []) if item.get("name")]


def requested_questions_for_result(result: dict[str, Any]) -> list[str]:
    questions = result.get("requested_questions")
    if questions:
        return list(questions)
    return [
        evidence.get("question", "")
        for evidence in result.get("research_question_evidence", [])
        if evidence.get("question")
    ]


def result_to_flat_row(result: dict[str, Any]) -> dict[str, str]:
    article = result.get("article", {})
    row = {
        "File names": result.get("source_file", ""),
        "title": article.get("title", ""),
        "authors": article.get("authors", ""),
        "year": article.get("year", ""),
        "journal": article.get("journal", ""),
        "overall_confidence": article.get("overall_confidence", ""),
        "review_warnings": "; ".join(result.get("review_warnings", [])),
    }

    fields_by_name = {
        item.get("name", "").casefold(): item for item in result.get("structured_fields", [])
    }
    for field in requested_fields_for_result(result):
        item = fields_by_name.get(field.casefold(), {})
        row[field] = clean_text(item.get("value"))
        row[f"{field} confidence"] = clean_text(item.get("confidence"), "low")

    evidence_items = result.get("research_question_evidence", [])
    for index, _question in enumerate(requested_questions_for_result(result), start=1):
        evidence = evidence_items[index - 1] if index - 1 < len(evidence_items) else {}
        row[f"RQ{index} summary"] = clean_text(evidence.get("answer_summary"))
        row[f"RQ{index} confidence"] = clean_text(evidence.get("confidence"), "low")

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
        excerpts = evidence.get("excerpts", []) or [{"text": "Not found", "source_location": "Not found", "relevance_note": "Not found"}]
        for excerpt in excerpts:
            rows.append(
                {
                    "File names": result.get("source_file", ""),
                    "title": clean_text(article.get("title")),
                    "research question": f"RQ{question_index}",
                    "answer_summary": clean_text(evidence.get("answer_summary")),
                    "confidence": clean_text(evidence.get("confidence"), "low"),
                    "excerpt": clean_text(excerpt.get("text")),
                    "source_location": clean_text(excerpt.get("source_location")),
                    "relevance_note": clean_text(excerpt.get("relevance_note")),
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

    header_by_column = {
        cell.column: str(cell.value or "").casefold()
        for cell in sheet[1]
    }

    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = border
            is_confidence_cell = "confidence" in header_by_column.get(cell.column, "")
            if is_confidence_cell and isinstance(cell.value, str) and confidence_needs_review(cell.value):
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
            --app-bg: #f6f8fb;
            --panel-bg: #ffffff;
            --sidebar-bg: #eef2f7;
            --text-main: #111827;
            --text-muted: #64748b;
            --border: #dfe6ef;
            --border-strong: #c7d2df;
            --accent: #334155;
            --accent-soft: #eef3f8;
            --shadow: 0 18px 46px rgba(15, 23, 42, 0.08);
        }

        html, body, [class*="css"] {
            font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display", "SF Pro Text", Inter, "Segoe UI", sans-serif;
        }

        .stApp {
            background: var(--app-bg);
            color: var(--text-main);
        }

        [data-testid="stSidebar"] {
            background: var(--sidebar-bg);
            border-right: 1px solid var(--border);
        }

        [data-testid="stHeader"],
        header[data-testid="stHeader"],
        [data-testid="stAppViewContainer"] > header,
        .stAppHeader {
            background: var(--app-bg) !important;
            box-shadow: none !important;
            border-bottom: 1px solid var(--border) !important;
        }

        [data-testid="stToolbar"],
        .stAppToolbar {
            background: transparent !important;
        }

        [data-testid="stSidebar"] > div:first-child {
            padding-top: 1.35rem;
            padding-left: 1.15rem;
            padding-right: 1.15rem;
        }

        [data-testid="stSidebar"] h2,
        [data-testid="stSidebar"] h3 {
            color: var(--text-main);
            letter-spacing: 0;
            font-size: 1rem;
            font-weight: 680;
        }

        .main .block-container {
            max-width: 1120px;
            padding-top: 2.15rem;
            padding-bottom: 2.8rem;
        }

        h1 {
            color: var(--text-main);
            font-weight: 650;
            font-size: clamp(2.15rem, 4vw, 3.35rem);
            letter-spacing: 0;
            line-height: 1.04;
            margin-bottom: 0.45rem;
        }

        div[data-testid="stCaptionContainer"] {
            color: var(--text-muted);
        }

        label, .stMarkdown p {
            color: var(--text-main);
        }

        .stTextInput input,
        .stTextArea textarea {
            border-radius: 11px;
            border: 1px solid var(--border-strong);
            background: #ffffff;
            color: var(--text-main);
            box-shadow: none;
        }

        .stTextInput input:focus,
        .stTextArea textarea:focus {
            border-color: var(--accent);
            box-shadow: 0 0 0 1px rgba(89, 104, 92, 0.24);
        }

        .stButton button {
            border-radius: 999px;
            border: 1px solid var(--border);
            background: #ffffff;
            color: var(--text-main);
            font-weight: 620;
            min-height: 2.6rem;
            padding-left: 1.05rem;
            padding-right: 1.05rem;
        }

        .stButton button:hover {
            border-color: var(--border-strong);
            color: var(--text-main);
            background: #f8fafc;
        }

        [data-testid="stSidebar"] .stButton button {
            min-height: 2.35rem;
            padding-left: 0.85rem;
            padding-right: 0.85rem;
            width: 100%;
        }

        .stButton button[kind="primary"] {
            background: var(--text-main);
            border-color: var(--text-main);
            color: #ffffff;
        }

        .stButton button[kind="primary"]:hover {
            background: #343431;
            border-color: #343431;
            color: #ffffff;
        }

        div[data-testid="stFileUploader"] section {
            border-radius: 15px;
            border-color: var(--border);
            background: #f8fafc;
            padding: 1.15rem;
        }

        div[data-testid="stFileUploader"] section:hover {
            border-color: var(--border-strong);
            background: #f1f5f9;
        }

        div[data-testid="stDataFrame"] {
            border: 1px solid var(--border);
            border-radius: 14px;
            overflow: hidden;
        }

        .rq-label {
            font-weight: 640;
            color: var(--text-muted);
            padding-top: 0.45rem;
        }

        .rq-control-button {
            margin-top: 0.15rem;
        }

        .section-note {
            color: var(--text-muted);
            font-size: 0.86rem;
            margin-top: -0.35rem;
            margin-bottom: 0.75rem;
        }

        .app-hero {
            margin-bottom: 1.25rem;
            max-width: 860px;
        }

        .app-kicker {
            display: inline-flex;
            align-items: center;
            gap: 0.5rem;
            color: var(--text-muted);
            font-size: 0.82rem;
            font-weight: 560;
            margin-bottom: 0.75rem;
            letter-spacing: 0.01em;
        }

        .app-kicker svg,
        .panel-title svg,
        .upload-title svg,
        .metric-card svg {
            width: 17px;
            height: 17px;
            stroke: currentColor;
            stroke-width: 1.75;
            fill: none;
            stroke-linecap: round;
            stroke-linejoin: round;
        }

        .workspace-panel {
            border: 1px solid var(--border);
            background: var(--panel-bg);
            border-radius: 18px;
            padding: 1.05rem;
            box-shadow: var(--shadow);
            min-height: 100%;
        }

        .panel-title {
            display: flex;
            align-items: center;
            gap: 0.52rem;
            font-weight: 660;
            color: var(--text-main);
            margin-bottom: 0.25rem;
        }

        .panel-subtitle {
            color: var(--text-muted);
            font-size: 0.9rem;
            line-height: 1.45;
            margin-bottom: 0.95rem;
        }

        .metric-grid {
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 0.62rem;
            margin: 0.85rem 0 0;
        }

        .metric-card {
            border: 1px solid var(--border);
            border-radius: 14px;
            padding: 0.78rem;
            background: #f8fafc;
            min-height: 78px;
        }

        .metric-card svg {
            color: var(--accent);
            margin-bottom: 0.34rem;
        }

        .metric-label {
            color: var(--text-muted);
            font-size: 0.76rem;
            line-height: 1.2;
        }

        .metric-value {
            color: var(--text-main);
            font-weight: 650;
            font-size: 1.02rem;
            margin-top: 0.16rem;
            overflow-wrap: anywhere;
        }

        .upload-shell {
            border: 1px solid var(--border);
            background: var(--panel-bg);
            border-radius: 18px;
            padding: 1.05rem 1.05rem 0.95rem;
            box-shadow: var(--shadow);
            margin-bottom: 1rem;
        }

        .upload-title {
            display: flex;
            align-items: center;
            gap: 0.52rem;
            font-weight: 660;
            color: var(--text-main);
            margin-bottom: 0.22rem;
        }

        .upload-subtitle {
            color: var(--text-muted);
            font-size: 0.9rem;
            line-height: 1.45;
            margin-bottom: 0.85rem;
        }

        .action-row-note {
            color: var(--text-muted);
            font-size: 0.84rem;
            line-height: 1.45;
            padding-top: 0.55rem;
        }

        .results-spacer {
            margin-top: 1.15rem;
        }

        @media (max-width: 900px) {
            .metric-grid {
                grid-template-columns: repeat(2, minmax(0, 1fr));
            }

            .main .block-container {
                padding-left: 1rem;
                padding-right: 1rem;
                padding-top: 1.4rem;
            }

            .upload-shell,
            .workspace-panel {
                padding: 0.9rem;
                border-radius: 15px;
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
            <div style="color:#64748b; max-width:760px; line-height:1.55; font-size:1rem;">
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
            <div class="panel-subtitle">A compact check of the current batch setup before you run extraction.</div>
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
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_upload_intro() -> None:
    st.markdown(
        f"""
        <div class="upload-shell">
            <div class="upload-title">{svg_icon("file")} Source PDFs</div>
            <div class="upload-subtitle">Upload the articles for this extraction run. Files stay local until each PDF is sent to the configured model provider.</div>
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
        st.markdown('<div class="rq-control-button"></div>', unsafe_allow_html=True)
        st.button("＋", key="add_rq", help="Add a research question", on_click=add_research_question)

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
                "×",
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
        value="Study design\nPopulation / sample\nIntervention or exposure\nMain findings\nLimitations",
        height=150,
        key="structured_fields_text",
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

    upload_col, summary_col = st.columns([0.64, 0.36], gap="medium", vertical_alignment="top")
    with upload_col:
        render_upload_intro()
        uploaded_files = st.file_uploader(
            "Upload PDF articles",
            type=["pdf"],
            accept_multiple_files=True,
            label_visibility="collapsed",
            help="Select multiple PDFs. Browser folder upload is not required in this first version.",
        )

    with summary_col:
        render_workspace_panel(
            uploaded_count=len(uploaded_files or []),
            field_count=len(fields),
            question_count=len(questions),
            model=model,
        )

    col_run, col_clear, col_note = st.columns([0.22, 0.24, 0.54], vertical_alignment="center")

    with col_run:
        run = st.button("Run extraction", type="primary")
    with col_clear:
        if st.button("Clear results"):
            st.session_state.results = []
            st.session_state.errors = []
            st.rerun()
    with col_note:
        st.markdown(
            '<div class="action-row-note">Results export to a stable Excel workbook with article summary, evidence excerpts, and the exact prompt used.</div>',
            unsafe_allow_html=True,
        )

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

    st.markdown('<div class="results-spacer"></div>', unsafe_allow_html=True)
    render_results()


if __name__ == "__main__":
    main()
