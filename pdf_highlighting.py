import csv
import io
import re
import zipfile
from pathlib import Path
from typing import Any

from config import HIGHLIGHT_COLORS


PAGE_RE = re.compile(r"\b(?:page|p\.?|pp\.?)\s*[:.]?\s*(\d{1,4})\b", re.IGNORECASE)
INVALID_FILENAME_CHARS_RE = re.compile(r'[\\/:*?"<>|]+')


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def safe_zip_filename(filename: str, suffix: str) -> str:
    path = Path(filename)
    stem = INVALID_FILENAME_CHARS_RE.sub("_", path.stem or "article")
    extension = path.suffix if path.suffix.casefold() == ".pdf" else ".pdf"
    return f"{stem}{suffix}{extension}"


def unique_name(filename: str, used_names: set[str]) -> str:
    if filename not in used_names:
        used_names.add(filename)
        return filename

    path = Path(filename)
    counter = 2
    while True:
        candidate = f"{path.stem}_{counter}{path.suffix}"
        if candidate not in used_names:
            used_names.add(candidate)
            return candidate
        counter += 1


def parse_page_index(source_location: str, page_count: int) -> int | None:
    match = PAGE_RE.search(source_location or "")
    if not match:
        return None
    page_number = int(match.group(1))
    if 1 <= page_number <= page_count:
        return page_number - 1
    return None


def excerpt_variants(text: str) -> list[tuple[str, str]]:
    stripped = clean_text(text)
    variants: list[tuple[str, str]] = []

    def add(method: str, value: str) -> None:
        value = clean_text(value)
        if value and value.casefold() not in {existing.casefold() for _, existing in variants}:
            variants.append((method, value))

    add("exact", stripped)
    add("collapsed whitespace", re.sub(r"\s+", " ", stripped))
    add("dehyphenated line breaks", re.sub(r"(\w)-\s+(\w)", r"\1\2", stripped))
    add(
        "dehyphenated + collapsed whitespace",
        re.sub(r"\s+", " ", re.sub(r"(\w)-\s+(\w)", r"\1\2", stripped)),
    )
    return variants


def color_for_question(question_index: int) -> dict[str, Any]:
    return HIGHLIGHT_COLORS[(question_index - 1) % len(HIGHLIGHT_COLORS)]


def apply_highlight(page: Any, rectangles: list[Any], color: tuple[float, float, float]) -> int:
    count = 0
    for rectangle in rectangles:
        annotation = page.add_highlight_annot(rectangle)
        annotation.set_colors(stroke=color)
        annotation.update()
        count += 1
    return count


def search_excerpt(
    document: Any,
    text: str,
    source_location: str,
) -> tuple[int, str, str]:
    target_page_index = parse_page_index(source_location, len(document))
    if target_page_index is None:
        page_indexes = range(len(document))
        page_note = "searched all pages"
    else:
        page_indexes = [target_page_index]
        page_note = f"searched page {target_page_index + 1}"

    for method, variant in excerpt_variants(text):
        for page_index in page_indexes:
            page = document[page_index]
            rectangles = page.search_for(variant)
            if rectangles:
                return page_index, method, page_note
    return -1, "", page_note


def highlight_extraction_result(
    pdf_bytes: bytes,
    result: dict[str, Any],
) -> tuple[bytes, list[dict[str, str]]]:
    try:
        import fitz
    except ImportError as exc:
        raise RuntimeError(
            "PyMuPDF is not installed. Please install requirements.txt again before using PDF highlighting."
        ) from exc

    source_file = clean_text(result.get("source_file")) or "uploaded.pdf"
    report_rows: list[dict[str, str]] = []
    document = fitz.open(stream=pdf_bytes, filetype="pdf")

    try:
        for question_index, evidence in enumerate(
            result.get("research_question_evidence", []),
            start=1,
        ):
            question_label = f"RQ{question_index}"
            color = color_for_question(question_index)
            excerpts = evidence.get("excerpts", []) or []
            for excerpt_index, excerpt in enumerate(excerpts, start=1):
                excerpt_text = clean_text(excerpt.get("text"))
                source_location = clean_text(excerpt.get("source_location"))
                row = {
                    "file": source_file,
                    "research_question": question_label,
                    "excerpt_index": str(excerpt_index),
                    "source_location": source_location,
                    "status": "unmatched",
                    "highlight_count": "0",
                    "match_method": "",
                    "note": "",
                    "excerpt_preview": excerpt_text[:500],
                }

                if not excerpt_text or excerpt_text.casefold() in {"not found", "not relevant to this article"}:
                    row["note"] = "No highlightable excerpt text."
                    report_rows.append(row)
                    continue

                page_index, method, page_note = search_excerpt(document, excerpt_text, source_location)
                row["note"] = page_note
                if page_index < 0:
                    row["note"] = f"{page_note}; no exact text match after light cleanup."
                    report_rows.append(row)
                    continue

                rectangles = []
                for _, variant in excerpt_variants(excerpt_text):
                    rectangles = document[page_index].search_for(variant)
                    if rectangles:
                        break
                highlight_count = apply_highlight(
                    document[page_index],
                    rectangles,
                    color["rgb"],
                )
                row["status"] = "highlighted"
                row["highlight_count"] = str(highlight_count)
                row["match_method"] = method
                report_rows.append(row)

        output = io.BytesIO()
        document.save(output, garbage=4, deflate=True)
        return output.getvalue(), report_rows
    finally:
        document.close()


def build_legend_rows(results: list[dict[str, Any]]) -> list[dict[str, str]]:
    max_question_count = max(
        (len(result.get("research_question_evidence", [])) for result in results),
        default=0,
    )
    rows: list[dict[str, str]] = []
    for question_index in range(1, max_question_count + 1):
        color = color_for_question(question_index)
        question_text = ""
        for result in results:
            questions = result.get("requested_questions", [])
            if len(questions) >= question_index:
                question_text = clean_text(questions[question_index - 1])
                break
        rows.append(
            {
                "research_question": f"RQ{question_index}",
                "color_name": color["name"],
                "color_hex": color["hex"],
                "question_text": question_text,
            }
        )
    return rows


def rows_to_csv_bytes(rows: list[dict[str, str]], default_headers: list[str]) -> bytes:
    output = io.StringIO()
    headers = default_headers[:]
    for row in rows:
        for key in row:
            if key not in headers:
                headers.append(key)
    writer = csv.DictWriter(output, fieldnames=headers)
    writer.writeheader()
    writer.writerows(rows)
    return output.getvalue().encode("utf-8-sig")


def build_highlight_zip(
    results: list[dict[str, Any]],
    uploaded_pdf_bytes: dict[str, bytes],
) -> tuple[bytes, list[dict[str, str]]]:
    zip_buffer = io.BytesIO()
    all_report_rows: list[dict[str, str]] = []
    used_names: set[str] = set()

    with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for result in results:
            source_file = clean_text(result.get("source_file"))
            pdf_bytes = uploaded_pdf_bytes.get(source_file)
            if not pdf_bytes:
                all_report_rows.append(
                    {
                        "file": source_file,
                        "research_question": "",
                        "excerpt_index": "",
                        "source_location": "",
                        "status": "file missing",
                        "highlight_count": "0",
                        "match_method": "",
                        "note": "Original uploaded PDF bytes are no longer available. Re-upload the PDFs and run extraction again.",
                        "excerpt_preview": "",
                    }
                )
                continue

            try:
                highlighted_pdf, report_rows = highlight_extraction_result(pdf_bytes, result)
                all_report_rows.extend(report_rows)
                pdf_name = unique_name(safe_zip_filename(source_file, "_highlighted"), used_names)
                archive.writestr(pdf_name, highlighted_pdf)
            except Exception as exc:
                all_report_rows.append(
                    {
                        "file": source_file,
                        "research_question": "",
                        "excerpt_index": "",
                        "source_location": "",
                        "status": "failed",
                        "highlight_count": "0",
                        "match_method": "",
                        "note": str(exc),
                        "excerpt_preview": "",
                    }
                )

        legend_rows = build_legend_rows(results)
        archive.writestr(
            "highlight_legend.csv",
            rows_to_csv_bytes(
                legend_rows,
                ["research_question", "color_name", "color_hex", "question_text"],
            ),
        )
        archive.writestr(
            "highlight_report.csv",
            rows_to_csv_bytes(
                all_report_rows,
                [
                    "file",
                    "research_question",
                    "excerpt_index",
                    "source_location",
                    "status",
                    "highlight_count",
                    "match_method",
                    "note",
                    "excerpt_preview",
                ],
            ),
        )

    return zip_buffer.getvalue(), all_report_rows


def summarize_report(report_rows: list[dict[str, str]]) -> tuple[int, int]:
    total = 0
    highlighted = 0
    for row in report_rows:
        if not row.get("research_question") or not row.get("excerpt_index"):
            continue
        total += 1
        if row.get("status") == "highlighted":
            highlighted += 1
    return highlighted, total
