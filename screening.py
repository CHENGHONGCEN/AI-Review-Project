from __future__ import annotations

from collections import Counter
from copy import deepcopy
from datetime import datetime
import re
from typing import Any


SCREENING_PROVIDERS = ("openai", "gemini")
VALID_SCREENING_DECISIONS = {"include", "exclude", "unsure"}

LEGACY_OPENAI_FIELDS = {
    "decision": "ai_decision",
    "suggested_action": "ai_suggested_action",
    "suggested_inclusion": "ai_suggested_inclusion",
    "suggested_exclusion": "ai_suggested_exclusion",
    "matched_criteria": "ai_matched_criteria",
    "matched_exclusion_criteria": "ai_matched_exclusion_criteria",
    "exclusion_reason": "ai_exclusion_reason",
    "reason": "ai_reason",
    "evidence": "ai_evidence",
    "needs_human_review": "needs_human_review",
}


def _text(value: Any, default: str = "") -> str:
    if value is None:
        return default
    return str(value).strip()


def _bool(value: Any, default: bool = False) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        normalized = value.strip().casefold()
        if normalized in {"true", "yes", "1"}:
            return True
        if normalized in {"false", "no", "0"}:
            return False
    return default


def _criterion_numbers(value: Any, criterion_count: int) -> str:
    if criterion_count <= 0:
        return ""
    numbers: list[str] = []
    for match in re.finditer(r"\b\d{1,3}\b", _text(value)):
        number = int(match.group(0))
        if 1 <= number <= criterion_count and str(number) not in numbers:
            numbers.append(str(number))
    return ",".join(numbers)


def normalize_provider_result(
    decision: dict[str, Any],
    provider: str,
    model: str,
    inclusion_criteria_count: int,
    exclusion_criteria_count: int,
    completed_at: str | None = None,
) -> dict[str, Any] | None:
    raw_decision = _text(
        decision.get("ai_decision", decision.get("decision")),
    ).casefold()
    if raw_decision not in VALID_SCREENING_DECISIONS:
        if "ai_suggested_inclusion" not in decision:
            return None
        raw_decision = (
            "include"
            if _bool(decision.get("ai_suggested_inclusion"), True)
            else "exclude"
        )

    suggested_action = _text(decision.get("suggested_action"))
    if not suggested_action:
        suggested_action = {
            "include": "keep for full-text review",
            "unsure": "human review",
            "exclude": "exclude",
        }[raw_decision]

    return {
        "decision": raw_decision,
        "suggested_action": suggested_action,
        "suggested_inclusion": raw_decision in {"include", "unsure"},
        "suggested_exclusion": raw_decision == "exclude",
        "matched_criteria": _criterion_numbers(
            decision.get("matched_criteria"),
            inclusion_criteria_count,
        ),
        "matched_exclusion_criteria": _criterion_numbers(
            decision.get("matched_exclusion_criteria"),
            exclusion_criteria_count,
        ),
        "exclusion_reason": _text(decision.get("exclusion_reason")),
        "reason": _text(decision.get("reason")),
        "evidence": _text(decision.get("evidence")),
        "needs_human_review": _bool(
            decision.get("needs_human_review"),
            raw_decision == "unsure",
        ),
        "provider": provider,
        "model": _text(model),
        "completed_at": completed_at or datetime.now().isoformat(timespec="seconds"),
    }


def sync_openai_legacy_fields(record: dict[str, Any]) -> None:
    result = provider_result(record, "openai")
    if result is None:
        return
    for result_field, legacy_field in LEGACY_OPENAI_FIELDS.items():
        record[legacy_field] = deepcopy(result.get(result_field))


def apply_provider_decisions(
    records: list[dict[str, Any]],
    response_data: dict[str, Any],
    provider: str,
    model: str,
    inclusion_criteria_count: int,
    exclusion_criteria_count: int,
) -> tuple[list[dict[str, Any]], list[str]]:
    if provider not in SCREENING_PROVIDERS:
        raise ValueError(f"Unsupported screening provider: {provider}")

    returned_records = response_data.get("records", [])
    if not isinstance(returned_records, list):
        raise ValueError("The model response does not contain a records list.")

    decisions: dict[str, dict[str, Any]] = {}
    for item in returned_records:
        if not isinstance(item, dict):
            continue
        record_id = _text(item.get("record_id"))
        if record_id:
            decisions[record_id] = item

    completed_ids: list[str] = []
    for record in records:
        record_id = _text(record.get("record_id"))
        decision = decisions.get(record_id)
        if decision is None:
            continue
        normalized = normalize_provider_result(
            decision=decision,
            provider=provider,
            model=model,
            inclusion_criteria_count=inclusion_criteria_count,
            exclusion_criteria_count=exclusion_criteria_count,
        )
        if normalized is None:
            continue
        results = record.setdefault("screening_results", {})
        if not isinstance(results, dict):
            results = {}
            record["screening_results"] = results
        results[provider] = normalized
        if provider == "openai":
            sync_openai_legacy_fields(record)
        completed_ids.append(record_id)

    if records and not completed_ids:
        raise ValueError("The model response did not contain a valid result for this batch.")
    return records, completed_ids


def provider_result(record: dict[str, Any], provider: str) -> dict[str, Any] | None:
    results = record.get("screening_results")
    if not isinstance(results, dict):
        return None
    result = results.get(provider)
    if not isinstance(result, dict):
        return None
    decision = _text(result.get("decision")).casefold()
    if decision not in VALID_SCREENING_DECISIONS:
        return None
    return result


def missing_provider_records(
    records: list[dict[str, Any]],
    provider: str,
) -> list[dict[str, Any]]:
    return [record for record in records if provider_result(record, provider) is None]


def clear_provider_results(
    records: list[dict[str, Any]],
    providers: tuple[str, ...] = SCREENING_PROVIDERS,
) -> None:
    for record in records:
        results = record.get("screening_results")
        if isinstance(results, dict):
            for provider in providers:
                results.pop(provider, None)
        if "openai" in providers:
            record.update(
                {
                    "ai_decision": "unsure",
                    "ai_suggested_action": "human review",
                    "ai_suggested_inclusion": True,
                    "ai_suggested_exclusion": False,
                    "ai_matched_criteria": "",
                    "ai_matched_exclusion_criteria": "",
                    "ai_exclusion_reason": "",
                    "ai_reason": "",
                    "ai_evidence": "",
                    "needs_human_review": False,
                }
            )


def migrate_v1_record(record: dict[str, Any], marked: bool) -> dict[str, Any]:
    migrated = deepcopy(record)
    migrated.setdefault("screening_results", {})
    if not marked or provider_result(migrated, "openai") is not None:
        return migrated

    legacy_decision = _text(migrated.get("ai_decision")).casefold()
    if legacy_decision not in VALID_SCREENING_DECISIONS:
        return migrated

    result = {
        result_field: deepcopy(migrated.get(legacy_field))
        for result_field, legacy_field in LEGACY_OPENAI_FIELDS.items()
    }
    result["decision"] = legacy_decision
    result["provider"] = "openai"
    result["model"] = ""
    result["completed_at"] = ""
    migrated["screening_results"]["openai"] = result
    sync_openai_legacy_fields(migrated)
    return migrated


def comparison_for_record(record: dict[str, Any]) -> dict[str, Any]:
    openai_result = provider_result(record, "openai")
    gemini_result = provider_result(record, "gemini")
    if openai_result is None or gemini_result is None:
        status = "Incomplete"
    elif openai_result["decision"] == gemini_result["decision"]:
        status = "Agreement"
    else:
        status = "Disagreement"

    needs_human_review = (
        status != "Agreement"
        or any(
            result is not None and result.get("decision") == "unsure"
            for result in (openai_result, gemini_result)
        )
        or any(
            result is not None and _bool(result.get("needs_human_review"))
            for result in (openai_result, gemini_result)
        )
    )
    return {
        "status": status,
        "needs_human_review": needs_human_review,
        "openai_decision": (
            openai_result.get("decision") if openai_result is not None else ""
        ),
        "gemini_decision": (
            gemini_result.get("decision") if gemini_result is not None else ""
        ),
    }


def comparison_summary(records: list[dict[str, Any]]) -> dict[str, Any]:
    statuses = Counter()
    openai_decisions = Counter()
    gemini_decisions = Counter()
    disagreement_combinations = Counter()
    human_review_count = 0

    for record in records:
        comparison = comparison_for_record(record)
        statuses[comparison["status"]] += 1
        if comparison["needs_human_review"]:
            human_review_count += 1
        openai_decision = comparison["openai_decision"]
        gemini_decision = comparison["gemini_decision"]
        if openai_decision:
            openai_decisions[openai_decision] += 1
        if gemini_decision:
            gemini_decisions[gemini_decision] += 1
        if comparison["status"] == "Disagreement":
            disagreement_combinations[
                f"{openai_decision} / {gemini_decision}"
            ] += 1

    completed_pairs = statuses["Agreement"] + statuses["Disagreement"]
    agreement_rate = (
        statuses["Agreement"] / completed_pairs if completed_pairs else None
    )
    return {
        "total": len(records),
        "agreement": statuses["Agreement"],
        "disagreement": statuses["Disagreement"],
        "incomplete": statuses["Incomplete"],
        "human_review": human_review_count,
        "completed_pairs": completed_pairs,
        "agreement_rate": agreement_rate,
        "openai_decisions": dict(openai_decisions),
        "gemini_decisions": dict(gemini_decisions),
        "disagreement_combinations": dict(disagreement_combinations),
    }


def provider_suggests_inclusion(record: dict[str, Any], provider: str) -> bool:
    result = provider_result(record, provider)
    return bool(result and result.get("decision") in {"include", "unsure"})
