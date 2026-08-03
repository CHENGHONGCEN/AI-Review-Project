import unittest
from unittest.mock import patch
from types import ModuleType, SimpleNamespace
import json
import sys

import app
import screening


class SessionState(dict):
    def __getattr__(self, name):
        try:
            return self[name]
        except KeyError as exc:
            raise AttributeError(name) from exc

    def __setattr__(self, name, value):
        self[name] = value


def record(record_id: str) -> dict:
    return {
        "record_id": record_id,
        "title": f"Title {record_id}",
        "abstract": "Abstract",
        "ai_decision": "unsure",
        "ai_suggested_inclusion": True,
        "ai_suggested_exclusion": False,
        "screening_results": {},
    }


def decision(record_id: str, value: str, review: bool = False) -> dict:
    return {
        "record_id": record_id,
        "ai_decision": value,
        "suggested_action": "human review" if value == "unsure" else value,
        "ai_suggested_inclusion": value != "exclude",
        "matched_criteria": "1",
        "matched_exclusion_criteria": "1" if value == "exclude" else "",
        "exclusion_reason": "",
        "reason": f"{value} reason",
        "evidence": "evidence",
        "needs_human_review": review,
    }


def apply(records: list[dict], provider: str, values: list[str]) -> None:
    screening.apply_provider_decisions(
        records=records,
        response_data={
            "records": [
                decision(item["record_id"], value)
                for item, value in zip(records, values)
            ]
        },
        provider=provider,
        model=f"{provider}-model",
        inclusion_criteria_count=2,
        exclusion_criteria_count=2,
    )


class ScreeningComparisonTests(unittest.TestCase):
    def test_all_nine_decision_combinations(self):
        values = ["include", "exclude", "unsure"]
        records = [
            record(f"C{index:02d}")
            for index in range(1, 10)
        ]
        openai_values = [left for left in values for _right in values]
        gemini_values = [right for _left in values for right in values]
        apply(records, "openai", openai_values)
        apply(records, "gemini", gemini_values)

        comparisons = [
            screening.comparison_for_record(item)
            for item in records
        ]
        self.assertEqual(
            [item["status"] for item in comparisons],
            [
                "Agreement",
                "Disagreement",
                "Disagreement",
                "Disagreement",
                "Agreement",
                "Disagreement",
                "Disagreement",
                "Disagreement",
                "Agreement",
            ],
        )
        self.assertFalse(comparisons[0]["needs_human_review"])
        self.assertFalse(comparisons[4]["needs_human_review"])
        self.assertTrue(comparisons[8]["needs_human_review"])
        self.assertTrue(all(item["needs_human_review"] for item in comparisons[1:4]))

    def test_missing_provider_is_incomplete(self):
        records = [record("C01")]
        apply(records, "openai", ["include"])
        comparison = screening.comparison_for_record(records[0])
        self.assertEqual(comparison["status"], "Incomplete")
        self.assertTrue(comparison["needs_human_review"])
        self.assertEqual(
            screening.missing_provider_records(records, "gemini"),
            records,
        )

    def test_invalid_and_missing_record_ids_are_not_applied(self):
        records = [record("C01"), record("C02")]
        with self.assertRaisesRegex(ValueError, "valid result"):
            screening.apply_provider_decisions(
                records=records,
                response_data={"records": [decision("UNKNOWN", "include")]},
                provider="gemini",
                model="gemini-model",
                inclusion_criteria_count=1,
                exclusion_criteria_count=1,
            )
        self.assertEqual(
            screening.missing_provider_records(records, "gemini"),
            records,
        )

    def test_v1_record_migrates_to_openai_namespace(self):
        legacy = record("C01")
        legacy.update(
            {
                "ai_decision": "exclude",
                "ai_reason": "Not eligible",
                "ai_suggested_inclusion": False,
                "ai_suggested_exclusion": True,
            }
        )
        migrated = screening.migrate_v1_record(legacy, marked=True)
        result = screening.provider_result(migrated, "openai")
        self.assertIsNotNone(result)
        self.assertEqual(result["decision"], "exclude")
        self.assertEqual(result["reason"], "Not eligible")


class ProviderClientTests(unittest.TestCase):
    def test_openai_structured_response_is_saved_to_openai_namespace(self):
        records = [record("C01")]
        response = SimpleNamespace(
            output_text=json.dumps(
                {"records": [decision("C01", "include")]}
            )
        )
        fake_client = SimpleNamespace(
            responses=SimpleNamespace(
                create=lambda **_kwargs: response,
            )
        )
        with patch.object(app, "OpenAI", return_value=fake_client):
            app.mark_citation_inclusions(
                records=records,
                inclusion_criteria=["Eligible"],
                exclusion_criteria=["Ineligible"],
                api_key="secret",
                base_url="https://api.openai.com/v1",
                model="openai-model",
                prompt_template=app.DEFAULT_INCLUSION_PROMPT_TEMPLATE,
            )
        self.assertEqual(
            screening.provider_result(records[0], "openai")["decision"],
            "include",
        )
        self.assertEqual(records[0]["ai_decision"], "include")

    def test_gemini_structured_response_is_saved_to_gemini_namespace(self):
        records = [record("C01")]
        captured = {}

        class FakeModels:
            def generate_content(self, **kwargs):
                captured.update(kwargs)
                return SimpleNamespace(
                    parsed={"records": [decision("C01", "exclude")]},
                    text="",
                )

        class FakeGenai:
            @staticmethod
            def Client(**kwargs):
                captured["client"] = kwargs
                return SimpleNamespace(models=FakeModels())

        fake_google = ModuleType("google")
        fake_google.genai = FakeGenai
        with patch.dict(sys.modules, {"google": fake_google}):
            app.mark_citation_inclusions_gemini(
                records=records,
                inclusion_criteria=["Eligible"],
                exclusion_criteria=["Ineligible"],
                api_key="secret",
                model="gemini-model",
                prompt_template=app.DEFAULT_INCLUSION_PROMPT_TEMPLATE,
            )

        self.assertEqual(
            screening.provider_result(records[0], "gemini")["decision"],
            "exclude",
        )
        self.assertEqual(captured["client"]["http_options"], {"api_version": "v1"})
        self.assertEqual(
            captured["config"]["response_mime_type"],
            "application/json",
        )
        self.assertIn("response_json_schema", captured["config"])


class ScreeningRunTests(unittest.TestCase):
    def setUp(self):
        self.session_state = SessionState(
            citation_ai_prompts_used={},
            citation_screening_run={},
            citation_ai_prompt_used="",
            citation_export_timestamp="",
        )
        self.fake_streamlit = type(
            "FakeStreamlit",
            (),
            {"session_state": self.session_state},
        )()

    @staticmethod
    def marker(calls, fail_provider=None):
        def fake_mark(**kwargs):
            provider = kwargs["provider"]
            provider_records = kwargs["records"]
            calls.append(
                (provider, [item["record_id"] for item in provider_records])
            )
            if provider == fail_provider:
                raise RuntimeError(f"{provider} unavailable")
            apply(
                provider_records,
                provider,
                ["include"] * len(provider_records),
            )
            checkpoint = kwargs.get("checkpoint")
            if checkpoint is not None:
                checkpoint(provider_records)
            return provider_records, "prompt"

        return fake_mark

    def run_models(self, records, retry=False):
        return app.run_screening_models(
            records=records,
            inclusion_criteria=["Eligible"],
            exclusion_criteria=["Ineligible"],
            openai_api_key="openai-secret",
            base_url="https://api.openai.com/v1",
            openai_model="openai-model",
            prompt_template=app.DEFAULT_INCLUSION_PROMPT_TEMPLATE,
            dual_model_enabled=True,
            gemini_api_key="gemini-secret",
            gemini_model="gemini-model",
            export_timestamp="20260731_1200",
            retry_incomplete=retry,
        )

    def test_retry_calls_only_missing_provider_records(self):
        records = [record("C01"), record("C02")]
        calls = []
        with (
            patch.object(app, "st", self.fake_streamlit),
            patch.object(
                app,
                "mark_citation_inclusions_batched",
                side_effect=self.marker(calls),
            ),
        ):
            self.run_models(records)
            records[1]["screening_results"].pop("gemini")
            calls.clear()
            self.run_models(records, retry=True)

        self.assertEqual(calls, [("gemini", ["C02"])])
        self.assertNotIn(
            "openai-secret",
            str(self.session_state["citation_screening_run"]),
        )
        self.assertNotIn(
            "gemini-secret",
            str(self.session_state["citation_screening_run"]),
        )

    def test_one_provider_failure_does_not_block_the_other(self):
        records = [record("C01"), record("C02")]
        calls = []
        with (
            patch.object(app, "st", self.fake_streamlit),
            patch.object(
                app,
                "mark_citation_inclusions_batched",
                side_effect=self.marker(calls, fail_provider="openai"),
            ),
        ):
            _records, _prompts, metadata, errors = self.run_models(records)

        self.assertEqual(
            calls,
            [
                ("openai", ["C01", "C02"]),
                ("gemini", ["C01", "C02"]),
            ],
        )
        self.assertEqual(metadata["providers"]["openai"]["status"], "failed")
        self.assertEqual(metadata["providers"]["gemini"]["status"], "complete")
        self.assertTrue(errors)
        self.assertEqual(
            screening.missing_provider_records(records, "gemini"),
            [],
        )

    def test_schema_v2_backup_round_trip_and_no_api_keys(self):
        records = [record("C01")]
        apply(records, "openai", ["include"])
        apply(records, "gemini", ["exclude"])
        self.session_state.update(
            {
                "citation_records": records,
                "citation_duplicate_log": [],
                "citation_import_log": [],
                "citation_errors": [],
                "citation_ai_prompt_used": "prompt",
                "citation_ai_prompts_used": {
                    "openai": "prompt",
                    "gemini": "prompt",
                },
                "citation_screening_run": {
                    "dual_model_enabled": True,
                    "providers": {
                        "openai": {"model": "openai-model"},
                        "gemini": {"model": "gemini-model"},
                    },
                },
                "citation_imported_count": 1,
                "citation_export_timestamp": "20260731_1200",
                "citation_inclusion_criteria": "Eligible",
                "citation_exclusion_criteria": "Ineligible",
            }
        )
        with patch.object(app, "st", self.fake_streamlit):
            payload = app.citation_state_payload()
        serialized = json.dumps(payload)
        self.assertEqual(payload["schema_version"], 2)
        self.assertNotIn("openai-secret", serialized)
        self.assertNotIn("gemini-secret", serialized)

        restored_state = SessionState()
        restored_streamlit = type(
            "FakeStreamlit",
            (),
            {"session_state": restored_state},
        )()
        with patch.object(app, "st", restored_streamlit):
            app.apply_citation_state(payload, restore_criteria=True)
        self.assertEqual(
            screening.provider_result(
                restored_state["citation_records"][0],
                "gemini",
            )["decision"],
            "exclude",
        )
        self.assertTrue(
            restored_state["citation_screening_run"]["dual_model_enabled"]
        )


if __name__ == "__main__":
    unittest.main()
