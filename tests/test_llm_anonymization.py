import json

from app.llm_report import _anonymize_llm_payload


def test_llm_payload_anonymization_preserves_aggregates_and_redacts_identifiers():
    payload = {
        "source_name": "Billing Account customer@example.com invoice-123.csv",
        "baseline_total_cost": 12345.67,
        "currency": "BRL",
        "billing_account_id": "01E96E-7127A7-65CDF4",
        "project_id": "certisign-production-project",
        "top_services": [
            {
                "service_name_original": "Compute Engine arn:aws:ec2:us-east-1:123456789012:instance/i-123",
                "total_cost": 10000.0,
            }
        ],
        "notes": "Owner customer@example.com at 10.20.30.40; resource projects/private-prod/zones/us-east1-b.",
    }

    anonymized = _anonymize_llm_payload(payload)
    serialized = json.dumps(anonymized, ensure_ascii=False)

    assert anonymized["baseline_total_cost"] == 12345.67
    assert anonymized["currency"] == "BRL"
    assert anonymized["top_services"][0]["total_cost"] == 10000.0
    assert anonymized["anonymization"] == {
        "enabled": True,
        "version": "1",
        "scope": "aggregated_only",
        "raw_rows_sent": False,
    }
    assert anonymized["source_name"].startswith("[SOURCE_NAME_")
    assert anonymized["billing_account_id"].startswith("[BILLING_ACCOUNT_ID_")
    assert anonymized["project_id"].startswith("[PROJECT_ID_")
    for secret in (
        "customer@example.com",
        "01E96E-7127A7-65CDF4",
        "certisign-production-project",
        "123456789012",
        "10.20.30.40",
        "projects/private-prod",
    ):
        assert secret not in serialized


def test_sensitive_aliases_are_stable_without_revealing_the_value():
    payload = {"source_name": "customer-billing.csv"}

    first = _anonymize_llm_payload(payload)
    second = _anonymize_llm_payload(payload)

    assert first["source_name"] == second["source_name"]
    assert "customer-billing.csv" not in first["source_name"]

def test_network_boundary_anonymizes_payload(monkeypatch):
    import app.llm_report as llm_report

    captured = {}

    def fake_request(*, body, api_key, model):
        captured["body"] = body
        return {"executive_summary": "ok"}, "", ""

    monkeypatch.setenv("GEMINI_API_KEY", "test-key")
    monkeypatch.setattr(llm_report, "_request_gemini_once", fake_request)

    report, error = llm_report._request_llm_report(
        {
            "source_name": "confidential-customer.csv",
            "project_id": "private-production-project",
            "baseline_total_cost": 500.0,
        },
        "gemini-test",
    )

    serialized = json.dumps(captured["body"], ensure_ascii=False)
    assert report == {"executive_summary": "ok"}
    assert error == ""
    assert "confidential-customer.csv" not in serialized
    assert "private-production-project" not in serialized
    assert "aggregated_only" in serialized
