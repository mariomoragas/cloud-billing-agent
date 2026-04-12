from __future__ import annotations

from pathlib import Path

import pandas as pd


def load_mapping_table(mapping_csv: Path) -> pd.DataFrame:
    mapping_df = pd.read_csv(mapping_csv).fillna("")
    required = {
        "source_cloud",
        "source_service",
        "source_product_code",
        "source_pattern",
        "oci_service",
        "rule_type",
        "confidence",
    }
    missing = required.difference(mapping_df.columns)
    if missing:
        raise ValueError(
            "Arquivo de mapeamento OCI sem colunas obrigatorias: "
            + ", ".join(sorted(missing))
        )
    mapping_df["source_cloud"] = mapping_df["source_cloud"].astype(str).str.lower()
    return mapping_df


def build_oci_mapping(
    service_summary: pd.DataFrame, mapping_df: pd.DataFrame
) -> pd.DataFrame:
    rows: list[dict[str, object]] = []

    for _, service_row in service_summary.iterrows():
        match = _find_best_match(
            cloud=service_row["cloud"],
            service_name=service_row["service_name_original"],
            product_code=service_row.get("primary_product_code", ""),
            mapping_df=mapping_df,
        )

        if match is None:
            rows.append(
                {
                    "cloud": service_row["cloud"],
                    "service_name_original": service_row["service_name_original"],
                    "primary_product_code": service_row.get("primary_product_code", ""),
                    "total_usage_quantity": service_row["total_usage_quantity"],
                    "primary_unit": service_row["primary_unit"],
                    "total_cost": service_row["total_cost"],
                    "primary_currency": service_row["primary_currency"],
                    "oci_service": "REVIEW_REQUIRED",
                    "rule_type": "unmapped",
                    "confidence": 0.0,
                    "notes": "Servico sem regra de equivalencia. Revisao manual necessaria.",
                }
            )
            continue

        rows.append(
            {
                "cloud": service_row["cloud"],
                "service_name_original": service_row["service_name_original"],
                "primary_product_code": service_row.get("primary_product_code", ""),
                "total_usage_quantity": service_row["total_usage_quantity"],
                "primary_unit": service_row["primary_unit"],
                "total_cost": service_row["total_cost"],
                "primary_currency": service_row["primary_currency"],
                "oci_service": match["oci_service"],
                "rule_type": match["rule_type"],
                "confidence": float(match["confidence"]),
                "notes": _build_note(match),
            }
        )

    return pd.DataFrame(rows).sort_values(
        by=["confidence", "total_cost"], ascending=[True, False]
    )


def _find_best_match(
    cloud: str,
    service_name: str,
    product_code: str,
    mapping_df: pd.DataFrame,
) -> pd.Series | None:
    candidates = mapping_df[mapping_df["source_cloud"] == str(cloud).lower()].copy()
    if candidates.empty:
        return None

    if str(product_code).strip():
        exact_code = candidates[
            candidates["source_product_code"].astype(str).str.lower()
            == str(product_code).lower()
        ]
        if not exact_code.empty:
            return exact_code.sort_values("confidence", ascending=False).iloc[0]

    exact = candidates[
        candidates["source_service"].astype(str).str.lower() == str(service_name).lower()
    ]
    if not exact.empty:
        return exact.sort_values("confidence", ascending=False).iloc[0]

    for _, row in candidates.sort_values("confidence", ascending=False).iterrows():
        pattern = str(row["source_pattern"]).strip()
        if not pattern:
            continue
        if pattern == ".*":
            continue
        if pd.Series([service_name]).str.contains(pattern, case=False, regex=True).iloc[0]:
            return row

    wildcard = candidates[candidates["source_pattern"].astype(str).str.strip() == ".*"]
    if not wildcard.empty:
        same_service = wildcard[
            wildcard["source_service"].astype(str).str.lower()
            == str(service_name).lower()
        ]
        if not same_service.empty:
            return same_service.sort_values("confidence", ascending=False).iloc[0]
    return None


def _build_note(match: pd.Series) -> str:
    if str(match["rule_type"]).lower() == "exact":
        return "Mapeamento realizado por regra exata."
    return f"Mapeamento realizado por padrao: {match['source_pattern']}"
