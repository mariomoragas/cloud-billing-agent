from __future__ import annotations

from pathlib import Path
from typing import Iterable

import pandas as pd

from app.schemas import BillingRecord

REQUIRED_COLUMNS = {
    "service_name_original",
    "usage_quantity",
    "cost",
}

COLUMN_ALIASES = {
    "cloud": ["cloud", "provider", "source_cloud"],
    "service_name_original": [
        "service_name_original",
        "service",
        "service_name",
        "product_name",
        "meter_category",
    ],
    "sku": ["sku", "sku_name", "usage_type", "meter_name"],
    "region": ["region", "location", "resource_location"],
    "usage_quantity": ["usage_quantity", "quantity", "usage", "consumed_quantity"],
    "usage_unit": ["usage_unit", "unit", "pricing_unit"],
    "cost": ["cost", "amount", "pretax_cost", "unblended_cost"],
    "currency": ["currency", "currency_code", "billing_currency"],
    "period": ["period", "billing_period", "invoice_month", "date"],
}


def _normalize_headers(columns: Iterable[str]) -> dict[str, str]:
    normalized = {}
    for original in columns:
        key = original.strip().lower().replace(" ", "_")
        normalized[key] = original
    return normalized


def _resolve_columns(df: pd.DataFrame) -> dict[str, str]:
    normalized_headers = _normalize_headers(df.columns)
    resolved: dict[str, str] = {}

    for target, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            if alias in normalized_headers:
                resolved[target] = normalized_headers[alias]
                break

    missing = REQUIRED_COLUMNS.difference(resolved)
    if missing:
        missing_display = ", ".join(sorted(missing))
        raise ValueError(
            "CSV sem colunas obrigatorias para o modelo normalizado: "
            f"{missing_display}"
        )

    return resolved


def load_and_normalize_csv(csv_path: Path, default_cloud: str) -> pd.DataFrame:
    raw_df = pd.read_csv(csv_path)
    column_map = _resolve_columns(raw_df)

    normalized = pd.DataFrame(
        {
            "cloud": raw_df[column_map["cloud"]]
            if "cloud" in column_map
            else default_cloud,
            "service_name_original": raw_df[column_map["service_name_original"]],
            "sku": raw_df[column_map["sku"]] if "sku" in column_map else "",
            "region": raw_df[column_map["region"]] if "region" in column_map else "",
            "usage_quantity": raw_df[column_map["usage_quantity"]],
            "usage_unit": (
                raw_df[column_map["usage_unit"]] if "usage_unit" in column_map else ""
            ),
            "cost": raw_df[column_map["cost"]],
            "currency": (
                raw_df[column_map["currency"]] if "currency" in column_map else "USD"
            ),
            "period": raw_df[column_map["period"]] if "period" in column_map else "",
        }
    )

    normalized["cloud"] = normalized["cloud"].fillna(default_cloud).astype(str).str.lower()
    normalized["service_name_original"] = (
        normalized["service_name_original"].fillna("").astype(str).str.strip()
    )
    normalized["sku"] = normalized["sku"].fillna("").astype(str).str.strip()
    normalized["region"] = normalized["region"].fillna("").astype(str).str.strip()
    normalized["usage_unit"] = normalized["usage_unit"].fillna("").astype(str).str.strip()
    normalized["currency"] = normalized["currency"].fillna("USD").astype(str).str.strip()
    normalized["period"] = normalized["period"].fillna("").astype(str).str.strip()
    normalized["usage_quantity"] = pd.to_numeric(
        normalized["usage_quantity"], errors="coerce"
    ).fillna(0.0)
    normalized["cost"] = pd.to_numeric(normalized["cost"], errors="coerce").fillna(0.0)

    return normalized


def dataframe_to_records(df: pd.DataFrame) -> list[BillingRecord]:
    return [
        BillingRecord(
            cloud=row["cloud"],
            service_name_original=row["service_name_original"],
            sku=row["sku"],
            region=row["region"],
            usage_quantity=float(row["usage_quantity"]),
            usage_unit=row["usage_unit"],
            cost=float(row["cost"]),
            currency=row["currency"],
            period=row["period"],
        )
        for _, row in df.iterrows()
    ]
