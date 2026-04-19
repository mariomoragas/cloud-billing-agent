from __future__ import annotations

import csv
import re
from pathlib import Path

import pandas as pd

from app.report_types import ParserResult


GCP_COST_TABLE_REQUIRED_COLUMNS = {
    "Billing account name",
    "Billing account ID",
    "Project name",
    "Project ID",
    "Project hierarchy",
    "Service description",
    "Service ID",
    "SKU description",
    "SKU ID",
    "Credit type",
    "Cost type",
    "Usage start date",
    "Usage end date",
    "Usage amount",
    "Usage unit",
}


def load_gcp_cost_table_csv(csv_path: Path) -> ParserResult:
    metadata, header_row_index = _read_metadata_and_header_index(csv_path)
    raw_df = pd.read_csv(
        csv_path,
        skiprows=header_row_index,
        dtype=str,
        keep_default_na=False,
        encoding="utf-8-sig",
    )
    _validate_required_columns(raw_df)

    cost_column = _resolve_cost_column(raw_df.columns)
    unrounded_cost_column = _resolve_unrounded_cost_column(raw_df.columns)
    currency = _resolve_currency(metadata, cost_column)

    raw_df = raw_df.copy()
    raw_df["_cost"] = _parse_number_series(raw_df[cost_column])
    raw_df["_unrounded_cost"] = (
        _parse_number_series(raw_df[unrounded_cost_column])
        if unrounded_cost_column
        else raw_df["_cost"]
    )
    raw_df["_usage_amount"] = _parse_number_series(raw_df["Usage amount"])
    raw_df["_service"] = raw_df["Service description"].fillna("").astype(str).str.strip()
    raw_df["_cost_type"] = raw_df["Cost type"].fillna("").astype(str).str.strip()

    total_row_cost = _total_row_cost(raw_df)
    raw_cost_total = total_row_cost if total_row_cost is not None else raw_df["_cost"].sum()

    usage_df = _filter_usage_rows(raw_df)

    normalized = pd.DataFrame(
        {
            "cloud": "gcp",
            "billing_account_name": usage_df["Billing account name"],
            "billing_account_id": usage_df["Billing account ID"],
            "project_name": usage_df["Project name"],
            "project_id": usage_df["Project ID"],
            "project_hierarchy": usage_df["Project hierarchy"],
            "service_name_original": usage_df["Service description"],
            "product_code": usage_df["Service ID"],
            "sku": usage_df["SKU description"],
            "sku_id": usage_df["SKU ID"],
            "region": _infer_region(usage_df),
            "usage_quantity": usage_df["_usage_amount"],
            "usage_unit": usage_df["Usage unit"],
            "cost": usage_df["_cost"],
            "currency": currency,
            "period": _period(usage_df),
            "usage_start_date": usage_df["Usage start date"],
            "usage_end_date": usage_df["Usage end date"],
            "cost_type": usage_df["Cost type"],
            "credit_type": usage_df["Credit type"],
            "unrounded_cost": usage_df["_unrounded_cost"],
            "item_description": usage_df["SKU description"],
        }
    )

    for column in normalized.columns:
        if column not in {"usage_quantity", "cost", "unrounded_cost"}:
            normalized[column] = normalized[column].fillna("").astype(str).str.strip()

    data_quality = _build_data_quality(
        metadata=metadata,
        raw_df=raw_df,
        final_df=normalized,
        raw_cost_total=float(raw_cost_total),
        currency=currency,
    )
    return ParserResult(dataframe=normalized.reset_index(drop=True), data_quality=data_quality)


def _read_metadata_and_header_index(csv_path: Path) -> tuple[dict[str, str], int]:
    metadata: dict[str, str] = {}
    with csv_path.open("r", encoding="utf-8-sig", newline="") as file:
        reader = csv.reader(file)
        for index, row in enumerate(reader):
            if row and row[0].strip() == "Billing account name":
                return metadata, index
            if len(row) >= 2 and row[0].strip():
                metadata[row[0].strip()] = row[1].strip()
    raise ValueError("CSV GCP Cost table sem cabecalho 'Billing account name'.")


def _validate_required_columns(raw_df: pd.DataFrame) -> None:
    missing = GCP_COST_TABLE_REQUIRED_COLUMNS.difference(raw_df.columns)
    if missing:
        raise ValueError(
            "CSV GCP Cost table sem colunas obrigatorias: "
            + ", ".join(sorted(missing))
        )


def _resolve_cost_column(columns: pd.Index) -> str:
    for column in columns:
        if str(column).strip().lower().startswith("cost ("):
            return str(column)
    raise ValueError("CSV GCP Cost table sem coluna de custo final, ex.: 'Cost (R$)'.")


def _resolve_unrounded_cost_column(columns: pd.Index) -> str:
    for column in columns:
        if str(column).strip().lower().startswith("unrounded cost"):
            return str(column)
    return ""


def _resolve_currency(metadata: dict[str, str], cost_column: str) -> str:
    metadata_currency = metadata.get("Currency", "").strip()
    if metadata_currency:
        return metadata_currency

    match = re.search(r"\(([^)]+)\)", cost_column)
    if not match:
        return "USD"

    symbol = match.group(1).strip()
    return {"R$": "BRL", "$": "USD", "EUR": "EUR", "GBP": "GBP"}.get(symbol, symbol)


def _parse_number_series(series: pd.Series) -> pd.Series:
    return series.fillna("").astype(str).map(_parse_number).astype(float)


def _parse_number(value: str) -> float:
    cleaned = str(value or "").strip().replace('"', "").replace("R$", "").replace(" ", "")
    if cleaned in {"", "-"}:
        return 0.0

    has_comma = "," in cleaned
    has_dot = "." in cleaned
    if has_comma and has_dot:
        if cleaned.rfind(".") > cleaned.rfind(","):
            cleaned = cleaned.replace(",", "")
        else:
            cleaned = cleaned.replace(".", "").replace(",", ".")
    elif has_comma:
        cleaned = cleaned.replace(".", "").replace(",", ".")

    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def _total_row_cost(raw_df: pd.DataFrame) -> float | None:
    total_rows = raw_df[raw_df["_cost_type"].str.lower() == "total"]
    if total_rows.empty:
        return None
    return float(total_rows["_cost"].sum())


def _filter_usage_rows(raw_df: pd.DataFrame) -> pd.DataFrame:
    mask = (raw_df["_cost_type"].str.lower() == "usage") & (raw_df["_service"] != "")
    return raw_df.loc[mask].reset_index(drop=True)


def _period(raw_df: pd.DataFrame) -> pd.Series:
    start_dates = pd.to_datetime(raw_df["Usage start date"], errors="coerce")
    return start_dates.dt.strftime("%Y-%m").fillna("")


def _infer_region(raw_df: pd.DataFrame) -> pd.Series:
    sku = raw_df["SKU description"].fillna("").astype(str)
    region = pd.Series(["global"] * len(raw_df), index=raw_df.index)

    region_patterns = {
        "Americas": "americas",
        "Sao Paulo": "southamerica-east1",
        "Virginia": "us-east4",
        "Iowa": "us-central1",
        "Oregon": "us-west1",
        "Netherlands": "europe-west4",
        "Finland": "europe-north1",
        "Belgium": "europe-west1",
        "Frankfurt": "europe-west3",
        "London": "europe-west2",
        "Tokyo": "asia-northeast1",
        "Seoul": "asia-northeast3",
        "Singapore": "asia-southeast1",
        "Sydney": "australia-southeast1",
    }
    for marker, gcp_region in region_patterns.items():
        region = region.mask(
            sku.str.contains(marker, case=False, na=False, regex=False),
            gcp_region,
        )
    return region


def _build_data_quality(
    *,
    metadata: dict[str, str],
    raw_df: pd.DataFrame,
    final_df: pd.DataFrame,
    raw_cost_total: float,
    currency: str,
) -> pd.DataFrame:
    usage_rows = raw_df[raw_df["_cost_type"].str.lower() == "usage"]
    usage_cost_total = float(usage_rows["_cost"].sum())
    final_cost_total = float(final_df["cost"].sum())

    rows = [
        {"metric": "parser", "value": "gcp-cost-table"},
        {"metric": "currency", "value": currency},
        {"metric": "invoice_number", "value": metadata.get("Invoice number", "")},
        {"metric": "invoice_date", "value": metadata.get("Invoice date", "")},
        {"metric": "billing_account_id", "value": metadata.get("Billing account ID", "")},
        {
            "metric": "currency_exchange_rate",
            "value": metadata.get("Currency exchange rate", ""),
        },
        {"metric": "total_amount_due_raw", "value": metadata.get("Total amount due", "")},
        {"metric": "input_rows_raw", "value": len(raw_df)},
        {"metric": "input_cost_total_row", "value": raw_cost_total},
        {"metric": "usage_rows", "value": len(usage_rows)},
        {"metric": "usage_cost_total", "value": usage_cost_total},
        {"metric": "final_rows", "value": len(final_df)},
        {"metric": "final_cost_total", "value": final_cost_total},
        {"metric": "removed_non_usage_rows", "value": len(raw_df) - len(usage_rows)},
        {
            "metric": "removed_blank_service_usage_rows",
            "value": len(usage_rows) - len(final_df),
        },
        {
            "metric": "reconciliation_delta_total_vs_final",
            "value": raw_cost_total - final_cost_total,
        },
    ]
    return pd.DataFrame(rows)
