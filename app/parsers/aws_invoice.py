from __future__ import annotations

from pathlib import Path

import pandas as pd

from app.input_validation import validate_billing_input_file
from app.report_types import ParserResult


AWS_INVOICE_REQUIRED_COLUMNS = {
    "InvoiceID",
    "RecordType",
    "ProductCode",
    "ProductName",
    "UsageType",
    "UsageQuantity",
    "CurrencyCode",
    "CostBeforeTax",
    "TaxAmount",
    "TotalCost",
}

REGION_PREFIX_MAP = {
    "SAE1": "sa-east-1",
    "USE1": "us-east-1",
    "USE2": "us-east-2",
    "USW1": "us-west-1",
    "USW2": "us-west-2",
    "EUW1": "eu-west-1",
    "EUW2": "eu-west-2",
    "EUC1": "eu-central-1",
    "APS1": "ap-south-1",
    "APN1": "ap-northeast-1",
    "APN2": "ap-northeast-2",
    "APS2": "ap-southeast-2",
    "CAC1": "ca-central-1",
}

DESCRIPTION_REGION_MAP = {
    "South America (Sao Paulo)": "sa-east-1",
    "US East (Northern Virginia)": "us-east-1",
    "US East (Ohio)": "us-east-2",
    "US West (Oregon)": "us-west-2",
    "US West (Northern California)": "us-west-1",
    "EU (London)": "eu-west-2",
    "EU (Germany)": "eu-central-1",
    "Canada (Central)": "ca-central-1",
}


def load_aws_invoice_csv(csv_path: Path) -> ParserResult:
    validate_billing_input_file(csv_path, "aws-invoice")
    raw_df = pd.read_csv(csv_path, low_memory=False)
    _validate_required_columns(raw_df)
    raw_cost_total = pd.to_numeric(raw_df["TotalCost"], errors="coerce").fillna(0.0).sum()

    normalized = pd.DataFrame(
        {
            "cloud": "aws",
            "invoice_id": raw_df["InvoiceID"],
            "product_code": raw_df["ProductCode"],
            "service_name_original": raw_df["ProductName"],
            "sku": raw_df["UsageType"],
            "region": _infer_region(raw_df),
            "usage_quantity": pd.to_numeric(
                raw_df["UsageQuantity"], errors="coerce"
            ).fillna(0.0),
            "usage_unit": _infer_usage_unit(raw_df),
            "cost": pd.to_numeric(raw_df["TotalCost"], errors="coerce").fillna(0.0),
            "currency": raw_df["CurrencyCode"].fillna("USD"),
            "period": _period(raw_df),
            "usage_type": raw_df["UsageType"].fillna(""),
            "operation": raw_df["Operation"].fillna("") if "Operation" in raw_df.columns else "",
            "record_type": raw_df["RecordType"].fillna(""),
            "line_item_type": raw_df["RecordType"].fillna(""),
            "payer_account_id": _optional(raw_df, "PayerAccountId"),
            "payer_account_name": _optional(raw_df, "PayerAccountName"),
            "linked_account_id": _optional(raw_df, "LinkedAccountId"),
            "linked_account_name": _optional(raw_df, "LinkedAccountName"),
            "seller_of_record": _optional(raw_df, "SellerOfRecord"),
            "item_description": _optional(raw_df, "ItemDescription"),
            "cost_before_tax": pd.to_numeric(
                raw_df["CostBeforeTax"], errors="coerce"
            ).fillna(0.0),
            "tax_amount": pd.to_numeric(raw_df["TaxAmount"], errors="coerce").fillna(0.0),
            "credits": pd.to_numeric(_optional(raw_df, "Credits"), errors="coerce").fillna(0.0),
            "tax_type": _optional(raw_df, "TaxType"),
            "purchase_option": "Invoice",
        }
    )

    for column in normalized.columns:
        if column not in {"usage_quantity", "cost", "cost_before_tax", "tax_amount", "credits"}:
            normalized[column] = normalized[column].fillna("").astype(str).str.strip()

    detail_df = _filter_invoice_detail_rows(normalized)
    linked_only_df = _filter_missing_linked_account_name(detail_df)
    deduped_df = _deduplicate_payer_linked_rows(linked_only_df)

    data_quality = _build_data_quality(
        raw_df=raw_df,
        detail_df=detail_df,
        linked_only_df=linked_only_df,
        final_df=deduped_df,
        raw_cost_total=raw_cost_total,
    )
    return ParserResult(dataframe=deduped_df, data_quality=data_quality)


def _validate_required_columns(raw_df: pd.DataFrame) -> None:
    missing = AWS_INVOICE_REQUIRED_COLUMNS.difference(raw_df.columns)
    if missing:
        raise ValueError(
            "CSV AWS Invoice sem colunas obrigatorias: " + ", ".join(sorted(missing))
        )


def _period(raw_df: pd.DataFrame) -> pd.Series:
    return pd.to_datetime(raw_df["BillingPeriodStartDate"], errors="coerce").dt.strftime(
        "%Y-%m"
    )


def _infer_region(raw_df: pd.DataFrame) -> pd.Series:
    usage_type = raw_df["UsageType"].fillna("").astype(str).str.strip()
    description = _optional(raw_df, "ItemDescription").fillna("").astype(str)
    prefix = usage_type.str.extract(r"^([A-Z0-9]+)[-:]")[0].fillna("")
    region = prefix.map(REGION_PREFIX_MAP).fillna("")

    for description_text, aws_region in DESCRIPTION_REGION_MAP.items():
        region = region.mask(
            (region == "")
            & description.str.contains(
                description_text,
                case=False,
                na=False,
                regex=False,
            ),
            aws_region,
        )

    global_products = raw_df["ProductCode"].fillna("").astype(str).isin(
        ["AmazonRoute53", "AWSDeveloperSupport", "AmazonCloudFront"]
    )
    region = region.mask(global_products, "global")
    return region


def _infer_usage_unit(raw_df: pd.DataFrame) -> pd.Series:
    usage_type = raw_df["UsageType"].fillna("").astype(str)
    description = _optional(raw_df, "ItemDescription").fillna("").astype(str)

    unit = pd.Series([""] * len(raw_df))
    unit = unit.mask(
        usage_type.str.contains("TimedStorage-ByteHrs|GlacierByteHrs", case=False, na=False),
        "GB-Month",
    )
    unit = unit.mask(
        (unit == "")
        & (
            usage_type.str.contains("ByteHrs", case=False, na=False)
        | (
            description.str.contains("per GB", case=False, na=False)
            & ~description.str.contains("per hour", case=False, na=False)
        )),
        "GB",
    )
    unit = unit.mask(
        (unit == "")
        & (
            usage_type.str.contains("Requests", case=False, na=False)
            | description.str.contains("request", case=False, na=False)
        ),
        "Requests",
    )
    unit = unit.mask(
        (unit == "")
        & (
            usage_type.str.contains(
                "Hrs|Hours|BoxUsage|InstanceUsage", case=False, na=False
            )
            | description.str.contains("per hour", case=False, na=False)
        ),
        "Hours",
    )
    unit = unit.mask(
        (unit == "") & usage_type.str.contains("HostedZone", case=False, na=False),
        "HostedZone",
    )
    return unit


def _filter_invoice_detail_rows(normalized_df: pd.DataFrame) -> pd.DataFrame:
    keep_record_types = {"LinkedLineItem", "PayerLineItem"}
    mask = normalized_df["record_type"].isin(keep_record_types)
    return normalized_df.loc[mask].reset_index(drop=True)


def _filter_missing_linked_account_name(normalized_df: pd.DataFrame) -> pd.DataFrame:
    linked_account_name = (
        normalized_df["linked_account_name"].fillna("").astype(str).str.strip()
    )
    mask = linked_account_name != ""
    return normalized_df.loc[mask].reset_index(drop=True)


def _deduplicate_payer_linked_rows(normalized_df: pd.DataFrame) -> pd.DataFrame:
    dedupe_key = [
        "invoice_id",
        "service_name_original",
        "sku",
        "operation",
        "item_description",
        "usage_quantity",
        "cost_before_tax",
        "tax_amount",
        "cost",
        "period",
    ]

    working = normalized_df.copy()
    working["_has_linked_account"] = (
        working["linked_account_name"].fillna("").astype(str).str.strip() != ""
    ) | (working["linked_account_id"].fillna("").astype(str).str.strip() != "")
    working["_record_rank"] = working["record_type"].map(
        {"LinkedLineItem": 0, "PayerLineItem": 1}
    ).fillna(2)
    working["_row_order"] = range(len(working))

    keep_rows: list[pd.DataFrame] = []

    for _, group in working.groupby(dedupe_key, dropna=False, sort=False):
        if len(group) == 1:
            keep_rows.append(group)
            continue

        has_linked = bool(group["_has_linked_account"].any())
        has_unlinked = bool((~group["_has_linked_account"]).any())

        if has_linked and has_unlinked:
            keep_group = group[group["_has_linked_account"]].copy()
        else:
            keep_group = group.copy()

        keep_group = keep_group.sort_values(
            by=["_record_rank", "_row_order"], ascending=[True, True]
        )
        keep_rows.append(keep_group)

    deduped = pd.concat(keep_rows, ignore_index=True)
    return deduped.drop(columns=["_has_linked_account", "_record_rank", "_row_order"])


def _optional(raw_df: pd.DataFrame, column: str) -> pd.Series:
    if column in raw_df.columns:
        return raw_df[column]
    return pd.Series([""] * len(raw_df))


def _build_data_quality(
    raw_df: pd.DataFrame,
    detail_df: pd.DataFrame,
    linked_only_df: pd.DataFrame,
    final_df: pd.DataFrame,
    raw_cost_total: float,
) -> pd.DataFrame:
    detail_cost_total = detail_df["cost"].sum()
    linked_only_cost_total = linked_only_df["cost"].sum()
    final_cost_total = final_df["cost"].sum()

    removed_non_detail_rows = len(raw_df) - len(detail_df)
    removed_missing_linked_rows = len(detail_df) - len(linked_only_df)
    removed_duplicate_rows = len(linked_only_df) - len(final_df)

    rows = [
        {"metric": "input_rows_raw", "value": len(raw_df)},
        {"metric": "input_cost_raw", "value": raw_cost_total},
        {"metric": "detail_rows", "value": len(detail_df)},
        {"metric": "detail_cost_total", "value": detail_cost_total},
        {"metric": "removed_non_detail_rows", "value": removed_non_detail_rows},
        {
            "metric": "removed_non_detail_cost",
            "value": raw_cost_total - detail_cost_total,
        },
        {"metric": "linked_only_rows", "value": len(linked_only_df)},
        {"metric": "linked_only_cost_total", "value": linked_only_cost_total},
        {"metric": "removed_missing_linked_rows", "value": removed_missing_linked_rows},
        {
            "metric": "removed_missing_linked_cost",
            "value": detail_cost_total - linked_only_cost_total,
        },
        {"metric": "final_rows", "value": len(final_df)},
        {"metric": "final_cost_total", "value": final_cost_total},
        {"metric": "removed_duplicate_rows", "value": removed_duplicate_rows},
        {
            "metric": "removed_duplicate_cost",
            "value": linked_only_cost_total - final_cost_total,
        },
    ]
    return pd.DataFrame(rows)
