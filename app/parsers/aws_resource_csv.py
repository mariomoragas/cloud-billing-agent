from __future__ import annotations

import re
from pathlib import Path

import pandas as pd

from app.input_validation import validate_billing_input_file
from app.report_types import ParserResult


REQUIRED_COLUMNS = {
    "invoice",
    "conta_id",
    "conta_nome",
    "servico",
    "regiao",
    "usage_type",
    "quantidade_uso",
    "unidade",
    "custo_usd",
}

PRODUCT_CODE_BY_SERVICE = {
    "amazon elastic compute cloud": "AmazonEC2",
    "amazon simple storage service": "AmazonS3",
    "amazon elastic file system": "AmazonEFS",
    "aws data transfer": "AWSDataTransfer",
    "amazon elastic load balancing": "AWSELB",
    "elastic load balancing": "AWSELB",
    "amazon relational database service": "AmazonRDS",
    "amazon virtual private cloud": "AmazonVPC",
    "amazoncloudwatch": "AmazonCloudWatch",
    "amazon managed streaming for apache kafka": "AmazonMSK",
    "aws key management service": "awskms",
    "amazon route 53": "AmazonRoute53",
}

UNIT_ALIASES = {
    "hrs": "Hours",
    "hours": "Hours",
    "gb-mo": "GB-Month",
    "gb-month": "GB-Month",
    "gb": "GB",
    "requests": "Requests",
    "months": "Months",
}


def load_aws_resource_csv(csv_path: Path) -> ParserResult:
    validate_billing_input_file(csv_path, "aws-resource-csv")
    raw_df = pd.read_csv(
        csv_path,
        sep=";",
        dtype=str,
        encoding="utf-8-sig",
        keep_default_na=False,
        low_memory=False,
    )
    raw_df.columns = [str(column).strip() for column in raw_df.columns]
    missing = REQUIRED_COLUMNS.difference(raw_df.columns)
    if missing:
        raise ValueError(
            "CSV AWS por recurso sem colunas obrigatorias: "
            + ", ".join(sorted(missing))
        )

    service = raw_df["servico"].astype(str).str.strip()
    usage_type = raw_df["usage_type"].astype(str).str.strip()
    instance_type = _optional(raw_df, "tipo_instancia")
    usage_quantity = _parse_localized_number(raw_df["quantidade_uso"])
    cost = _parse_localized_number(raw_df["custo_usd"])
    list_cost = _parse_localized_number(_optional(raw_df, "custo_preco_tabela_usd"))
    credits = (cost - list_cost).where((list_cost > 0) & (list_cost >= cost), 0.0)

    normalized = pd.DataFrame(
        {
            "cloud": "aws",
            "invoice_id": raw_df["invoice"],
            "product_code": service.str.casefold().map(PRODUCT_CODE_BY_SERVICE).fillna(""),
            "service_name_original": service,
            "sku": usage_type,
            "region": raw_df["regiao"],
            "usage_quantity": usage_quantity,
            "usage_unit": _normalize_unit(raw_df["unidade"]),
            "cost": cost,
            "currency": "USD",
            "period": _period_from_filename(csv_path, len(raw_df)),
            "usage_type": usage_type,
            "operation": _optional(raw_df, "operacao"),
            "record_type": _optional(raw_df, "tipo_item"),
            "line_item_type": _optional(raw_df, "tipo_item"),
            "instance_type": instance_type,
            "instance_name": instance_type,
            "resource_id": _optional(raw_df, "resource_id"),
            "resource_name": _optional(raw_df, "resource_id"),
            "linked_account_id": raw_df["conta_id"],
            "linked_account_name": raw_df["conta_nome"],
            "seller_of_record": _optional(raw_df, "emissor"),
            "item_description": _optional(raw_df, "descricao"),
            "cost_before_tax": cost,
            "tax_amount": 0.0,
            "credits": credits,
            "list_cost": list_cost,
            "tag_sector": _optional(raw_df, "tag_SETOR"),
            "tag_created_by": _optional(raw_df, "tag_createdBy"),
            "purchase_option": "Usage",
        }
    )

    numeric_columns = {
        "usage_quantity",
        "cost",
        "cost_before_tax",
        "tax_amount",
        "credits",
        "list_cost",
    }
    for column in normalized.columns:
        if column not in numeric_columns:
            normalized[column] = normalized[column].fillna("").astype(str).str.strip()

    data_quality = pd.DataFrame(
        [
            {"metric": "input_rows_raw", "value": len(raw_df)},
            {"metric": "final_rows", "value": len(normalized)},
            {"metric": "input_cost_raw", "value": float(cost.sum())},
            {"metric": "final_cost_total", "value": float(normalized["cost"].sum())},
            {
                "metric": "rows_missing_service",
                "value": int(normalized["service_name_original"].eq("").sum()),
            },
            {
                "metric": "rows_invalid_cost",
                "value": int(raw_df["custo_usd"].astype(str).str.strip().ne("").sum() - cost.notna().sum()),
            },
        ]
    )
    return ParserResult(dataframe=normalized, data_quality=data_quality)


def _optional(raw_df: pd.DataFrame, column: str) -> pd.Series:
    if column in raw_df.columns:
        return raw_df[column]
    return pd.Series([""] * len(raw_df), index=raw_df.index, dtype="object")


def _parse_localized_number(series: pd.Series) -> pd.Series:
    text = series.fillna("").astype(str).str.strip().str.replace(" ", "", regex=False)
    comma_decimal = text.str.contains(",", regex=False)
    normalized = text.copy()
    normalized.loc[comma_decimal] = (
        normalized.loc[comma_decimal]
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
    )
    return pd.to_numeric(normalized, errors="coerce").fillna(0.0)


def _normalize_unit(series: pd.Series) -> pd.Series:
    text = series.fillna("").astype(str).str.strip()
    return text.str.casefold().map(UNIT_ALIASES).fillna(text)


def _period_from_filename(csv_path: Path, size: int) -> pd.Series:
    match = re.search(r"(20\d{2})[-_](0[1-9]|1[0-2])", csv_path.stem)
    period = f"{match.group(1)}-{match.group(2)}" if match else ""
    return pd.Series([period] * size, dtype="object")

