from __future__ import annotations

from pathlib import Path
import re
import unicodedata

import pandas as pd

from app.input_validation import validate_billing_input_file
from app.report_types import ParserResult

LEGACY_REQUIRED_COLUMNS = {
    "categoria",
    "nome_do_produto",
    "qtd",
    "un",
    "consumo",
}

USAGE_REQUIRED_COLUMNS = {
    "date",
    "service_name",
    "service_type",
    "service_region",
    "service_resource",
    "metric",
    "quantity",
    "cost",
}


def load_azure_cost_csv(csv_path: Path) -> ParserResult:
    validate_billing_input_file(csv_path, "azure-cost-csv")
    raw_df = _read_csv_with_fallbacks(csv_path)
    raw_df = raw_df.rename(columns=_normalize_columns(raw_df.columns))
    schema = _detect_schema(raw_df)

    period_source = "data" if schema == "legacy_ptbr" else "date"
    period_series = _parse_period(raw_df.get(period_source, pd.Series([""] * len(raw_df))))
    currency = _infer_currency(raw_df)
    normalized = _build_normalized_dataframe(
        raw_df=raw_df,
        schema=schema,
        currency=currency,
        period_series=period_series,
    )

    data_quality = pd.DataFrame(
        [
            {"metric": "input_rows_raw", "value": len(raw_df)},
            {"metric": "final_rows", "value": len(normalized)},
            {"metric": "final_cost_total", "value": float(normalized["cost"].sum())},
            {"metric": "detected_currency", "value": currency},
            {"metric": "source_schema", "value": schema},
            {
                "metric": "distinct_services",
                "value": int(normalized["service_name_original"].nunique()),
            },
            {
                "metric": "distinct_regions",
                "value": int(normalized["region"].replace("", pd.NA).dropna().nunique()),
            },
        ]
    )

    return ParserResult(dataframe=normalized, data_quality=data_quality)


def _read_csv_with_fallbacks(csv_path: Path) -> pd.DataFrame:
    attempts = [
        {"sep": ",", "encoding": "utf-8-sig"},
        {"sep": ";", "encoding": "utf-8-sig"},
        {"sep": ",", "encoding": "cp1252"},
        {"sep": ";", "encoding": "cp1252"},
        {"sep": ",", "encoding": "latin-1"},
        {"sep": ";", "encoding": "latin-1"},
    ]

    errors: list[str] = []
    for options in attempts:
        try:
            df = pd.read_csv(csv_path, **options)
            if len(df.columns) > 1:
                return df
        except Exception as exc:  # pragma: no cover - defensive parser fallback
            errors.append(f"{options}: {exc}")
    raise ValueError(
        "Falha ao ler CSV de custo Azure com delimitadores/encodings conhecidos. "
        + " | ".join(errors[:3])
    )


def _normalize_columns(columns) -> dict[str, str]:
    return {original: _normalize_column_name(original) for original in columns}


def _normalize_column_name(value: object) -> str:
    text = str(value).strip().replace('"', "")
    text = re.sub(r"(?<=[a-z0-9])(?=[A-Z])", "_", text)
    text = unicodedata.normalize("NFKD", text)
    text = text.encode("ascii", "ignore").decode("ascii")
    text = text.lower()
    text = re.sub(r"[^a-z0-9]+", "_", text)
    return text.strip("_")


def _detect_schema(df: pd.DataFrame) -> str:
    if LEGACY_REQUIRED_COLUMNS.issubset(df.columns):
        return "legacy_ptbr"
    if USAGE_REQUIRED_COLUMNS.issubset(df.columns):
        return "usage_export"

    missing_legacy = LEGACY_REQUIRED_COLUMNS.difference(df.columns)
    missing_usage = USAGE_REQUIRED_COLUMNS.difference(df.columns)
    raise ValueError(
        "CSV Azure sem colunas obrigatorias conhecidas. "
        f"Formato legado ausente: {', '.join(sorted(missing_legacy))}. "
        f"Formato usage export ausente: {', '.join(sorted(missing_usage))}."
    )


def _build_normalized_dataframe(
    *,
    raw_df: pd.DataFrame,
    schema: str,
    currency: str,
    period_series: pd.Series,
) -> pd.DataFrame:
    if schema == "legacy_ptbr":
        return pd.DataFrame(
            {
                "cloud": "azure",
                "service_name_original": raw_df["categoria"].fillna("").astype(str).str.strip(),
                "sku": raw_df["nome_do_produto"].fillna("").astype(str).str.strip(),
                "region": raw_df.get("local", pd.Series([""] * len(raw_df)))
                .fillna("")
                .astype(str)
                .str.strip(),
                "usage_quantity": pd.to_numeric(raw_df["qtd"], errors="coerce").fillna(0.0),
                "usage_unit": raw_df["un"].fillna("").astype(str).str.strip(),
                "cost": pd.to_numeric(raw_df["consumo"], errors="coerce").fillna(0.0),
                "currency": currency,
                "period": period_series,
                "resource_group": raw_df.get("grupo", pd.Series([""] * len(raw_df)))
                .fillna("")
                .astype(str)
                .str.strip(),
                "resource_name": raw_df.get(
                    "nome_do_recurso", pd.Series([""] * len(raw_df))
                )
                .fillna("")
                .astype(str)
                .str.strip(),
                "instance_name": raw_df.get("instancia", pd.Series([""] * len(raw_df)))
                .fillna("")
                .astype(str)
                .str.strip(),
                "subcategory": raw_df.get("subcategoria", pd.Series([""] * len(raw_df)))
                .fillna("")
                .astype(str)
                .str.strip(),
            }
        )

    return pd.DataFrame(
        {
            "cloud": "azure",
            "service_name_original": raw_df["service_name"].fillna("").astype(str).str.strip(),
            "sku": raw_df["service_type"].fillna("").astype(str).str.strip(),
            "region": raw_df.get("service_region", pd.Series([""] * len(raw_df)))
            .fillna("")
            .astype(str)
            .str.strip(),
            "usage_quantity": pd.to_numeric(raw_df["quantity"], errors="coerce").fillna(0.0),
            "usage_unit": raw_df["metric"].fillna("").astype(str).str.strip(),
            "cost": pd.to_numeric(raw_df["cost"], errors="coerce").fillna(0.0),
            "currency": currency,
            "period": period_series,
            "resource_group": pd.Series([""] * len(raw_df)),
            "resource_name": raw_df.get("service_resource", pd.Series([""] * len(raw_df)))
            .fillna("")
            .astype(str)
            .str.strip(),
            "instance_name": raw_df.get("resource_guid", pd.Series([""] * len(raw_df)))
            .fillna("")
            .astype(str)
            .str.strip(),
            "subcategory": raw_df.get("service_type", pd.Series([""] * len(raw_df)))
            .fillna("")
            .astype(str)
            .str.strip(),
        }
    )


def _parse_period(series: pd.Series) -> pd.Series:
    parsed = pd.to_datetime(series.astype(str).str.strip(), errors="coerce")
    if parsed.isna().all():
        parsed = pd.to_datetime(series.astype(str).str.strip(), dayfirst=True, errors="coerce")
    return parsed.dt.strftime("%Y-%m").fillna("")


def _infer_currency(df: pd.DataFrame) -> str:
    lower_cols = set(df.columns)
    for column in ("moeda", "currency"):
        if column not in lower_cols:
            continue
        mode = df[column].fillna("").astype(str).str.strip()
        mode = mode[mode != ""]
        if not mode.empty:
            return str(mode.mode().iloc[0]).upper()
    return "USD"
