from __future__ import annotations

from pathlib import Path

import pandas as pd

from app.input_validation import validate_billing_input_file
from app.report_types import ParserResult

REQUIRED_COLUMNS = {
    "categoria",
    "nome_do_produto",
    "qtd",
    "un",
    "consumo",
}


def load_azure_cost_csv(csv_path: Path) -> ParserResult:
    validate_billing_input_file(csv_path, "azure-cost-csv")
    raw_df = _read_csv_with_fallbacks(csv_path)
    normalized_columns = _normalize_columns(raw_df.columns)
    raw_df = raw_df.rename(columns=normalized_columns)
    _validate_required_columns(raw_df)

    period_series = _parse_period(raw_df.get("data", pd.Series([""] * len(raw_df))))
    currency = _infer_currency(raw_df)

    normalized = pd.DataFrame(
        {
            "cloud": "azure",
            "service_name_original": raw_df["categoria"].fillna("").astype(str).str.strip(),
            "sku": raw_df["nome_do_produto"].fillna("").astype(str).str.strip(),
            "region": raw_df.get("local", pd.Series([""] * len(raw_df))).fillna("").astype(str).str.strip(),
            "usage_quantity": pd.to_numeric(raw_df["qtd"], errors="coerce").fillna(0.0),
            "usage_unit": raw_df["un"].fillna("").astype(str).str.strip(),
            "cost": pd.to_numeric(raw_df["consumo"], errors="coerce").fillna(0.0),
            "currency": currency,
            "period": period_series,
            "resource_group": raw_df.get("grupo", pd.Series([""] * len(raw_df))).fillna("").astype(str).str.strip(),
            "resource_name": raw_df.get("nome_do_recurso", pd.Series([""] * len(raw_df))).fillna("").astype(str).str.strip(),
            "instance_name": raw_df.get("instancia", pd.Series([""] * len(raw_df))).fillna("").astype(str).str.strip(),
            "subcategory": raw_df.get("subcategoria", pd.Series([""] * len(raw_df))).fillna("").astype(str).str.strip(),
        }
    )

    data_quality = pd.DataFrame(
        [
            {"metric": "input_rows_raw", "value": len(raw_df)},
            {"metric": "final_rows", "value": len(normalized)},
            {"metric": "final_cost_total", "value": float(normalized["cost"].sum())},
            {"metric": "detected_currency", "value": currency},
            {"metric": "distinct_services", "value": int(normalized["service_name_original"].nunique())},
            {"metric": "distinct_regions", "value": int(normalized["region"].replace("", pd.NA).dropna().nunique())},
        ]
    )

    return ParserResult(dataframe=normalized, data_quality=data_quality)


def _read_csv_with_fallbacks(csv_path: Path) -> pd.DataFrame:
    attempts = [
        {"sep": ",", "encoding": "utf-8-sig"},
        {"sep": ";", "encoding": "utf-8-sig"},
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
        + " | ".join(errors[:2])
    )


def _normalize_columns(columns) -> dict[str, str]:
    mapping: dict[str, str] = {}
    for original in columns:
        normalized = (
            str(original)
            .strip()
            .lower()
            .replace(" ", "_")
            .replace('"', "")
            .replace("ã", "a")
            .replace("á", "a")
            .replace("à", "a")
            .replace("â", "a")
            .replace("é", "e")
            .replace("ê", "e")
            .replace("í", "i")
            .replace("ó", "o")
            .replace("ô", "o")
            .replace("õ", "o")
            .replace("ú", "u")
            .replace("ç", "c")
        )
        mapping[original] = normalized
    return mapping


def _validate_required_columns(df: pd.DataFrame) -> None:
    missing = REQUIRED_COLUMNS.difference(df.columns)
    if missing:
        raise ValueError(
            "CSV Azure sem colunas obrigatorias: " + ", ".join(sorted(missing))
        )


def _parse_period(series: pd.Series) -> pd.Series:
    parsed = pd.to_datetime(series.astype(str).str.strip(), dayfirst=True, errors="coerce")
    return parsed.dt.strftime("%Y-%m").fillna("")


def _infer_currency(df: pd.DataFrame) -> str:
    lower_cols = set(df.columns)
    if "moeda" in lower_cols:
        mode = df["moeda"].fillna("").astype(str).str.strip()
        mode = mode[mode != ""]
        if not mode.empty:
            return str(mode.mode().iloc[0]).upper()
    return "USD"
