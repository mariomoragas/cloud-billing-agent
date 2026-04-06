from __future__ import annotations

import pandas as pd


def summarize_by_service(df: pd.DataFrame) -> pd.DataFrame:
    aggregations: dict[str, tuple[str, str | callable]] = {
        "total_usage_quantity": ("usage_quantity", "sum"),
        "total_cost": ("cost", "sum"),
        "record_count": ("service_name_original", "count"),
        "primary_unit": ("usage_unit", lambda values: _most_common_nonempty(values)),
        "primary_currency": ("currency", lambda values: _most_common_nonempty(values)),
    }
    if "cost_before_tax" in df.columns:
        aggregations["total_cost_before_tax"] = ("cost_before_tax", "sum")
    if "tax_amount" in df.columns:
        aggregations["total_tax_amount"] = ("tax_amount", "sum")
    if "credits" in df.columns:
        aggregations["total_credits"] = ("credits", "sum")

    summary = (
        df.groupby(["cloud", "service_name_original"], dropna=False, as_index=False)
        .agg(**aggregations)
        .sort_values("total_cost", ascending=False)
    )
    return summary


def summarize_by_region(df: pd.DataFrame) -> pd.DataFrame:
    aggregations: dict[str, tuple[str, str | callable]] = {
        "total_usage_quantity": ("usage_quantity", "sum"),
        "total_cost": ("cost", "sum"),
        "primary_currency": ("currency", lambda values: _most_common_nonempty(values)),
    }
    if "cost_before_tax" in df.columns:
        aggregations["total_cost_before_tax"] = ("cost_before_tax", "sum")
    if "tax_amount" in df.columns:
        aggregations["total_tax_amount"] = ("tax_amount", "sum")

    summary = (
        df.groupby(["cloud", "region"], dropna=False, as_index=False)
        .agg(**aggregations)
        .sort_values("total_cost", ascending=False)
    )
    return summary


def summarize_by_column(
    df: pd.DataFrame,
    column: str,
    label: str | None = None,
    top_n: int | None = None,
) -> pd.DataFrame:
    if column not in df.columns:
        return pd.DataFrame()

    display_column = label or column
    working_df = df.copy()
    working_df[column] = working_df[column].fillna("").astype(str).str.strip()
    working_df[column] = working_df[column].replace("", "UNSPECIFIED")

    summary = (
        working_df.groupby([column], dropna=False, as_index=False)
        .agg(
            total_usage_quantity=("usage_quantity", "sum"),
            total_cost=("cost", "sum"),
            record_count=(column, "count"),
            primary_unit=("usage_unit", lambda values: _most_common_nonempty(values)),
            primary_currency=("currency", lambda values: _most_common_nonempty(values)),
        )
        .sort_values("total_cost", ascending=False)
        .rename(columns={column: display_column})
    )

    if top_n is not None:
        summary = summary.head(top_n)
    return summary


def build_aws_enterprise_summaries(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    summaries: dict[str, pd.DataFrame] = {}
    summary_specs = [
        ("linked_account_name", "linked_account_name", 20),
        ("usage_type", "usage_type", 20),
    ]

    for column, label, top_n in summary_specs:
        summary = summarize_by_column(df, column=column, label=label, top_n=top_n)
        if not summary.empty:
            summaries[label] = summary

    tag_summary = summarize_aws_tags(df, top_n_keys=10)
    if not tag_summary.empty:
        summaries["cost_allocation_tags"] = tag_summary

    return summaries


def summarize_aws_tags(df: pd.DataFrame, top_n_keys: int = 10) -> pd.DataFrame:
    tag_columns = [column for column in df.columns if column.startswith("tag:")]
    if not tag_columns:
        return pd.DataFrame()

    rows: list[dict[str, object]] = []
    for tag_column in tag_columns:
        tag_values = df[tag_column].fillna("").astype(str).str.strip()
        mask = tag_values != ""
        if not mask.any():
            continue

        scoped = df.loc[mask].copy()
        scoped[tag_column] = tag_values.loc[mask]
        grouped = (
            scoped.groupby(tag_column, as_index=False)
            .agg(
                total_usage_quantity=("usage_quantity", "sum"),
                total_cost=("cost", "sum"),
                primary_unit=("usage_unit", lambda values: _most_common_nonempty(values)),
                primary_currency=("currency", lambda values: _most_common_nonempty(values)),
            )
            .sort_values("total_cost", ascending=False)
            .head(3)
        )

        for _, row in grouped.iterrows():
            rows.append(
                {
                    "tag_key": tag_column.removeprefix("tag:"),
                    "tag_value": row[tag_column],
                    "total_usage_quantity": row["total_usage_quantity"],
                    "primary_unit": row["primary_unit"],
                    "total_cost": row["total_cost"],
                    "primary_currency": row["primary_currency"],
                }
            )

    if not rows:
        return pd.DataFrame()

    summary = pd.DataFrame(rows).sort_values("total_cost", ascending=False)
    top_keys = summary.groupby("tag_key")["total_cost"].sum().sort_values(ascending=False)
    keep_keys = set(top_keys.head(top_n_keys).index)
    return summary[summary["tag_key"].isin(keep_keys)].reset_index(drop=True)


def _most_common_nonempty(values: pd.Series) -> str:
    cleaned = values.fillna("").astype(str).str.strip()
    cleaned = cleaned[cleaned != ""]
    if cleaned.empty:
        return ""
    return cleaned.mode().iloc[0]
