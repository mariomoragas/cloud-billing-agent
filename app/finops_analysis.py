from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import pandas as pd


APP_ROOT = Path(__file__).resolve().parent
DEFAULT_PRICE_CATALOG = APP_ROOT / "mappings" / "oci_price_catalog.csv"


@dataclass
class FinOpsArtifacts:
    baseline_df: pd.DataFrame
    categories_df: pd.DataFrame
    effective_rates_df: pd.DataFrame
    coverage_df: pd.DataFrame
    priced_bom_df: pd.DataFrame
    savings_df: pd.DataFrame
    business_case_df: pd.DataFrame
    waterfall_df: pd.DataFrame
    sensitivity_df: pd.DataFrame


def build_finops_analysis(
    raw_df: pd.DataFrame,
    oci_bom_df: pd.DataFrame,
    *,
    price_catalog_path: Path = DEFAULT_PRICE_CATALOG,
) -> FinOpsArtifacts:
    working = _prepare_raw(raw_df)
    baseline_df = _build_baseline(working)
    categories_df = _build_categories(working)
    effective_rates_df = _build_effective_rates(working)
    priced_bom_df = _price_bom(oci_bom_df, price_catalog_path)
    savings_df = _build_savings(priced_bom_df)
    business_case_df, waterfall_df, sensitivity_df = _build_business_case(
        baseline_df, savings_df
    )
    coverage_df = _build_coverage(working, priced_bom_df, baseline_df)
    return FinOpsArtifacts(
        baseline_df=baseline_df,
        categories_df=categories_df,
        effective_rates_df=effective_rates_df,
        coverage_df=coverage_df,
        priced_bom_df=priced_bom_df,
        savings_df=savings_df,
        business_case_df=business_case_df,
        waterfall_df=waterfall_df,
        sensitivity_df=sensitivity_df,
    )


def _prepare_raw(raw_df: pd.DataFrame) -> pd.DataFrame:
    df = raw_df.copy()
    if "usage_unit" in df.columns:
        unit = df["usage_unit"].fillna("").astype(str).str.strip().str.lower()
        df = df.loc[unit != "reconciliation"].copy()
    for column in (
        "cloud", "service_name_original", "product_code", "sku", "usage_type",
        "usage_unit", "region", "currency", "linked_account_name",
    ):
        if column not in df.columns:
            df[column] = ""
        df[column] = df[column].fillna("").astype(str).str.strip()
    for column in ("usage_quantity", "cost", "cost_before_tax", "tax_amount", "credits"):
        if column not in df.columns:
            df[column] = 0.0
        df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0.0)
    if not df["cost_before_tax"].abs().gt(0).any():
        df["cost_before_tax"] = df["cost"] - df["tax_amount"]
    df["analysis_category"] = df.apply(_category_for_row, axis=1)
    return df


def _build_baseline(df: pd.DataFrame) -> pd.DataFrame:
    gross = float(df["cost"].sum())
    pretax = float(df["cost_before_tax"].sum())
    tax = float(df["tax_amount"].sum())
    credits = float(df["credits"].sum())
    currency = _mode(df["currency"], "USD")
    if abs(gross - (pretax + tax)) > 0.02 and abs(tax) < 0.01:
        pretax = gross
    rows = [
        ("Gross billed cost", gross, "Exact sum of normalized billing cost."),
        ("Comparable pre-tax baseline", pretax, "Taxes excluded when the source provides them."),
        ("Taxes", tax, "Reported separately; no automatic OCI equivalence."),
        ("Credits and discounts", credits, "Reported separately; sign follows the source export."),
    ]
    return pd.DataFrame(
        [{"metric": metric, "value": value, "currency": currency, "evidence": note}
         for metric, value, note in rows]
    )


def _build_categories(df: pd.DataFrame) -> pd.DataFrame:
    grouped = (
        df.groupby("analysis_category", as_index=False, dropna=False)
        .agg(
            gross_cost=("cost", "sum"),
            pre_tax_cost=("cost_before_tax", "sum"),
            tax_amount=("tax_amount", "sum"),
            record_count=("cost", "size"),
        )
    )
    total = float(grouped["gross_cost"].sum())
    grouped["gross_share_pct"] = grouped["gross_cost"].div(total).mul(100.0) if total else 0.0
    grouped["currency"] = _mode(df["currency"], "USD")
    return grouped.sort_values("gross_cost", ascending=False).reset_index(drop=True)


def _build_effective_rates(df: pd.DataFrame) -> pd.DataFrame:
    service = df["product_code"].where(df["product_code"] != "", df["service_name_original"])
    sku = df["usage_type"].where(df["usage_type"] != "", df["sku"])
    work = df.assign(source_service=service.replace("", "UNSPECIFIED"), source_sku=sku.replace("", "UNSPECIFIED"))
    work = work[(work["usage_quantity"] > 0) & (work["cost_before_tax"] > 0)].copy()
    if work.empty:
        return pd.DataFrame(columns=["source_service", "source_sku", "usage_unit", "quantity", "pre_tax_cost", "effective_rate", "currency", "confidence"])
    grouped = (
        work.groupby(["source_service", "source_sku", "usage_unit", "currency"], as_index=False, dropna=False)
        .agg(quantity=("usage_quantity", "sum"), pre_tax_cost=("cost_before_tax", "sum"), record_count=("cost", "size"))
    )
    grouped = grouped[grouped["quantity"] > 0].copy()
    grouped["effective_rate"] = grouped["pre_tax_cost"] / grouped["quantity"]
    grouped["confidence"] = grouped["usage_unit"].apply(lambda value: "High" if str(value).strip() else "Low")
    return grouped.sort_values("pre_tax_cost", ascending=False).head(50).reset_index(drop=True)


def _price_bom(bom_df: pd.DataFrame, price_catalog_path: Path) -> pd.DataFrame:
    columns = [
        "oci_part_number", "oci_sku", "target_profile", "workload_segment",
        "quantity", "unit", "unit_price_usd", "free_allowance", "chargeable_quantity",
        "estimated_oci_cost_usd", "source_cost", "source_cost_before_tax",
        "source_service", "source_shapes", "region", "confidence", "pricing_status",
        "pricing_note", "price_as_of", "catalog_url", "source_group_id",
    ]
    if bom_df is None or bom_df.empty or not price_catalog_path.exists():
        return pd.DataFrame(columns=columns)
    catalog = pd.read_csv(price_catalog_path).fillna("")
    priced = bom_df.copy()
    priced = priced.merge(catalog, on="oci_part_number", how="left", suffixes=("", "_price"))
    for column in ("quantity", "source_cost", "source_cost_before_tax", "unit_price_usd", "free_allowance"):
        if column not in priced.columns:
            priced[column] = 0.0
        priced[column] = pd.to_numeric(priced[column], errors="coerce").fillna(0.0)
    if "source_group_id" not in priced.columns:
        priced["source_group_id"] = priced.index.astype(str)
    priced["workload_segment"] = priced["target_profile"].map(_segment_for_profile)
    priced["chargeable_quantity"] = priced["quantity"]
    for part_number, indexes in priced.groupby("oci_part_number", sort=False).groups.items():
        allowance = float(priced.loc[list(indexes), "free_allowance"].max())
        remaining = allowance
        for index in indexes:
            quantity = float(priced.at[index, "quantity"])
            applied = min(quantity, remaining)
            priced.at[index, "chargeable_quantity"] = max(quantity - applied, 0.0)
            remaining = max(remaining - applied, 0.0)
    priced["estimated_oci_cost_usd"] = priced["chargeable_quantity"] * priced["unit_price_usd"]
    priced["pricing_status"] = priced["unit_price_usd"].gt(0).map({True: "Priced", False: "Free tier or zero price"})
    missing = priced["price_as_of"].astype(str).str.strip().eq("")
    priced.loc[missing, "pricing_status"] = "Price unavailable"
    priced.loc[missing, "estimated_oci_cost_usd"] = 0.0
    for column in columns:
        if column not in priced.columns:
            priced[column] = ""
    return priced[columns].sort_values("estimated_oci_cost_usd", ascending=False).reset_index(drop=True)


def _build_savings(priced: pd.DataFrame) -> pd.DataFrame:
    columns = ["workload_segment", "source_pre_tax_cost", "oci_estimated_cost", "savings_value", "savings_pct", "priced_skus", "confidence", "status"]
    if priced.empty:
        return pd.DataFrame(columns=columns)
    group_cost = (
        priced.groupby(["source_group_id", "workload_segment"], as_index=False)
        .agg(
            source_pre_tax_cost=("source_cost_before_tax", "max"),
            oci_estimated_cost=("estimated_oci_cost_usd", "sum"),
            priced_skus=("oci_part_number", "nunique"),
            confidence=("confidence", _lowest_confidence),
            unavailable=("pricing_status", lambda values: any(value == "Price unavailable" for value in values)),
        )
    )
    grouped = (
        group_cost.groupby("workload_segment", as_index=False)
        .agg(
            source_pre_tax_cost=("source_pre_tax_cost", "sum"),
            oci_estimated_cost=("oci_estimated_cost", "sum"),
            priced_skus=("priced_skus", "sum"),
            confidence=("confidence", _lowest_confidence),
            unavailable=("unavailable", "max"),
        )
    )
    grouped["savings_value"] = grouped["source_pre_tax_cost"] - grouped["oci_estimated_cost"]
    grouped["savings_pct"] = grouped["savings_value"].div(grouped["source_pre_tax_cost"]).mul(100.0).where(grouped["source_pre_tax_cost"].ne(0), 0.0)
    grouped["status"] = grouped["unavailable"].map({True: "Partial pricing", False: "Priced"})
    return grouped[columns].sort_values("source_pre_tax_cost", ascending=False).reset_index(drop=True)


def _build_business_case(baseline: pd.DataFrame, savings: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    baseline_value = _baseline_value(baseline, "Comparable pre-tax baseline")
    mapped_source = float(savings["source_pre_tax_cost"].sum()) if not savings.empty else 0.0
    mapped_oci = float(savings["oci_estimated_cost"].sum()) if not savings.empty else 0.0
    carry_forward = max(baseline_value - mapped_source, 0.0)
    gpu = savings[savings["workload_segment"].str.startswith("GPU", na=False)] if not savings.empty else savings
    gpu_source = float(gpu["source_pre_tax_cost"].sum()) if not gpu.empty else 0.0
    gpu_oci = float(gpu["oci_estimated_cost"].sum()) if not gpu.empty else 0.0
    scenarios = [
        _scenario("Capacity equivalent", baseline_value, mapped_oci + carry_forward, mapped_source, carry_forward, "No rightsizing or contract discount."),
        _scenario("Retain source GPU", baseline_value, mapped_oci - gpu_oci + gpu_source + carry_forward, mapped_source - gpu_source, carry_forward + gpu_source, "GPU remains at source cost; other priced workloads move to OCI."),
        _scenario("OCI price sensitivity -20%", baseline_value, mapped_oci * 0.80 + carry_forward, mapped_source, carry_forward, "Sensitivity only; requires an OCI commercial quote."),
    ]
    business = pd.DataFrame(scenarios)
    waterfall_rows = [{"step": "Comparable pre-tax baseline", "amount": baseline_value, "type": "total", "order": 0}]
    for index, row in savings.sort_values("source_pre_tax_cost", ascending=False).iterrows():
        waterfall_rows.append({"step": str(row["workload_segment"]), "amount": -float(row["savings_value"]), "type": "delta", "order": index + 1})
    waterfall_rows.append({"step": "Capacity-equivalent result", "amount": mapped_oci + carry_forward, "type": "total", "order": len(waterfall_rows)})
    sensitivity_rows = []
    for reduction in (0, 10, 20, 30, 40, 50):
        projected = mapped_oci - (gpu_oci * reduction / 100.0) + carry_forward
        saving = baseline_value - projected
        sensitivity_rows.append({
            "gpu_capacity_reduction_pct": reduction,
            "projected_total_cost": projected,
            "savings_value": saving,
            "savings_pct": (saving / baseline_value * 100.0) if baseline_value else 0.0,
            "assumption": "GPU OCI quantity sensitivity; validate occupancy, concurrency and benchmark results.",
        })
    sensitivity = pd.DataFrame(sensitivity_rows)
    return business, pd.DataFrame(waterfall_rows), sensitivity


def _scenario(name: str, baseline: float, projected: float, mapped_source: float, carry: float, assumption: str) -> dict[str, object]:
    saving = baseline - projected
    return {
        "scenario": name,
        "baseline_pre_tax": baseline,
        "projected_total_cost": projected,
        "savings_value": saving,
        "savings_pct": (saving / baseline * 100.0) if baseline else 0.0,
        "priced_source_scope": mapped_source,
        "unpriced_carry_forward": carry,
        "confidence": "Medium" if mapped_source > 0 else "Low",
        "assumption": assumption,
    }


def _build_coverage(df: pd.DataFrame, priced: pd.DataFrame, baseline: pd.DataFrame) -> pd.DataFrame:
    total = float(df["cost_before_tax"].abs().sum())
    baseline_value = _baseline_value(baseline, "Comparable pre-tax baseline")
    priced_source = 0.0
    if not priced.empty:
        priced_source = float(priced.groupby("source_group_id")["source_cost_before_tax"].max().sum())
    metrics = [
        ("Region cost coverage", _field_cost_coverage(df, "region"), "Source billing"),
        ("Usage unit cost coverage", _field_cost_coverage(df, "usage_unit"), "Source billing"),
        ("Product code cost coverage", _field_cost_coverage(df, "product_code"), "Source billing"),
        ("Account cost coverage", _field_cost_coverage(df, "linked_account_name"), "Source billing"),
        ("Tag cost coverage", _tag_cost_coverage(df), "Source billing"),
        ("OCI priced scope coverage", (priced_source / baseline_value * 100.0) if baseline_value else 0.0, "Versioned OCI price catalog"),
    ]
    rows = []
    for metric, value, source in metrics:
        rows.append({"metric": metric, "coverage_pct": value, "confidence": _coverage_confidence(value), "source": source})
    rows.append({"metric": "Absolute pre-tax cost analyzed", "coverage_pct": total, "confidence": "N/A", "source": "Billing currency amount, not a percentage"})
    return pd.DataFrame(rows)


def _field_cost_coverage(df: pd.DataFrame, column: str) -> float:
    if column not in df.columns:
        return 0.0
    total = float(df["cost_before_tax"].abs().sum())
    present = df[column].fillna("").astype(str).str.strip().ne("")
    return float(df.loc[present, "cost_before_tax"].abs().sum() / total * 100.0) if total else 0.0


def _tag_cost_coverage(df: pd.DataFrame) -> float:
    tag_columns = [column for column in df.columns if column.startswith("tag:")]
    if not tag_columns:
        return 0.0
    present = df[tag_columns].fillna("").astype(str).apply(lambda row: any(value.strip() for value in row), axis=1)
    total = float(df["cost_before_tax"].abs().sum())
    return float(df.loc[present, "cost_before_tax"].abs().sum() / total * 100.0) if total else 0.0


def _category_for_row(row: pd.Series) -> str:
    text = " ".join(str(row.get(column, "")) for column in ("product_code", "service_name_original", "sku", "usage_type")).lower()
    if any(token in text for token in ("tax", "imposto")):
        return "Taxes"
    if any(token in text for token in ("support", "suporte")):
        return "Support"
    if any(token in text for token in ("marketplace", "software subscription")):
        return "Marketplace"
    if any(token in text for token in ("ec2", "compute engine", "virtual machine", "block storage", "ebs", "disk")):
        return "Compute & Block Storage"
    if any(token in text for token in ("s3", "object storage", "blob", "file storage", "efs", "archive")):
        return "Storage"
    if any(token in text for token in ("data transfer", "network", "load balanc", "nat gateway", "cdn", "cloudfront", "dns", "route 53")):
        return "Network"
    if any(token in text for token in ("database", "rds", "sql", "dynamodb", "bigquery", "spanner", "cosmos", "redis", "oracle")):
        return "Database"
    if any(token in text for token in ("kubernetes", "container", "eks", "aks", "gke", "ecr")):
        return "Containers"
    if any(token in text for token in ("cloudwatch", "monitor", "security", "guard", "defender", "logging", "key management", "kms")):
        return "Security & Observability"
    return "Managed Services & Other"


def _segment_for_profile(profile: object) -> str:
    value = str(profile)
    mapping = {
        "compute_a1_ocpu": "Compute Arm/A1", "compute_a1_memory": "Compute Arm/A1",
        "compute_e5_ocpu": "Compute x86/E5", "compute_e5_memory": "Compute x86/E5",
        "gpu_l40s": "GPU L40S", "gpu_a10": "GPU A10",
        "block_storage": "Block Storage", "block_vpu": "Block Storage",
        "object_standard": "Object Storage", "object_requests": "Object Storage",
        "object_archive": "Object Storage", "file_storage": "File Storage",
        "egress_na_europe": "Data Transfer", "egress_apac_south_america": "Data Transfer",
        "load_balancer_base": "Load Balancer",
    }
    return mapping.get(value, value or "Other")


def _lowest_confidence(values: pd.Series) -> str:
    rank = {"High": 3, "Medium": 2, "Low": 1}
    cleaned = [str(value) for value in values if str(value) in rank]
    return min(cleaned, key=lambda value: rank[value]) if cleaned else "Low"


def _coverage_confidence(value: float) -> str:
    if value >= 90:
        return "High"
    if value >= 60:
        return "Medium"
    return "Low"


def _baseline_value(df: pd.DataFrame, metric: str) -> float:
    match = df.loc[df["metric"].eq(metric), "value"]
    return float(match.iloc[0]) if not match.empty else 0.0


def _mode(values: pd.Series, default: str) -> str:
    cleaned = values.fillna("").astype(str).str.strip()
    cleaned = cleaned[cleaned != ""]
    return str(cleaned.mode().iloc[0]) if not cleaned.empty else default
