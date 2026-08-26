from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re

import pandas as pd


APP_ROOT = Path(__file__).resolve().parent
DEFAULT_SKU_CATALOG = APP_ROOT / "mappings" / "oci_sku_catalog.csv"
DEFAULT_SHAPE_CATALOG = APP_ROOT / "mappings" / "cloud_shape_capacity.csv"
SCENARIO = "Capacity equivalent"

BOM_COLUMNS = [
    "scenario",
    "oci_part_number",
    "oci_sku",
    "oci_service_category",
    "target_profile",
    "quantity",
    "unit",
    "source_cloud",
    "source_service",
    "source_skus",
    "source_shapes",
    "source_quantity",
    "source_unit",
    "source_cost",
    "currency",
    "region",
    "confidence",
    "assumption",
    "catalog_last_updated",
    "catalog_url",
]

REVIEW_COLUMNS = [
    "source_cloud",
    "source_service",
    "source_sku",
    "region",
    "source_quantity",
    "source_unit",
    "source_cost",
    "currency",
    "review_reason",
]


@dataclass
class OciBomArtifacts:
    bom_df: pd.DataFrame
    review_df: pd.DataFrame
    summary_df: pd.DataFrame


def build_oci_bom(
    raw_df: pd.DataFrame,
    *,
    sku_catalog_path: Path = DEFAULT_SKU_CATALOG,
    shape_catalog_path: Path = DEFAULT_SHAPE_CATALOG,
) -> OciBomArtifacts:
    if raw_df.empty:
        return _empty_artifacts()

    working = raw_df.copy().reset_index(drop=True)
    working["_source_row"] = working.index
    for column in (
        "cloud",
        "product_code",
        "service_name_original",
        "sku",
        "usage_type",
        "usage_unit",
        "region",
        "currency",
    ):
        if column not in working.columns:
            working[column] = ""
        working[column] = working[column].fillna("").astype(str).str.strip()
    for column in ("usage_quantity", "cost"):
        if column not in working.columns:
            working[column] = 0.0
        working[column] = pd.to_numeric(working[column], errors="coerce").fillna(0.0)

    sku_catalog = pd.read_csv(sku_catalog_path, dtype=str).fillna("")
    shape_catalog = pd.read_csv(shape_catalog_path).fillna("")
    sku_by_profile = sku_catalog.set_index("target_profile").to_dict("index")

    bom_rows: list[dict[str, object]] = []
    converted_rows: set[int] = set()
    shape_rows_seen: set[int] = set()
    covered_shapes: set[str] = set()
    unresolved_shapes: set[str] = set()

    cloud = _mode(working["cloud"], "unknown")
    if cloud == "aws":
        (
            compute_rows,
            compute_matched,
            compute_seen,
            compute_covered,
            compute_unresolved,
        ) = _build_aws_compute_rows(working, shape_catalog, sku_by_profile)
        bom_rows.extend(compute_rows)
        converted_rows.update(compute_matched)
        shape_rows_seen.update(compute_seen)
        covered_shapes.update(compute_covered)
        unresolved_shapes.update(compute_unresolved)

        for builder in (
            _build_aws_block_storage_rows,
            _build_aws_object_storage_rows,
            _build_aws_file_storage_rows,
            _build_aws_data_transfer_rows,
            _build_aws_load_balancer_rows,
        ):
            rows, matched = builder(working, sku_by_profile)
            bom_rows.extend(rows)
            converted_rows.update(matched)

    bom_df = _consolidate_bom_rows(bom_rows)
    review_df = _build_review_df(
        working,
        converted_rows=converted_rows,
        unresolved_shapes=unresolved_shapes,
    )
    summary_df = _build_summary_df(
        working,
        bom_df,
        review_df,
        converted_rows,
        shape_rows_seen,
        covered_shapes,
        unresolved_shapes,
    )
    return OciBomArtifacts(bom_df=bom_df, review_df=review_df, summary_df=summary_df)


def _build_aws_compute_rows(
    df: pd.DataFrame,
    shape_catalog: pd.DataFrame,
    sku_by_profile: dict[str, dict[str, str]],
) -> tuple[list[dict[str, object]], set[int], set[int], set[str], set[str]]:
    usage = df["usage_type"].str.lower()
    mask = (df["product_code"].str.lower() == "amazonec2") & usage.str.contains(
        r"(?:boxusage|spotusage|instanceusage)(?:iooptimized)?:", regex=True, na=False
    )
    compute = df.loc[mask].copy()
    if compute.empty:
        return [], set(), set(), set(), set()

    compute["shape"] = compute["usage_type"].map(_extract_aws_shape)
    catalog = shape_catalog[shape_catalog["cloud"].astype(str).str.lower() == "aws"].copy()
    catalog["shape"] = catalog["shape"].astype(str).str.lower()
    compute = compute.merge(catalog, how="left", on="shape", suffixes=("", "_catalog"))

    rows: list[dict[str, object]] = []
    seen = set(compute["_source_row"].astype(int).tolist())
    matched: set[int] = set()
    covered: set[str] = set()
    unresolved: set[str] = set()

    for shape, group in compute.groupby("shape", dropna=False, sort=False):
        shape = str(shape or "").strip()
        if not shape or group["target_profile"].fillna("").astype(str).str.strip().eq("").all():
            if shape:
                unresolved.add(shape)
            continue

        covered.add(shape)
        matched.update(group["_source_row"].astype(int).tolist())
        profile = str(group["target_profile"].iloc[0]).strip()
        hours = float(group["usage_quantity"].sum())
        vcpu = float(pd.to_numeric(group["vcpu"], errors="coerce").fillna(0.0).iloc[0])
        memory_gb = float(
            pd.to_numeric(group["memory_gb"], errors="coerce").fillna(0.0).iloc[0]
        )
        gpu_count = float(
            pd.to_numeric(group["gpu_count"], errors="coerce").fillna(0.0).iloc[0]
        )
        confidence = str(group["confidence"].iloc[0]).strip() or "Medium"
        common = _source_context(group, source_shapes=shape, source_quantity=hours)
        common["source_unit"] = "Hours"

        if profile == "compute_e5":
            rows.append(
                _sku_row(
                    sku_by_profile,
                    "compute_e5_ocpu",
                    quantity=hours * vcpu / 2.0,
                    unit="OCPU-hour",
                    confidence=confidence,
                    assumption=(
                        "Capacity-equivalent x86 sizing: 1 OCI OCPU = 2 source vCPUs; "
                        "no utilization-based rightsizing applied."
                    ),
                    **common,
                )
            )
            rows.append(
                _sku_row(
                    sku_by_profile,
                    "compute_e5_memory",
                    quantity=hours * memory_gb,
                    unit="GB-hour",
                    confidence=confidence,
                    assumption="Source provisioned memory converted without utilization-based rightsizing.",
                    **common,
                )
            )
        elif profile == "compute_a1":
            rows.append(
                _sku_row(
                    sku_by_profile,
                    "compute_a1_ocpu",
                    quantity=hours * vcpu,
                    unit="OCPU-hour",
                    confidence=confidence,
                    assumption=(
                        "Arm-to-Arm capacity equivalent: one OCI Ampere A1 OCPU per source vCPU; "
                        "application compatibility must be validated."
                    ),
                    **common,
                )
            )
            rows.append(
                _sku_row(
                    sku_by_profile,
                    "compute_a1_memory",
                    quantity=hours * memory_gb,
                    unit="GB-hour",
                    confidence=confidence,
                    assumption="Source provisioned memory converted without utilization-based rightsizing.",
                    **common,
                )
            )
        elif profile in {"gpu_l40s", "gpu_a10"}:
            gpu_note = (
                "Direct L40S GPU-hour equivalence; CPU and memory require final OCI shape validation."
                if profile == "gpu_l40s"
                else "Architectural substitution from NVIDIA T4 to OCI A10; performance and software compatibility validation required."
            )
            rows.append(
                _sku_row(
                    sku_by_profile,
                    profile,
                    quantity=hours * gpu_count,
                    unit="GPU-hour",
                    confidence=confidence,
                    assumption=gpu_note,
                    **common,
                )
            )

    return rows, matched, seen, covered, unresolved


def _build_aws_block_storage_rows(
    df: pd.DataFrame, sku_by_profile: dict[str, dict[str, str]]
) -> tuple[list[dict[str, object]], set[int]]:
    mask = (df["product_code"].str.lower() == "amazonec2") & df["usage_type"].str.contains(
        r"EBS:VolumeUsage\.(?:gp2|gp3)$", case=False, regex=True, na=False
    )
    rows: list[dict[str, object]] = []
    matched: set[int] = set()
    for _, group in df.loc[mask].groupby("region", dropna=False):
        quantity = float(group["usage_quantity"].sum())
        matched.update(group["_source_row"].astype(int).tolist())
        common = _source_context(group, source_quantity=quantity)
        rows.append(
            _sku_row(
                sku_by_profile,
                "block_storage",
                quantity=quantity,
                unit="GB-month",
                confidence="High",
                assumption="AWS gp2/gp3 provisioned storage converted at equal GB-month capacity.",
                **common,
            )
        )
        rows.append(
            _sku_row(
                sku_by_profile,
                "block_vpu",
                quantity=quantity * 10.0,
                unit="VPU-GB-month",
                confidence="Medium",
                assumption=(
                    "OCI Balanced performance assumption of 10 VPU per GB; validate IOPS and throughput requirements."
                ),
                **common,
            )
        )
    return rows, matched


def _build_aws_object_storage_rows(
    df: pd.DataFrame, sku_by_profile: dict[str, dict[str, str]]
) -> tuple[list[dict[str, object]], set[int]]:
    is_s3 = df["product_code"].str.lower() == "amazons3"
    rules = [
        (r"(?:^|-)TimedStorage-ByteHrs$", "object_standard", "GB-month", 1.0, "High", "Standard S3 storage converted at equal GB-month capacity."),
        (r"(?:^|-)TimedStorage-GlacierByteHrs$", "object_archive", "GB-month", 1.0, "Medium", "S3 Glacier storage converted to OCI Archive Storage; retrieval and minimum-duration behavior must be validated."),
        (r"(?:^|-)Requests-Tier[12]$", "object_requests", "10K requests", 0.0001, "Medium", "AWS S3 Tier 1/2 requests consolidated into OCI 10,000-request billing units; operation mix must be validated."),
    ]
    rows: list[dict[str, object]] = []
    matched: set[int] = set()
    for pattern, profile, unit, factor, confidence, assumption in rules:
        mask = is_s3 & df["usage_type"].str.contains(pattern, case=False, regex=True, na=False)
        for _, group in df.loc[mask].groupby("region", dropna=False):
            source_quantity = float(group["usage_quantity"].sum())
            matched.update(group["_source_row"].astype(int).tolist())
            rows.append(
                _sku_row(
                    sku_by_profile,
                    profile,
                    quantity=source_quantity * factor,
                    unit=unit,
                    confidence=confidence,
                    assumption=assumption,
                    **_source_context(group, source_quantity=source_quantity),
                )
            )
    return rows, matched


def _build_aws_file_storage_rows(
    df: pd.DataFrame, sku_by_profile: dict[str, dict[str, str]]
) -> tuple[list[dict[str, object]], set[int]]:
    mask = (df["product_code"].str.lower() == "amazonefs") & df["usage_type"].str.contains(
        r"(?:^|-)TimedStorage-ByteHrs$", case=False, regex=True, na=False
    )
    rows: list[dict[str, object]] = []
    matched: set[int] = set()
    for _, group in df.loc[mask].groupby("region", dropna=False):
        quantity = float(group["usage_quantity"].sum())
        matched.update(group["_source_row"].astype(int).tolist())
        rows.append(
            _sku_row(
                sku_by_profile,
                "file_storage",
                quantity=quantity,
                unit="GB-month",
                confidence="Medium",
                assumption="AWS EFS storage converted at equal GB-month capacity; performance tier must be validated.",
                **_source_context(group, source_quantity=quantity),
            )
        )
    return rows, matched


def _build_aws_data_transfer_rows(
    df: pd.DataFrame, sku_by_profile: dict[str, dict[str, str]]
) -> tuple[list[dict[str, object]], set[int]]:
    mask = (df["product_code"].str.lower() == "awsdatatransfer") & df["usage_type"].str.contains(
        r"(?:^|-)DataTransfer-Out-Bytes$", case=False, regex=True, na=False
    )
    rows: list[dict[str, object]] = []
    matched: set[int] = set()
    for region, group in df.loc[mask].groupby("region", dropna=False):
        region = str(region or "").strip()
        quantity = float(group["usage_quantity"].sum())
        matched.update(group["_source_row"].astype(int).tolist())
        south_apac = region == "sa-east-1" or region.startswith(("ap-", "me-", "af-"))
        profile = "egress_apac_south_america" if south_apac else "egress_na_europe"
        confidence = "High" if region else "Medium"
        assumption = (
            "Outbound internet data converted at equal GB from a South America/APAC origin."
            if south_apac
            else "Outbound internet data converted at equal GB; blank source region is assumed North America/Europe and requires validation."
        )
        rows.append(
            _sku_row(
                sku_by_profile,
                profile,
                quantity=quantity,
                unit="GB",
                confidence=confidence,
                assumption=assumption,
                **_source_context(group, source_quantity=quantity),
            )
        )
    return rows, matched


def _build_aws_load_balancer_rows(
    df: pd.DataFrame, sku_by_profile: dict[str, dict[str, str]]
) -> tuple[list[dict[str, object]], set[int]]:
    mask = (df["product_code"].str.lower() == "awselb") & df["usage_type"].str.contains(
        r"(?:^|-)LoadBalancerUsage$", case=False, regex=True, na=False
    )
    rows: list[dict[str, object]] = []
    matched: set[int] = set()
    for _, group in df.loc[mask].groupby("region", dropna=False):
        quantity = float(group["usage_quantity"].sum())
        matched.update(group["_source_row"].astype(int).tolist())
        rows.append(
            _sku_row(
                sku_by_profile,
                "load_balancer_base",
                quantity=quantity,
                unit="load-balancer-hour",
                confidence="Medium",
                assumption="AWS load balancer hours converted to OCI flexible load balancer base hours; bandwidth requires separate validation.",
                **_source_context(group, source_quantity=quantity),
            )
        )
    return rows, matched


def _sku_row(
    sku_by_profile: dict[str, dict[str, str]],
    profile: str,
    *,
    quantity: float,
    unit: str,
    confidence: str,
    assumption: str,
    **source: object,
) -> dict[str, object]:
    sku = sku_by_profile[profile]
    return {
        "scenario": SCENARIO,
        "oci_part_number": sku["part_number"],
        "oci_sku": sku["display_name"],
        "oci_service_category": sku["service_category"],
        "target_profile": profile,
        "quantity": float(quantity),
        "unit": unit,
        "confidence": confidence,
        "assumption": assumption,
        "catalog_last_updated": sku["catalog_last_updated"],
        "catalog_url": sku["catalog_url"],
        **source,
    }


def _source_context(
    group: pd.DataFrame,
    *,
    source_shapes: str = "",
    source_quantity: float,
) -> dict[str, object]:
    return {
        "source_cloud": _mode(group["cloud"], "unknown"),
        "source_service": _join_unique(group["service_name_original"]),
        "source_skus": _join_unique(group["usage_type"], limit=12),
        "source_shapes": source_shapes,
        "source_quantity": source_quantity,
        "source_unit": _mode(group["usage_unit"], "source units"),
        "source_cost": float(group["cost"].sum()),
        "currency": _mode(group["currency"], "USD"),
        "region": _mode(group["region"], "unspecified"),
    }


def _consolidate_bom_rows(rows: list[dict[str, object]]) -> pd.DataFrame:
    if not rows:
        return pd.DataFrame(columns=BOM_COLUMNS)
    df = pd.DataFrame(rows)
    numeric = ["quantity", "source_quantity", "source_cost"]
    keys = [
        "scenario",
        "oci_part_number",
        "oci_sku",
        "oci_service_category",
        "target_profile",
        "unit",
        "source_cloud",
        "source_unit",
        "currency",
        "region",
        "confidence",
        "assumption",
        "catalog_last_updated",
        "catalog_url",
    ]
    grouped = (
        df.groupby(keys, as_index=False, dropna=False)
        .agg(
            quantity=("quantity", "sum"),
            source_quantity=("source_quantity", "sum"),
            source_cost=("source_cost", "sum"),
            source_service=("source_service", _join_unique),
            source_skus=("source_skus", _join_unique),
            source_shapes=("source_shapes", _join_unique),
        )
        .sort_values(["oci_service_category", "oci_part_number", "region"])
    )
    for column in numeric:
        grouped[column] = pd.to_numeric(grouped[column], errors="coerce").fillna(0.0).round(4)
    return grouped[BOM_COLUMNS].reset_index(drop=True)


def _build_review_df(
    df: pd.DataFrame,
    *,
    converted_rows: set[int],
    unresolved_shapes: set[str],
) -> pd.DataFrame:
    pending = df.loc[~df["_source_row"].isin(converted_rows) & (df["cost"].abs() > 0.000001)].copy()
    if pending.empty and not unresolved_shapes:
        return pd.DataFrame(columns=REVIEW_COLUMNS)

    if not pending.empty:
        pending["review_reason"] = pending.apply(_review_reason, axis=1)
        grouped = (
            pending.groupby(
                [
                    "cloud",
                    "service_name_original",
                    "usage_type",
                    "region",
                    "usage_unit",
                    "currency",
                    "review_reason",
                ],
                as_index=False,
                dropna=False,
            )
            .agg(source_quantity=("usage_quantity", "sum"), source_cost=("cost", "sum"))
            .sort_values("source_cost", key=lambda value: value.abs(), ascending=False)
            .rename(
                columns={
                    "cloud": "source_cloud",
                    "service_name_original": "source_service",
                    "usage_type": "source_sku",
                    "usage_unit": "source_unit",
                }
            )
        )
    else:
        grouped = pd.DataFrame(columns=REVIEW_COLUMNS)

    existing_skus = set(grouped.get("source_sku", pd.Series(dtype=str)).astype(str))
    unresolved_records = []
    for shape in sorted(unresolved_shapes):
        if shape in existing_skus:
            continue
        unresolved_records.append(
            {
                "source_cloud": "aws",
                "source_service": "Amazon Elastic Compute Cloud",
                "source_sku": shape,
                "region": "",
                "source_quantity": 0.0,
                "source_unit": "Hours",
                "source_cost": 0.0,
                "currency": _mode(df["currency"], "USD"),
                "review_reason": "Shape not present in the versioned capacity catalog; add verified vCPU and memory before sizing.",
            }
        )
    if unresolved_records:
        grouped = pd.concat([grouped, pd.DataFrame(unresolved_records)], ignore_index=True)
    return grouped[REVIEW_COLUMNS]


def _review_reason(row: pd.Series) -> str:
    product = str(row.get("product_code", "")).lower()
    usage = str(row.get("usage_type", ""))
    if re.search(r"HeavyUsage|SavingsPlan|Reserved", usage, re.IGNORECASE):
        return "Contract or reservation charge excluded from capacity quantities to avoid double counting usage hours."
    if product in {"amazonrds", "amazonneptune", "amazondynamodb", "amazones"}:
        return "Managed data service requires engine, instance class, storage, throughput, SLA and HA topology before selecting OCI SKUs."
    if "natgateway" in usage.lower():
        return "NAT Gateway has no safe one-to-one quantity conversion from the aggregated billing line."
    if "lcuusage" in usage.lower():
        return "AWS LCU does not directly convert to OCI Mbps per hour; validate observed bandwidth and connections."
    if product == "amazons3" and "timedstorage" in usage.lower():
        return "Storage tier has different retrieval or lifecycle semantics; validate target OCI storage tier."
    if product == "amazonec2" and re.search(r"(?:BoxUsage|SpotUsage|InstanceUsage)", usage, re.IGNORECASE):
        return "Compute shape could not be resolved in the versioned capacity catalog."
    return "No auditable quantity conversion rule is available for this aggregated billing line."


def _build_summary_df(
    df: pd.DataFrame,
    bom_df: pd.DataFrame,
    review_df: pd.DataFrame,
    converted_rows: set[int],
    shape_rows_seen: set[int],
    covered_shapes: set[str],
    unresolved_shapes: set[str],
) -> pd.DataFrame:
    total_abs_cost = float(df["cost"].abs().sum())
    converted_abs_cost = float(
        df.loc[df["_source_row"].isin(converted_rows), "cost"].abs().sum()
    )
    shape_hours = float(df.loc[df["_source_row"].isin(shape_rows_seen), "usage_quantity"].sum())
    return pd.DataFrame(
        [
            {"metric": "scenario", "value": SCENARIO, "note": "Technical capacity equivalence; no OCI prices applied."},
            {"metric": "oci_part_numbers", "value": int(bom_df["oci_part_number"].nunique()) if not bom_df.empty else 0, "note": "Unique official OCI catalog part numbers generated."},
            {"metric": "bom_lines", "value": len(bom_df), "note": "Consolidated by part number, region, unit and assumption."},
            {"metric": "review_lines", "value": len(review_df), "note": "Source billing groups requiring manual sizing or validation."},
            {"metric": "converted_cost_coverage_pct", "value": round((converted_abs_cost / total_abs_cost * 100.0) if total_abs_cost else 0.0, 2), "note": "Indicative only; contract charges may carry cost while usage lines carry quantity."},
            {"metric": "compute_shape_hours", "value": round(shape_hours, 4), "note": "Detected EC2 BoxUsage/SpotUsage/InstanceUsage hours."},
            {"metric": "covered_compute_shapes", "value": len(covered_shapes), "note": ", ".join(sorted(covered_shapes))},
            {"metric": "unresolved_compute_shapes", "value": len(unresolved_shapes), "note": ", ".join(sorted(unresolved_shapes))},
            {"metric": "catalog_last_updated", "value": _mode(bom_df.get("catalog_last_updated", pd.Series(dtype=str)), "N/A"), "note": "Version captured from the OCI public price catalog."},
        ]
    )


def _extract_aws_shape(value: str) -> str:
    match = re.search(
        r"(?:BoxUsage|SpotUsage|InstanceUsage(?:IOOptimized)?):(?:db\.)?([A-Za-z0-9.]+)",
        str(value),
        re.IGNORECASE,
    )
    return match.group(1).lower() if match else ""


def _join_unique(values: pd.Series, limit: int = 20) -> str:
    tokens: list[str] = []
    for value in values.fillna("").astype(str):
        for token in value.split("; "):
            token = token.strip()
            if token and token not in tokens:
                tokens.append(token)
            if len(tokens) >= limit:
                break
        if len(tokens) >= limit:
            break
    return "; ".join(tokens)


def _mode(values: pd.Series, default: str) -> str:
    cleaned = values.fillna("").astype(str).str.strip()
    cleaned = cleaned[cleaned != ""]
    if cleaned.empty:
        return default
    return str(cleaned.mode().iloc[0])


def _empty_artifacts() -> OciBomArtifacts:
    return OciBomArtifacts(
        bom_df=pd.DataFrame(columns=BOM_COLUMNS),
        review_df=pd.DataFrame(columns=REVIEW_COLUMNS),
        summary_df=pd.DataFrame(columns=["metric", "value", "note"]),
    )
