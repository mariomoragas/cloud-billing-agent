from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import pandas as pd

APP_ROOT = Path(__file__).resolve().parent
DEFAULT_RULES_PATH = APP_ROOT / "mappings" / "cloud_shape_processors.csv"
DEFAULT_BENCHMARKS_PATH = APP_ROOT / "mappings" / "processor_benchmarks.csv"

DETAIL_COLUMNS = [
    "cloud",
    "source_shape",
    "shape_family",
    "source_processor_model",
    "architecture",
    "mapping_confidence",
    "source_cost",
    "positive_cost_weight",
    "usage_quantity",
    "record_count",
    "cpu_mark",
    "single_thread_mark",
    "physical_cores",
    "benchmark_samples",
    "benchmark_date",
    "benchmark_retrieved_at",
    "benchmark_confidence",
    "comparison_status",
    "shape_source_url",
    "benchmark_url",
    "notes",
]

COMPARISON_COLUMNS = [
    "source_cloud",
    "source_shape_families",
    "source_processor_models",
    "target_shape_family",
    "target_processor_model",
    "target_cpu_mark",
    "target_cpu_count",
    "target_physical_cores",
    "target_single_thread_mark",
    "source_weighted_single_thread_mark",
    "estimated_gain_pct",
    "source_benchmark_coverage_pct",
    "confidence",
    "target_benchmark_samples",
    "benchmark_retrieved_at",
    "methodology",
    "benchmark_url",
]


@dataclass(frozen=True)
class ProcessorAnalysisArtifacts:
    detail_df: pd.DataFrame
    comparison_df: pd.DataFrame


def build_processor_analysis(
    raw_df: pd.DataFrame,
    *,
    rules_path: Path = DEFAULT_RULES_PATH,
    benchmarks_path: Path = DEFAULT_BENCHMARKS_PATH,
) -> ProcessorAnalysisArtifacts:
    if raw_df.empty:
        return _empty_artifacts()

    rules = _load_rules(rules_path)
    benchmarks = _load_benchmarks(benchmarks_path)
    working = raw_df.reset_index(drop=True).copy()
    working["cloud"] = working.get("cloud", "").fillna("").astype(str).str.lower()
    working["_text"] = _combined_text(working)
    working["source_shape"] = _extract_shapes(working)
    working["shape_family"] = ""
    working["source_processor_model"] = ""
    working["architecture"] = ""
    working["mapping_confidence"] = 0.0
    working["shape_source_url"] = ""
    working["notes"] = ""

    for _, rule in rules.iterrows():
        available = working["source_processor_model"].eq("")
        cloud_match = working["cloud"].eq(rule["cloud"])
        pattern_match = working["_text"].str.contains(
            rule["shape_pattern"], case=False, regex=True, na=False
        )
        mask = available & cloud_match & pattern_match
        if not mask.any():
            continue
        for column in (
            "shape_family",
            "processor_model",
            "architecture",
            "mapping_confidence",
            "source_url",
            "notes",
        ):
            target = {
                "processor_model": "source_processor_model",
                "source_url": "shape_source_url",
            }.get(column, column)
            working.loc[mask, target] = rule[column]
        working.loc[mask & working["source_shape"].eq(""), "source_shape"] = rule[
            "shape_family"
        ]

    candidate = working["source_shape"].ne("") | working["source_processor_model"].ne("")
    working = working.loc[candidate].copy()
    if working.empty:
        return _empty_artifacts()

    unresolved = working["source_processor_model"].eq("")
    working.loc[unresolved, "source_processor_model"] = "UNRESOLVED"
    working.loc[unresolved, "shape_family"] = working.loc[unresolved, "source_shape"]
    working.loc[unresolved, "notes"] = "Shape detected without a unique processor rule."
    working["cost"] = pd.to_numeric(working.get("cost", 0.0), errors="coerce").fillna(0.0)
    working["usage_quantity"] = pd.to_numeric(
        working.get("usage_quantity", 0.0), errors="coerce"
    ).fillna(0.0)
    working["positive_cost_weight"] = working["cost"].clip(lower=0.0)

    detail = (
        working.groupby(
            [
                "cloud",
                "source_shape",
                "shape_family",
                "source_processor_model",
                "architecture",
                "shape_source_url",
                "notes",
            ],
            dropna=False,
            as_index=False,
        )
        .agg(
            mapping_confidence=("mapping_confidence", "min"),
            source_cost=("cost", "sum"),
            positive_cost_weight=("positive_cost_weight", "sum"),
            usage_quantity=("usage_quantity", "sum"),
            record_count=("source_shape", "size"),
        )
        .sort_values("source_cost", ascending=False)
        .reset_index(drop=True)
    )

    source_benchmarks = benchmarks[benchmarks["platform"].eq("source")].copy()
    source_benchmarks = source_benchmarks.rename(
        columns={
            "processor_model": "source_processor_model",
            "samples": "benchmark_samples",
            "retrieved_at": "benchmark_retrieved_at",
        }
    )
    detail = detail.merge(
        source_benchmarks[
            [
                "source_processor_model",
                "cpu_mark",
                "single_thread_mark",
                "physical_cores",
                "benchmark_samples",
                "benchmark_date",
                "benchmark_retrieved_at",
                "benchmark_url",
            ]
        ],
        on="source_processor_model",
        how="left",
    )
    detail["benchmark_confidence"] = detail["benchmark_samples"].apply(
        _benchmark_confidence
    )
    detail["comparison_status"] = detail.apply(_comparison_status, axis=1)
    detail = detail.reindex(columns=DETAIL_COLUMNS)

    comparison = _build_comparison(detail, benchmarks)
    return ProcessorAnalysisArtifacts(detail_df=detail, comparison_df=comparison)


def _load_rules(path: Path) -> pd.DataFrame:
    rules = pd.read_csv(path).fillna("")
    required = {
        "priority",
        "cloud",
        "shape_family",
        "shape_pattern",
        "processor_model",
        "architecture",
        "mapping_confidence",
        "source_url",
        "notes",
    }
    missing = required.difference(rules.columns)
    if missing:
        raise ValueError("Processor rules missing columns: " + ", ".join(sorted(missing)))
    rules["priority"] = pd.to_numeric(rules["priority"], errors="raise")
    rules["mapping_confidence"] = pd.to_numeric(
        rules["mapping_confidence"], errors="raise"
    )
    rules["cloud"] = rules["cloud"].astype(str).str.lower().str.strip()
    return rules.sort_values("priority")


def _load_benchmarks(path: Path) -> pd.DataFrame:
    benchmarks = pd.read_csv(path).fillna("")
    required = {
        "platform",
        "shape_family",
        "processor_model",
        "cpu_mark",
        "single_thread_mark",
        "physical_cores",
        "cpu_count",
        "samples",
        "retrieved_at",
        "benchmark_url",
    }
    missing = required.difference(benchmarks.columns)
    if missing:
        raise ValueError("Processor benchmarks missing columns: " + ", ".join(sorted(missing)))
    for column in (
        "cpu_mark",
        "single_thread_mark",
        "physical_cores",
        "cpu_count",
        "samples",
    ):
        benchmarks[column] = pd.to_numeric(benchmarks[column], errors="coerce")
    return benchmarks


def _combined_text(df: pd.DataFrame) -> pd.Series:
    columns = [
        "sku",
        "item_description",
        "instance_name",
        "resource_name",
        "subcategory",
        "service_name_original",
        "product_code",
    ]
    combined = pd.Series("", index=df.index, dtype="object")
    for column in columns:
        if column not in df.columns:
            continue
        combined = combined.str.cat(
            df[column].fillna("").astype(str).str.lower(), sep=" | "
        )
    return combined


def _extract_shapes(df: pd.DataFrame) -> pd.Series:
    result = pd.Series("", index=df.index, dtype="object")
    specs = {
        "aws": r"\b((?:[a-z][a-z0-9]*\d[a-z0-9]*)\.(?:metal|nano|micro|small|medium|large|xlarge|\d+xlarge))\b",
        "azure": r"\b((?:standard[_-])?[a-z][a-z0-9_-]*v\d+)\b",
        "gcp": r"\b(t2d|n2d|n2|c2|c3|e2)(?:[- _]|\b)",
    }
    for cloud, pattern in specs.items():
        mask = df["cloud"].eq(cloud)
        if not mask.any():
            continue
        extracted = df.loc[mask, "_text"].str.extract(pattern, expand=False).fillna("")
        result.loc[mask] = extracted
    return result


def _comparison_status(row: pd.Series) -> str:
    processor = str(row.get("source_processor_model", ""))
    architecture = str(row.get("architecture", ""))
    score = pd.to_numeric(row.get("single_thread_mark"), errors="coerce")
    if processor == "UNRESOLVED":
        return "Processor model requires validation"
    if architecture == "arm64":
        return "Architecture differs; no x86 performance estimate"
    if processor.startswith("GPU host"):
        return "GPU workload; CPU-only comparison excluded"
    if pd.isna(score) or float(score) <= 0:
        return "No comparable PassMark record in catalog"
    return "Comparable single-thread benchmark"


def _benchmark_confidence(samples: object) -> str:
    numeric = pd.to_numeric(samples, errors="coerce")
    if pd.isna(numeric) or float(numeric) < 3:
        return "Low"
    if float(numeric) < 10:
        return "Medium"
    return "High"


def _build_comparison(detail: pd.DataFrame, benchmarks: pd.DataFrame) -> pd.DataFrame:
    valid = detail[
        detail["single_thread_mark"].notna()
        & detail["single_thread_mark"].gt(0)
        & detail["architecture"].eq("x86_64")
    ].copy()
    if valid.empty:
        return pd.DataFrame(columns=COMPARISON_COLUMNS)

    candidate_cost = float(detail["positive_cost_weight"].sum())
    valid_cost = float(valid["positive_cost_weight"].sum())
    if candidate_cost > 0 and valid_cost <= 0:
        return pd.DataFrame(columns=COMPARISON_COLUMNS)

    weights = valid["positive_cost_weight"].clip(lower=0.0)
    if float(weights.sum()) <= 0:
        weights = valid["usage_quantity"].clip(lower=0.0)
    if float(weights.sum()) <= 0:
        weights = pd.Series(1.0, index=valid.index)

    source_score = float((valid["single_thread_mark"] * weights).sum() / weights.sum())
    source_cloud = ", ".join(sorted(valid["cloud"].dropna().astype(str).unique()))
    source_shape_families = ", ".join(
        sorted(valid["shape_family"].dropna().astype(str).unique())
    )
    source_processor_models = ", ".join(
        sorted(valid["source_processor_model"].dropna().astype(str).unique())
    )
    coverage = (valid_cost / candidate_cost * 100.0) if candidate_cost > 0 else 100.0
    mapping_floor = float(valid["mapping_confidence"].min())
    source_sample_floor = int(
        pd.to_numeric(valid["benchmark_samples"], errors="coerce").fillna(0).min()
    )

    rows: list[dict[str, object]] = []
    targets = benchmarks[benchmarks["platform"].eq("oci")].copy()
    target_order = pd.Categorical(targets["shape_family"], ["X9", "E5", "E6"])
    targets = targets.assign(_order=target_order).sort_values("_order")
    for _, target in targets.iterrows():
        target_score = float(target["single_thread_mark"])
        target_samples = int(target["samples"])
        confidence = _overall_confidence(
            coverage_pct=coverage,
            mapping_floor=mapping_floor,
            sample_floor=min(source_sample_floor, target_samples),
        )
        rows.append(
            {
                "source_cloud": source_cloud,
                "source_shape_families": source_shape_families,
                "source_processor_models": source_processor_models,
                "target_shape_family": target["shape_family"],
                "target_processor_model": target["processor_model"],
                "target_cpu_mark": float(target["cpu_mark"]),
                "target_cpu_count": int(target["cpu_count"]),
                "target_physical_cores": int(target["physical_cores"]),
                "target_single_thread_mark": target_score,
                "source_weighted_single_thread_mark": source_score,
                "estimated_gain_pct": (target_score / source_score - 1.0) * 100.0,
                "source_benchmark_coverage_pct": coverage,
                "confidence": confidence,
                "target_benchmark_samples": target_samples,
                "benchmark_retrieved_at": target["retrieved_at"],
                "methodology": (
                    "PassMark CPU Mark Single Thread; source baseline weighted by positive "
                    "billing cost for identified compute shapes. Estimate, not workload benchmark."
                ),
                "benchmark_url": target["benchmark_url"],
            }
        )
    return pd.DataFrame(rows, columns=COMPARISON_COLUMNS)


def _overall_confidence(
    *, coverage_pct: float, mapping_floor: float, sample_floor: int
) -> str:
    if coverage_pct >= 80 and mapping_floor >= 0.85 and sample_floor >= 10:
        return "High"
    if coverage_pct >= 50 and mapping_floor >= 0.70 and sample_floor >= 3:
        return "Medium"
    return "Low"


def _empty_artifacts() -> ProcessorAnalysisArtifacts:
    return ProcessorAnalysisArtifacts(
        detail_df=pd.DataFrame(columns=DETAIL_COLUMNS),
        comparison_df=pd.DataFrame(columns=COMPARISON_COLUMNS),
    )
