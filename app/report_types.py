from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd


@dataclass
class ParserResult:
    dataframe: pd.DataFrame
    data_quality: pd.DataFrame = field(default_factory=pd.DataFrame)


@dataclass
class ProcessResult:
    output_path: Path
    presentation_path: Path | None
    raw_df: pd.DataFrame
    service_summary_df: pd.DataFrame
    region_summary_df: pd.DataFrame
    oci_mapping_df: pd.DataFrame
    data_quality_df: pd.DataFrame = field(default_factory=pd.DataFrame)
    processor_detail_df: pd.DataFrame = field(default_factory=pd.DataFrame)
    processor_comparison_df: pd.DataFrame = field(default_factory=pd.DataFrame)
    oci_bom_df: pd.DataFrame = field(default_factory=pd.DataFrame)
    oci_bom_review_df: pd.DataFrame = field(default_factory=pd.DataFrame)
    oci_bom_summary_df: pd.DataFrame = field(default_factory=pd.DataFrame)
    finops_baseline_df: pd.DataFrame = field(default_factory=pd.DataFrame)
    finops_categories_df: pd.DataFrame = field(default_factory=pd.DataFrame)
    effective_rates_df: pd.DataFrame = field(default_factory=pd.DataFrame)
    analysis_coverage_df: pd.DataFrame = field(default_factory=pd.DataFrame)
    oci_priced_bom_df: pd.DataFrame = field(default_factory=pd.DataFrame)
    savings_by_service_df: pd.DataFrame = field(default_factory=pd.DataFrame)
    business_case_df: pd.DataFrame = field(default_factory=pd.DataFrame)
    savings_waterfall_df: pd.DataFrame = field(default_factory=pd.DataFrame)
    sensitivity_df: pd.DataFrame = field(default_factory=pd.DataFrame)
    llm_report_df: pd.DataFrame = field(default_factory=pd.DataFrame)
    llm_migration_df: pd.DataFrame = field(default_factory=pd.DataFrame)
    llm_recommendations_df: pd.DataFrame = field(default_factory=pd.DataFrame)
    llm_confidence_df: pd.DataFrame = field(default_factory=pd.DataFrame)
