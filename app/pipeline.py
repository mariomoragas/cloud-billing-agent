from __future__ import annotations

from pathlib import Path

import pandas as pd

from app.aggregator import (
    build_aws_enterprise_summaries,
    summarize_by_column,
    summarize_by_region,
    summarize_by_service,
)
from app.excel_writer import write_billing_report
from app.input_validation import validate_billing_input_file
from app.llm_report import build_llm_report_artifacts
from app.normalizer import load_and_normalize_csv
from app.oci_mapper import build_oci_mapping, load_mapping_table
from app.parsers.aws_invoice import load_aws_invoice_csv
from app.parsers.azure_cost_csv import load_azure_cost_csv
from app.parsers.gcp_cost_table import load_gcp_cost_table_csv
from app.report_types import ProcessResult


def process_billing_file(
    *,
    input_path: Path,
    output_path: Path,
    presentation_path: Path | None = None,
    file_format: str,
    cloud: str,
    mapping_path: Path,
    company_name: str = "",
    project_name: str = "",
    llm_model: str = "gemini-2.5-flash",
) -> Path:
    data_quality_df = pd.DataFrame()
    validate_billing_input_file(input_path, file_format)

    if file_format == "aws-invoice":
        invoice_result = load_aws_invoice_csv(input_path)
        raw_df = invoice_result.dataframe
        data_quality_df = invoice_result.data_quality
    elif file_format == "aws-billing-pdf":
        from app.parsers.aws_billing_pdf import load_aws_billing_pdf

        pdf_result = load_aws_billing_pdf(input_path)
        raw_df = pdf_result.dataframe
        data_quality_df = pdf_result.data_quality
    elif file_format == "gcp-cost-table":
        gcp_result = load_gcp_cost_table_csv(input_path)
        raw_df = gcp_result.dataframe
        data_quality_df = gcp_result.data_quality
    elif file_format == "azure-cost-csv":
        azure_result = load_azure_cost_csv(input_path)
        raw_df = azure_result.dataframe
        data_quality_df = azure_result.data_quality
    else:
        raw_df = load_and_normalize_csv(input_path, default_cloud=cloud)

    service_summary_df = summarize_by_service(raw_df)
    if file_format == "aws-billing-pdf":
        service_summary_df = _apply_pdf_chart_grouping(service_summary_df)
    region_summary_df = summarize_by_region(raw_df)
    mapping_df = load_mapping_table(mapping_path)
    oci_mapping_df = build_oci_mapping(service_summary_df, mapping_df)
    extra_summaries = (
        build_aws_enterprise_summaries(raw_df)
        if file_format in {"aws-invoice", "aws-billing-pdf"}
        else {}
    )
    if file_format == "gcp-cost-table":
        extra_summaries = {
            "gcp_project_name": summarize_by_column(
                raw_df,
                column="project_name",
                label="project_name",
                top_n=20,
            ),
            "gcp_sku": summarize_by_column(
                raw_df,
                column="sku",
                label="sku",
                top_n=20,
            ),
            "gcp_credit_type": summarize_by_column(
                raw_df,
                column="credit_type",
                label="credit_type",
                top_n=20,
            ),
        }
    llm_artifacts = build_llm_report_artifacts(
        raw_df=raw_df,
        service_summary_df=service_summary_df,
        region_summary_df=region_summary_df,
        oci_mapping_df=oci_mapping_df,
        source_name=input_path.name,
        llm_model=llm_model,
    )

    write_billing_report(
        output_path=output_path,
        raw_df=raw_df,
        service_summary_df=service_summary_df,
        region_summary_df=region_summary_df,
        oci_mapping_df=oci_mapping_df,
        extra_summaries=extra_summaries,
        data_quality_df=data_quality_df,
        llm_report_df=llm_artifacts.summary_df,
        llm_migration_df=llm_artifacts.migration_df,
        llm_recommendations_df=llm_artifacts.recommendations_df,
        llm_confidence_df=llm_artifacts.confidence_df,
    )
    if presentation_path is not None:
        from app.powerpoint_writer import write_powerpoint_report

        write_powerpoint_report(
            output_path=presentation_path,
            raw_df=raw_df,
            service_summary_df=service_summary_df,
            region_summary_df=region_summary_df,
            oci_mapping_df=oci_mapping_df,
            llm_report_df=llm_artifacts.summary_df,
            llm_migration_df=llm_artifacts.migration_df,
            llm_recommendations_df=llm_artifacts.recommendations_df,
            llm_confidence_df=llm_artifacts.confidence_df,
            report_name=input_path.stem,
            company_name=company_name,
            project_name=project_name,
        )
    return ProcessResult(
        output_path=output_path,
        presentation_path=presentation_path,
        raw_df=raw_df,
        service_summary_df=service_summary_df,
        region_summary_df=region_summary_df,
        oci_mapping_df=oci_mapping_df,
        data_quality_df=data_quality_df,
        llm_report_df=llm_artifacts.summary_df,
        llm_migration_df=llm_artifacts.migration_df,
        llm_recommendations_df=llm_artifacts.recommendations_df,
        llm_confidence_df=llm_artifacts.confidence_df,
    )


def _apply_pdf_chart_grouping(service_summary_df: pd.DataFrame) -> pd.DataFrame:
    summary = service_summary_df.copy()
    if "primary_product_code" in summary.columns:
        product_code = summary["primary_product_code"].fillna("").astype(str).str.strip()
    else:
        product_code = pd.Series([""] * len(summary), index=summary.index)

    service_name = summary["service_name_original"].fillna("").astype(str).str.strip()
    summary["chart_group_label"] = product_code.where(product_code != "", service_name)
    return summary
