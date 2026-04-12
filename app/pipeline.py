from __future__ import annotations

from pathlib import Path

import pandas as pd

from app.aggregator import (
    build_aws_enterprise_summaries,
    summarize_by_region,
    summarize_by_service,
)
from app.excel_writer import write_billing_report
from app.normalizer import load_and_normalize_csv
from app.oci_mapper import build_oci_mapping, load_mapping_table
from app.parsers.aws_invoice import load_aws_invoice_csv
from app.powerpoint_writer import write_powerpoint_report
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
) -> Path:
    data_quality_df = pd.DataFrame()

    if file_format == "aws-invoice":
        invoice_result = load_aws_invoice_csv(input_path)
        raw_df = invoice_result.dataframe
        data_quality_df = invoice_result.data_quality
    else:
        raw_df = load_and_normalize_csv(input_path, default_cloud=cloud)

    service_summary_df = summarize_by_service(raw_df)
    region_summary_df = summarize_by_region(raw_df)
    mapping_df = load_mapping_table(mapping_path)
    oci_mapping_df = build_oci_mapping(service_summary_df, mapping_df)
    extra_summaries = (
        build_aws_enterprise_summaries(raw_df)
        if file_format == "aws-invoice"
        else {}
    )

    write_billing_report(
        output_path=output_path,
        raw_df=raw_df,
        service_summary_df=service_summary_df,
        region_summary_df=region_summary_df,
        oci_mapping_df=oci_mapping_df,
        extra_summaries=extra_summaries,
        data_quality_df=data_quality_df,
    )
    if presentation_path is not None:
        write_powerpoint_report(
            output_path=presentation_path,
            raw_df=raw_df,
            service_summary_df=service_summary_df,
            region_summary_df=region_summary_df,
            oci_mapping_df=oci_mapping_df,
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
    )
