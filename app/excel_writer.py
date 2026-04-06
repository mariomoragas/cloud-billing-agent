from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Font

DEFAULT_FONT = "Verdana"


def write_billing_report(
    output_path: Path,
    raw_df: pd.DataFrame,
    service_summary_df: pd.DataFrame,
    region_summary_df: pd.DataFrame,
    oci_mapping_df: pd.DataFrame,
    extra_summaries: dict[str, pd.DataFrame] | None = None,
    data_quality_df: pd.DataFrame | None = None,
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        raw_df.to_excel(writer, sheet_name="Raw_Data", index=False)
        service_summary_df.to_excel(writer, sheet_name="Resumo_Servicos", index=False)
        region_summary_df.to_excel(writer, sheet_name="Resumo_Regioes", index=False)
        oci_mapping_df.to_excel(writer, sheet_name="Mapeamento_OCI", index=False)
        if data_quality_df is not None and not data_quality_df.empty:
            data_quality_df.to_excel(writer, sheet_name="Data_Quality", index=False)
        if extra_summaries:
            for sheet_name, summary_df in extra_summaries.items():
                if summary_df.empty:
                    continue
                safe_name = _sheet_name(sheet_name)
                summary_df.to_excel(writer, sheet_name=safe_name, index=False)
        _build_pending_sheet(writer, oci_mapping_df)

    workbook = load_workbook(output_path)
    _format_headers(workbook)
    _autosize_columns(workbook)
    _create_charts_sheet(workbook)
    workbook.save(output_path)


def _sheet_name(value: str) -> str:
    sanitized = value.replace("/", "_").replace("\\", "_")
    return sanitized[:31]


def _build_pending_sheet(writer: pd.ExcelWriter, oci_mapping_df: pd.DataFrame) -> None:
    pending = oci_mapping_df[oci_mapping_df["oci_service"] == "REVIEW_REQUIRED"].copy()
    if pending.empty:
        pending = pd.DataFrame(
            [{"status": "Sem pendencias de mapeamento para revisao manual."}]
        )
    pending.to_excel(writer, sheet_name="Pendencias", index=False)


def _format_headers(workbook) -> None:
    for worksheet in workbook.worksheets:
        for row in worksheet.iter_rows():
            for cell in row:
                is_header = cell.row == 1
                cell.font = Font(name=DEFAULT_FONT, bold=is_header)


def _autosize_columns(workbook) -> None:
    for worksheet in workbook.worksheets:
        for column_cells in worksheet.columns:
            length = max(len(str(cell.value or "")) for cell in column_cells)
            worksheet.column_dimensions[column_cells[0].column_letter].width = min(
                max(length + 2, 12), 40
            )


def _create_charts_sheet(workbook) -> None:
    if "Charts" in workbook.sheetnames:
        del workbook["Charts"]

    charts_sheet = workbook.create_sheet("Charts")
    charts_sheet["A1"] = "Graficos de Billing"
    charts_sheet["A1"].font = Font(name=DEFAULT_FONT, bold=True, size=14)

    _add_service_cost_chart(workbook, charts_sheet)
    _add_service_cost_share_chart(workbook, charts_sheet)
    _add_region_cost_chart(workbook, charts_sheet)
    _add_optional_chart(
        workbook,
        charts_sheet,
        sheet_name="purchase_option",
        title="Custo por Modelo de Compra",
        category_axis="Modelo",
        anchor="X3",
    )
    _add_optional_chart(
        workbook,
        charts_sheet,
        sheet_name="linked_account_name",
        title="Top Contas por Custo",
        category_axis="Conta",
        anchor="X25",
    )
    _add_optional_chart(
        workbook,
        charts_sheet,
        sheet_name="operation",
        title="Top Operacoes por Custo",
        category_axis="Operacao",
        anchor="A47",
    )


def _add_service_cost_chart(workbook, charts_sheet) -> None:
    service_sheet = workbook["Resumo_Servicos"]
    chart = BarChart()
    chart.title = "Top 10 Servicos por Custo"
    chart.y_axis.title = "Custo"
    chart.x_axis.title = "Servico"
    chart.height = 10
    chart.width = 20

    max_row = min(service_sheet.max_row, 11)
    if max_row < 2:
        return

    data = Reference(service_sheet, min_col=4, min_row=1, max_row=max_row)
    categories = Reference(service_sheet, min_col=2, min_row=2, max_row=max_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)
    charts_sheet.add_chart(chart, "A3")


def _add_region_cost_chart(workbook, charts_sheet) -> None:
    region_sheet = workbook["Resumo_Regioes"]
    chart = BarChart()
    chart.title = "Top Regioes por Custo"
    chart.y_axis.title = "Custo"
    chart.x_axis.title = "Regiao"
    chart.height = 10
    chart.width = 20

    max_row = min(region_sheet.max_row, 11)
    if max_row < 2:
        return

    data = Reference(region_sheet, min_col=3, min_row=1, max_row=max_row)
    categories = Reference(region_sheet, min_col=2, min_row=2, max_row=max_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)
    charts_sheet.add_chart(chart, "A25")


def _add_service_cost_share_chart(workbook, charts_sheet) -> None:
    service_sheet = workbook["Resumo_Servicos"]
    max_row = min(service_sheet.max_row, 11)
    if max_row < 2:
        return

    chart = PieChart()
    chart.title = "Participacao Percentual do Custo por Servico"
    chart.height = 12
    chart.width = 18

    data = Reference(service_sheet, min_col=4, min_row=1, max_row=max_row)
    categories = Reference(service_sheet, min_col=2, min_row=2, max_row=max_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showPercent = True
    chart.dataLabels.showLeaderLines = True
    charts_sheet.add_chart(chart, "X47")


def _add_optional_chart(
    workbook,
    charts_sheet,
    sheet_name: str,
    title: str,
    category_axis: str,
    anchor: str,
) -> None:
    if sheet_name not in workbook.sheetnames:
        return

    data_sheet = workbook[sheet_name]
    max_row = min(data_sheet.max_row, 11)
    if max_row < 2:
        return

    chart = BarChart()
    chart.title = title
    chart.y_axis.title = "Custo"
    chart.x_axis.title = category_axis
    chart.height = 10
    chart.width = 20

    data = Reference(data_sheet, min_col=3, min_row=1, max_row=max_row)
    categories = Reference(data_sheet, min_col=1, min_row=2, max_row=max_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)
    charts_sheet.add_chart(chart, anchor)
