from __future__ import annotations

from pathlib import Path

import pandas as pd
from pptx import Presentation
from pptx.chart.data import ChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.util import Inches, Pt

FONT_NAME = "Verdana"


def write_powerpoint_report(
    output_path: Path,
    raw_df: pd.DataFrame,
    service_summary_df: pd.DataFrame,
    region_summary_df: pd.DataFrame,
    oci_mapping_df: pd.DataFrame,
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    presentation = Presentation()
    _add_title_slide(presentation, raw_df)
    _add_kpi_slide(presentation, raw_df, oci_mapping_df)
    _add_top_services_slide(presentation, service_summary_df)
    _add_service_share_slide(presentation, service_summary_df)
    _add_region_slide(presentation, region_summary_df)
    _add_unmapped_slide(presentation, oci_mapping_df)
    presentation.save(output_path)


def _add_title_slide(presentation: Presentation, raw_df: pd.DataFrame) -> None:
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    _add_title(slide, "Cloud Billing Executive Summary", 0.6, 0.5, 9.0)
    period = _mode_or_default(raw_df, "period", "N/A")
    cloud = _mode_or_default(raw_df, "cloud", "cloud").upper()
    subtitle = slide.shapes.add_textbox(Inches(0.6), Inches(1.5), Inches(8.8), Inches(0.8))
    paragraph = subtitle.text_frame.paragraphs[0]
    paragraph.text = f"Billing period: {period} | Source: {cloud}"
    _style_paragraph(paragraph, bold=False, size=22, color=(82, 96, 109))
    _add_footer(slide, "Generated automatically from the billing pipeline.")


def _add_kpi_slide(
    presentation: Presentation,
    raw_df: pd.DataFrame,
    oci_mapping_df: pd.DataFrame,
) -> None:
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    _add_title(slide, "Key Metrics", 0.6, 0.5, 5.0)

    total_cost = float(raw_df["cost"].sum()) if not raw_df.empty else 0.0
    currency = _mode_or_default(raw_df, "currency", "USD")
    service_count = int(raw_df["service_name_original"].nunique()) if not raw_df.empty else 0
    unmapped_count = (
        int((oci_mapping_df["oci_service"] == "REVIEW_REQUIRED").sum())
        if not oci_mapping_df.empty
        else 0
    )
    cards = [
        ("Total Cost", f"{total_cost:,.2f} {currency}"),
        ("Distinct Services", str(service_count)),
        ("Unmapped Services", str(unmapped_count)),
    ]
    for index, (label, value) in enumerate(cards):
        _add_kpi_card(slide, label, value, 0.6 + (index * 3.0), 1.6)
    _add_footer(slide, "Numbers reflect the same filtered dataset used in Excel.")


def _add_top_services_slide(presentation: Presentation, service_summary_df: pd.DataFrame) -> None:
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    _add_title(slide, "Top Services by Cost", 0.6, 0.5, 6.0)

    top_services = service_summary_df.head(8)
    chart_data = ChartData()
    chart_data.categories = list(top_services["service_name_original"])
    chart_data.add_series("Cost", list(top_services["total_cost"]))

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(0.6),
        Inches(1.3),
        Inches(8.8),
        Inches(4.9),
        chart_data,
    ).chart
    chart.has_legend = False
    chart.value_axis.has_major_gridlines = True
    _set_chart_fonts(chart)


def _add_service_share_slide(
    presentation: Presentation,
    service_summary_df: pd.DataFrame,
) -> None:
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    _add_title(slide, "Service Cost Share", 0.6, 0.5, 6.0)

    top_services = service_summary_df.head(6)
    chart_data = ChartData()
    chart_data.categories = list(top_services["service_name_original"])
    chart_data.add_series("Cost Share", list(top_services["total_cost"]))

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.PIE,
        Inches(0.8),
        Inches(1.2),
        Inches(8.2),
        Inches(5.0),
        chart_data,
    ).chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.RIGHT
    chart.plots[0].has_data_labels = True
    chart.plots[0].data_labels.show_percentage = True
    _set_chart_fonts(chart)


def _add_region_slide(presentation: Presentation, region_summary_df: pd.DataFrame) -> None:
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    _add_title(slide, "Top Regions by Cost", 0.6, 0.5, 6.0)

    top_regions = region_summary_df.head(8).copy()
    if "region" in top_regions.columns:
        top_regions["region"] = top_regions["region"].replace("", "UNSPECIFIED")
    chart_data = ChartData()
    chart_data.categories = list(top_regions["region"])
    chart_data.add_series("Cost", list(top_regions["total_cost"]))

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(0.6),
        Inches(1.3),
        Inches(8.8),
        Inches(4.9),
        chart_data,
    ).chart
    chart.has_legend = False
    chart.value_axis.has_major_gridlines = True
    _set_chart_fonts(chart)


def _add_unmapped_slide(presentation: Presentation, oci_mapping_df: pd.DataFrame) -> None:
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    _add_title(slide, "Services Requiring OCI Review", 0.6, 0.5, 7.0)

    unmapped_df = oci_mapping_df[oci_mapping_df["oci_service"] == "REVIEW_REQUIRED"].head(12)
    box = slide.shapes.add_textbox(Inches(0.7), Inches(1.4), Inches(8.6), Inches(4.9))
    frame = box.text_frame
    frame.word_wrap = True

    if unmapped_df.empty:
        paragraph = frame.paragraphs[0]
        paragraph.text = "All services in this report have an initial OCI mapping."
        _style_paragraph(paragraph, bold=False, size=18, color=(31, 42, 55))
        return

    first = True
    for _, row in unmapped_df.iterrows():
        paragraph = frame.paragraphs[0] if first else frame.add_paragraph()
        first = False
        paragraph.text = str(row["service_name_original"])
        _style_paragraph(paragraph, bold=False, size=18, color=(31, 42, 55))


def _add_title(slide, text: str, left: float, top: float, width: float) -> None:
    shape = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(0.8))
    paragraph = shape.text_frame.paragraphs[0]
    paragraph.text = text
    _style_paragraph(paragraph, bold=True, size=26, color=(31, 42, 55))


def _add_kpi_card(slide, label: str, value: str, left: float, top: float) -> None:
    shape = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
        Inches(left),
        Inches(top),
        Inches(2.6),
        Inches(2.0),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(255, 250, 242)
    shape.line.color.rgb = RGBColor(234, 223, 207)

    frame = shape.text_frame
    frame.clear()
    p1 = frame.paragraphs[0]
    p1.text = label
    _style_paragraph(p1, bold=False, size=16, color=(82, 96, 109))
    p2 = frame.add_paragraph()
    p2.text = value
    _style_paragraph(p2, bold=True, size=24, color=(31, 42, 55))


def _add_footer(slide, text: str) -> None:
    shape = slide.shapes.add_textbox(Inches(0.6), Inches(6.8), Inches(8.8), Inches(0.4))
    paragraph = shape.text_frame.paragraphs[0]
    paragraph.text = text
    _style_paragraph(paragraph, bold=False, size=10, color=(82, 96, 109))


def _style_paragraph(paragraph, *, bold: bool, size: int, color: tuple[int, int, int]) -> None:
    paragraph.font.name = FONT_NAME
    paragraph.font.size = Pt(size)
    paragraph.font.bold = bold
    paragraph.font.color.rgb = RGBColor(*color)


def _set_chart_fonts(chart) -> None:
    if chart.has_title and chart.chart_title is not None:
        for paragraph in chart.chart_title.text_frame.paragraphs:
            paragraph.font.name = FONT_NAME
    if chart.has_legend and chart.legend is not None:
        chart.legend.font.name = FONT_NAME
        chart.legend.font.size = Pt(10)
    try:
        category_axis = chart.category_axis
        category_axis.tick_labels.font.name = FONT_NAME
        category_axis.tick_labels.font.size = Pt(10)
    except (AttributeError, ValueError):
        pass
    try:
        value_axis = chart.value_axis
        value_axis.tick_labels.font.name = FONT_NAME
        value_axis.tick_labels.font.size = Pt(10)
    except (AttributeError, ValueError):
        pass


def _mode_or_default(df: pd.DataFrame, column: str, default: str) -> str:
    if df.empty or column not in df.columns:
        return default
    series = df[column].fillna("").astype(str).str.strip()
    series = series[series != ""]
    if series.empty:
        return default
    return str(series.mode().iloc[0])
