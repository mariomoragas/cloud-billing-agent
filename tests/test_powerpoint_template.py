import pandas as pd
from pptx import Presentation

from app.powerpoint_writer import (
    _add_closing_slide,
    _add_section_slide,
    _add_title_slide,
    _blank_slide_layout,
    _build_complexity_points,
    _configure_presentation_sections,
    _create_presentation,
    _find_slide_layout,
)


def test_bundled_powerpoint_template_is_reusable(monkeypatch, tmp_path):
    monkeypatch.delenv("PPT_TEMPLATE_PATH", raising=False)

    presentation = _create_presentation()

    assert len(presentation.slides) == 0
    assert round(presentation.slide_width / 914400, 3) == 13.333
    assert round(presentation.slide_height / 914400, 3) == 7.5
    assert _blank_slide_layout(presentation).name == "Light - Blank"
    assert _blank_slide_layout(presentation, dark=True).name == "Dark - Blank"
    assert _find_slide_layout(presentation, "Light - Safe Harbor").name == "Light - Safe Harbor"

    presentation.slides.add_slide(_blank_slide_layout(presentation))
    output_path = tmp_path / "template_roundtrip.pptx"
    presentation.save(output_path)

    reopened = Presentation(output_path)
    assert len(reopened.slides) == 1
    assert reopened.slides[0].slide_layout.name == "Light - Blank"


def test_title_and_section_slides_use_light_layout(monkeypatch):
    monkeypatch.delenv("PPT_TEMPLATE_PATH", raising=False)
    presentation = _create_presentation()
    raw_df = pd.DataFrame(
        [{"cost": 100.0, "currency": "USD", "period": "2026-09", "cloud": "aws"}]
    )

    _add_title_slide(presentation, raw_df, "report", "Company", "Project")
    _add_section_slide(presentation, title="Section", subtitle="Subtitle", slide_number=2)

    assert [slide.slide_layout.name for slide in presentation.slides] == [
        "Light - Blank",
        "Light - Blank",
    ]

def test_sections_move_analysis_to_commercial_and_close_in_final(monkeypatch):
    monkeypatch.delenv("PPT_TEMPLATE_PATH", raising=False)
    presentation = _create_presentation()
    for _ in range(4):
        presentation.slides.add_slide(_blank_slide_layout(presentation))

    _add_closing_slide(presentation)
    _configure_presentation_sections(presentation)

    assert presentation.slides[-1].slide_layout.name == "Light - Closing Slide"
    section_list = next(
        element
        for element in presentation.part._element.iter()
        if element.tag.endswith("}sectionLst")
    )
    sections = [(section.get("name"), len(list(section[0]))) for section in section_list]
    assert sections == [("Default Section", 3), ("Comercial", 1), ("Final", 1)]

def test_complexity_labels_use_gcp_description_and_aws_product_code():
    mapping = pd.DataFrame(
        [
            {
                "cloud": "gcp",
                "service_name_original": "Compute Engine",
                "primary_product_code": "6F81-5844-456A",
                "oci_service": "OCI Compute",
                "total_cost": 100.0,
                "complexity_score": 1,
            },
            {
                "cloud": "aws",
                "service_name_original": "Amazon Elastic Compute Cloud",
                "primary_product_code": "AmazonEC2",
                "oci_service": "OCI Compute",
                "total_cost": 80.0,
                "complexity_score": 1,
            },
        ]
    )

    points = _build_complexity_points(mapping, top_n=10)

    assert set(points["label"]) == {"Compute Engine", "AmazonEC2"}
    assert "6F81-5844-456A" not in set(points["label"])
