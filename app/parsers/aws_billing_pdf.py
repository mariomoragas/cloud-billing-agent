from __future__ import annotations

import re
from pathlib import Path

import pandas as pd
import pdfplumber

from app.report_types import ParserResult

HEADER_AMOUNT_RE = re.compile(
    r"^(?P<label>.+?)\s+USD\s+(?P<amount>[0-9,]+(?:\.\d+)?)\s*$"
)
PAREN_AMOUNT_LINE_RE = re.compile(
    r"^(?P<label>.+?)\s+\(USD\s*(?P<amount>[0-9,]+(?:\.\d+)?)\)\s*$"
)
GRAND_TOTAL_RE = re.compile(
    r"Grand total:\s*USD\s*(?P<amount>[0-9,]+(?:\.\d+)?)",
    flags=re.IGNORECASE,
)
USAGE_AMOUNT_SUFFIX_RE = re.compile(
    r"(?:\bUSD\s*(?P<plain>[0-9,]+(?:\.\d+)?)|\(USD\s*(?P<paren>[0-9,]+(?:\.\d+)?)\))\s*$"
)
QTY_UNIT_SUFFIX_RE = re.compile(
    r"(?P<qty>[0-9][0-9,]*(?:\.\d+)?)\s+"
    r"(?P<unit>[A-Za-z][A-Za-z0-9/%().:-]*(?:\s+[A-Za-z][A-Za-z0-9/%().:-]*){0,2})\s*$"
)
PERIOD_RE = re.compile(r"\b(20\d{2})[-/](0[1-9]|1[0-2])\b")

REGION_NAME_TO_CODE = {
    "South America (Sao Paulo)": "sa-east-1",
    "US East (N. Virginia)": "us-east-1",
    "US East (Northern Virginia)": "us-east-1",
    "US East (Ohio)": "us-east-2",
    "US West (Oregon)": "us-west-2",
    "US West (N. California)": "us-west-1",
    "EU (London)": "eu-west-2",
    "EU (Ireland)": "eu-west-1",
    "EU (Frankfurt)": "eu-central-1",
    "EU (Stockholm)": "eu-north-1",
    "Canada (Central)": "ca-central-1",
    "Asia Pacific (Sydney)": "ap-southeast-2",
    "Asia Pacific (Tokyo)": "ap-northeast-1",
    "Asia Pacific (Seoul)": "ap-northeast-2",
    "Asia Pacific (Singapore)": "ap-southeast-1",
    "Asia Pacific (Hong Kong)": "ap-east-1",
    "Asia Pacific (Mumbai)": "ap-south-1",
    "Middle East (Bahrain)": "me-south-1",
    "Africa (Cape Town)": "af-south-1",
    "Global": "global",
    "Any": "global",
}


def load_aws_billing_pdf(pdf_path: Path) -> ParserResult:
    rows: list[dict[str, object]] = []
    total_text_lines = 0
    skipped_lines = 0
    header_lines = 0

    current_service = ""
    current_region = ""
    period = ""
    grand_total = 0.0
    seen_discount_lines: set[str] = set()

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            lines = [line.strip() for line in page_text.splitlines() if line.strip()]
            total_text_lines += len(lines)

            if not period:
                period = _extract_period(page_text)
            if grand_total == 0.0:
                grand_total = _extract_grand_total(page_text)

            for line in lines:
                if line.lower() == "description usage quantity amount in usd":
                    continue

                parsed_line = _parse_usage_line(line)
                if parsed_line is not None:
                    description, usage_quantity, usage_unit, cost = parsed_line
                    inferred_region = _infer_region(description) or current_region or ""
                    service_name = _infer_service_name(description, current_service)

                    rows.append(
                        {
                            "cloud": "aws",
                            "service_name_original": service_name,
                            "product_code": _infer_product_code(service_name),
                            "sku": description,
                            "region": inferred_region,
                            "usage_quantity": usage_quantity,
                            "usage_unit": usage_unit,
                            "cost": cost,
                            "currency": "USD",
                            "period": period,
                        }
                    )
                    continue

                discount_line = _parse_discount_adjustment_line(line)
                if discount_line is not None:
                    service_name, amount = discount_line
                    if line in seen_discount_lines:
                        continue
                    seen_discount_lines.add(line)
                    rows.append(
                        {
                            "cloud": "aws",
                            "service_name_original": service_name,
                            "product_code": _infer_product_code(service_name),
                            "sku": line,
                            "region": current_region or "",
                            "usage_quantity": 1.0,
                            "usage_unit": "Adjustment",
                            "cost": amount,
                            "currency": "USD",
                            "period": period,
                        }
                    )
                    continue

                header_match = HEADER_AMOUNT_RE.match(line)
                if header_match:
                    header_lines += 1
                    label = header_match.group("label").strip()
                    if _is_region_label(label):
                        current_region = _normalize_region(label)
                    else:
                        current_service = label
                    continue

                skipped_lines += 1

    if not rows:
        raise ValueError(
            "Nao foi possivel extrair linhas de consumo do PDF AWS. "
            "Verifique se o layout segue o padrao de Billing and Cost Management."
        )

    normalized_df = pd.DataFrame(rows)
    normalized_df["service_name_original"] = (
        normalized_df["service_name_original"].fillna("").astype(str).str.strip()
    )
    normalized_df["sku"] = normalized_df["sku"].fillna("").astype(str).str.strip()
    normalized_df["region"] = normalized_df["region"].fillna("").astype(str).str.strip()
    normalized_df["usage_unit"] = (
        normalized_df["usage_unit"].fillna("").astype(str).str.strip()
    )
    normalized_df["period"] = normalized_df["period"].fillna("").astype(str).str.strip()
    normalized_df["usage_quantity"] = pd.to_numeric(
        normalized_df["usage_quantity"], errors="coerce"
    ).fillna(0.0)
    normalized_df["cost"] = pd.to_numeric(normalized_df["cost"], errors="coerce").fillna(0.0)
    parsed_total = float(normalized_df["cost"].sum())
    reconciliation_adjustment = grand_total - parsed_total if grand_total > 0 else 0.0
    if grand_total > 0 and abs(reconciliation_adjustment) >= 0.01:
        normalized_df = pd.concat(
            [
                normalized_df,
                pd.DataFrame(
                    [
                        {
                            "cloud": "aws",
                            "service_name_original": "Invoice Total Reconciliation",
                            "product_code": "",
                            "sku": f"Reconciliation to Grand total on PDF cover page ({grand_total:,.2f} USD)",
                            "region": "",
                            "usage_quantity": 1.0,
                            "usage_unit": "Reconciliation",
                            "cost": reconciliation_adjustment,
                            "currency": "USD",
                            "period": period,
                        }
                    ]
                ),
            ],
            ignore_index=True,
        )
        parsed_total = float(normalized_df["cost"].sum())

    data_quality = pd.DataFrame(
        [
            {"metric": "pdf_total_text_lines", "value": total_text_lines},
            {"metric": "pdf_header_lines_detected", "value": header_lines},
            {"metric": "pdf_usage_lines_parsed", "value": len(normalized_df)},
            {"metric": "pdf_lines_skipped", "value": skipped_lines},
            {"metric": "pdf_grand_total_page1", "value": grand_total},
            {"metric": "pdf_reconciliation_adjustment", "value": reconciliation_adjustment},
            {"metric": "final_rows", "value": len(normalized_df)},
            {"metric": "final_cost_total", "value": parsed_total},
        ]
    )
    return ParserResult(dataframe=normalized_df, data_quality=data_quality)


def _parse_usage_line(line: str) -> tuple[str, float, str, float] | None:
    amount_match = USAGE_AMOUNT_SUFFIX_RE.search(line)
    if amount_match is None:
        return None

    plain_amount = amount_match.group("plain")
    paren_amount = amount_match.group("paren")
    amount_raw = plain_amount or paren_amount
    if amount_raw is None:
        return None
    cost = _to_float(amount_raw)
    if paren_amount is not None:
        # In AWS PDF bills, parenthesized USD usually means discount/credit.
        cost = -abs(cost)

    without_amount = line[: amount_match.start()].strip()
    qty_unit_match = QTY_UNIT_SUFFIX_RE.search(without_amount)
    if qty_unit_match is None:
        return None

    quantity = _to_float(qty_unit_match.group("qty"))
    usage_unit = qty_unit_match.group("unit").strip()
    description = without_amount[: qty_unit_match.start()].strip()
    if not description:
        return None
    if len(description) <= 3 and "savings plan" in usage_unit.lower():
        # OCR line split like "EC 2.000 Instance Savings Plans USD 7140.91"
        return None
    return description, quantity, usage_unit, cost


def _parse_discount_adjustment_line(line: str) -> tuple[str, float] | None:
    match = PAREN_AMOUNT_LINE_RE.match(line)
    if match is None:
        return None

    label = match.group("label").strip()
    if not label:
        return None

    lowered = label.lower()
    # Keep only detailed discount lines to avoid duplicating section summaries.
    is_distributor_discount = lowered.startswith("distributor discounts")
    is_bundled_discount = "bundled discount" in lowered
    if not is_distributor_discount and not is_bundled_discount:
        return None

    amount = -abs(_to_float(match.group("amount")))
    service_name = "Discounts and Credits"
    if is_distributor_discount:
        service_name = "Distributor Discounts"
    elif is_bundled_discount:
        service_name = "Usage Based Discounts"
    return service_name, amount


def _extract_period(text: str) -> str:
    match = PERIOD_RE.search(text)
    if match is None:
        return ""
    return f"{match.group(1)}-{match.group(2)}"


def _extract_grand_total(text: str) -> float:
    match = GRAND_TOTAL_RE.search(text)
    if match is None:
        return 0.0
    return _to_float(match.group("amount"))


def _to_float(value: str) -> float:
    return float(str(value).replace(",", "").strip())


def _is_region_label(label: str) -> bool:
    cleaned = label.strip()
    if cleaned in REGION_NAME_TO_CODE:
        return True
    return bool(
        cleaned.startswith(("US ", "EU ", "Asia Pacific", "South America", "Canada", "Middle East", "Africa"))
        and "(" in cleaned
        and ")" in cleaned
    )


def _normalize_region(value: str) -> str:
    cleaned = value.strip()
    return REGION_NAME_TO_CODE.get(cleaned, cleaned.lower().replace(" ", "_"))


def _infer_region(description: str) -> str:
    for label, aws_region in REGION_NAME_TO_CODE.items():
        if label in description:
            return aws_region
    return ""


def _infer_service_name(description: str, current_service: str) -> str:
    if current_service.strip():
        return current_service.strip()

    description_lower = description.lower()
    fallback_rules = [
        ("nat gateway", "Amazon Elastic Compute Cloud NatGateway"),
        ("instance hour", "Amazon Elastic Compute Cloud"),
        ("linux/unix spot", "Amazon Elastic Compute Cloud"),
        ("ebs", "EBS"),
        ("snapshot data", "EBS"),
        ("data transfer", "Data Transfer"),
        ("cloudwatch", "Amazon CloudWatch"),
        ("elastic load balancing", "Elastic Load Balancing"),
        ("fargate", "Elastic Container Service"),
        ("virtual private cloud", "Amazon Virtual Private Cloud"),
    ]
    for token, label in fallback_rules:
        if token in description_lower:
            return label

    description_clean = description.strip()
    if description_clean:
        return description_clean[:120]
    return "AWS Service"


def _infer_product_code(service_name: str) -> str:
    service = service_name.lower()
    mapping = {
        "elastic compute cloud": "AmazonEC2",
        "ec2": "AmazonEC2",
        "s3": "AmazonS3",
        "elastic load balancing": "AWSELB",
        "cloudwatch": "AmazonCloudWatch",
        "data transfer": "AWSDataTransfer",
        "relational database service": "AmazonRDS",
        "dynamodb": "AmazonDynamoDB",
        "cloudfront": "AmazonCloudFront",
        "secrets manager": "AWSSecretsManager",
        "vpc": "AmazonVPC",
    }
    for key, code in mapping.items():
        if key in service:
            return code
    return ""
