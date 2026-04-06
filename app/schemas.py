from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class BillingRecord:
    cloud: str
    service_name_original: str
    sku: str
    region: str
    usage_quantity: float
    usage_unit: str
    cost: float
    currency: str
    period: str
