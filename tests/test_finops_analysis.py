from __future__ import annotations

import unittest

import pandas as pd

from app.finops_analysis import build_finops_analysis
from app.oci_bom import build_oci_bom


class FinOpsAnalysisTests(unittest.TestCase):
    def _raw(self) -> pd.DataFrame:
        return pd.DataFrame(
            [
                {
                    "cloud": "aws",
                    "service_name_original": "Amazon Elastic Compute Cloud",
                    "product_code": "AmazonEC2",
                    "sku": "BoxUsage:c6a.2xlarge",
                    "usage_type": "USE1-BoxUsage:c6a.2xlarge",
                    "region": "us-east-1",
                    "usage_quantity": 10.0,
                    "usage_unit": "Hours",
                    "cost": 11.0,
                    "cost_before_tax": 10.0,
                    "tax_amount": 1.0,
                    "credits": 0.0,
                    "currency": "USD",
                }
            ]
        )

    def test_baseline_separates_tax_from_comparable_cost(self) -> None:
        raw = self._raw()
        bom = build_oci_bom(raw).bom_df
        artifacts = build_finops_analysis(raw, bom)
        values = artifacts.baseline_df.set_index("metric")["value"]
        self.assertEqual(values["Gross billed cost"], 11.0)
        self.assertEqual(values["Comparable pre-tax baseline"], 10.0)
        self.assertEqual(values["Taxes"], 1.0)

    def test_source_cost_is_not_double_counted_across_ocpu_and_memory(self) -> None:
        raw = self._raw()
        bom = build_oci_bom(raw).bom_df
        self.assertEqual(len(bom), 2)
        artifacts = build_finops_analysis(raw, bom)
        self.assertAlmostEqual(
            float(artifacts.savings_df["source_pre_tax_cost"].sum()),
            10.0,
            places=6,
        )

    def test_unpriced_scope_is_carried_forward_without_assumed_saving(self) -> None:
        raw = self._raw()
        extra = raw.iloc[0].copy()
        extra["service_name_original"] = "Unsupported Managed Service"
        extra["product_code"] = "Unsupported"
        extra["usage_type"] = "UnsupportedUsage"
        extra["sku"] = "UnsupportedUsage"
        extra["cost"] = 5.0
        extra["cost_before_tax"] = 5.0
        extra["tax_amount"] = 0.0
        raw = pd.concat([raw, pd.DataFrame([extra])], ignore_index=True)
        bom = build_oci_bom(raw).bom_df
        artifacts = build_finops_analysis(raw, bom)
        capacity = artifacts.business_case_df.iloc[0]
        self.assertAlmostEqual(float(capacity["unpriced_carry_forward"]), 5.0, places=6)


if __name__ == "__main__":
    unittest.main()
