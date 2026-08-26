from __future__ import annotations

import unittest

import pandas as pd

from app.oci_bom import build_oci_bom


def _row(**overrides):
    row = {
        "cloud": "aws",
        "product_code": "AmazonEC2",
        "service_name_original": "Amazon Elastic Compute Cloud",
        "sku": "",
        "usage_type": "",
        "usage_quantity": 0.0,
        "usage_unit": "Hours",
        "cost": 1.0,
        "currency": "USD",
        "region": "sa-east-1",
    }
    row.update(overrides)
    row["sku"] = row["usage_type"]
    return row


def _quantity(bom, part_number):
    return float(bom.loc[bom["oci_part_number"] == part_number, "quantity"].sum())


class OciBomTests(unittest.TestCase):
    def test_compute_capacity_converts_x86_and_arm_without_contract_double_counting(self):
        raw = pd.DataFrame(
            [
                _row(usage_type="SAE1-BoxUsage:c6a.4xlarge", usage_quantity=10),
                _row(usage_type="SAE1-SpotUsage:t4g.xlarge", usage_quantity=5, usage_unit=""),
                _row(usage_type="SAE1-HeavyUsage:c6a.4xlarge", usage_quantity=10, cost=100),
            ]
        )

        artifacts = build_oci_bom(raw)

        self.assertAlmostEqual(_quantity(artifacts.bom_df, "B97384"), 80.0)
        self.assertAlmostEqual(_quantity(artifacts.bom_df, "B97385"), 320.0)
        self.assertAlmostEqual(_quantity(artifacts.bom_df, "B93297"), 20.0)
        self.assertAlmostEqual(_quantity(artifacts.bom_df, "B93298"), 80.0)
        arm_row = artifacts.bom_df[artifacts.bom_df["oci_part_number"] == "B93297"].iloc[0]
        self.assertEqual(arm_row["source_unit"], "Hours")
        self.assertTrue(artifacts.review_df["review_reason"].str.contains("double counting", case=False).any())

    def test_gpu_block_storage_and_s3_requests_generate_official_part_numbers(self):
        raw = pd.DataFrame(
            [
                _row(usage_type="BoxUsage:g6e.xlarge", usage_quantity=3, region=""),
                _row(usage_type="SAE1-EBS:VolumeUsage.gp3", usage_quantity=100, usage_unit="GB"),
                _row(
                    product_code="AmazonS3",
                    service_name_original="Amazon Simple Storage Service",
                    usage_type="SAE1-Requests-Tier1",
                    usage_quantity=20000,
                    usage_unit="Requests",
                ),
            ]
        )

        artifacts = build_oci_bom(raw)

        self.assertAlmostEqual(_quantity(artifacts.bom_df, "B109479"), 3.0)
        self.assertAlmostEqual(_quantity(artifacts.bom_df, "B91961"), 100.0)
        self.assertAlmostEqual(_quantity(artifacts.bom_df, "B91962"), 1000.0)
        self.assertAlmostEqual(_quantity(artifacts.bom_df, "B91627"), 2.0)
        self.assertEqual(set(artifacts.bom_df["scenario"]), {"Capacity equivalent"})

    def test_unknown_compute_shape_is_sent_to_review_instead_of_invented(self):
        raw = pd.DataFrame([_row(usage_type="BoxUsage:unknown.xlarge", usage_quantity=12, cost=25)])

        artifacts = build_oci_bom(raw)

        self.assertTrue(artifacts.bom_df.empty)
        self.assertTrue(artifacts.review_df["source_sku"].str.contains("unknown.xlarge", case=False).any())
        self.assertTrue(artifacts.review_df["review_reason"].str.contains("could not be resolved", case=False).any())


if __name__ == "__main__":
    unittest.main()
