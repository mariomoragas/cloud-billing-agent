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


def _gcp_row(**overrides):
    row = {
        "cloud": "gcp",
        "product_code": "service-id",
        "service_name_original": "Compute Engine",
        "sku": "",
        "sku_id": "sku-id",
        "usage_type": "",
        "usage_quantity": 0.0,
        "usage_unit": "hour",
        "cost": 1.0,
        "currency": "USD",
        "region": "southamerica-east1",
    }
    row.update(overrides)
    if not row["sku"]:
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


    def test_gcp_compute_uses_usage_rows_and_excludes_commitments(self):
        raw = pd.DataFrame(
            [
                _gcp_row(
                    usage_type="N2 Instance Core running in Sao Paulo",
                    usage_quantity=20,
                ),
                _gcp_row(
                    usage_type="N2 Instance Ram running in Sao Paulo",
                    usage_quantity=80,
                    usage_unit="gibibyte hour",
                ),
                _gcp_row(
                    usage_type="Commitment v1: N2 Cpu in Sao Paulo for 1 Year",
                    usage_quantity=20,
                    cost=50,
                ),
            ]
        )

        artifacts = build_oci_bom(raw)

        self.assertAlmostEqual(_quantity(artifacts.bom_df, "B97384"), 10.0)
        self.assertAlmostEqual(_quantity(artifacts.bom_df, "B97385"), 80.0)
        self.assertTrue(
            artifacts.review_df["review_reason"]
            .str.contains("double counting", case=False)
            .any()
        )
        self.assertTrue(
            artifacts.bom_df["source_skus"].str.contains("N2 Instance Core").any()
        )

    def test_gcp_storage_network_and_requests_generate_oci_part_numbers(self):
        raw = pd.DataFrame(
            [
                _gcp_row(
                    usage_type="Balanced PD Capacity in Sao Paulo",
                    usage_quantity=100,
                    usage_unit="gibibyte month",
                ),
                _gcp_row(
                    service_name_original="Cloud Storage",
                    usage_type="Standard Storage Sao Paulo",
                    usage_quantity=200,
                    usage_unit="gibibyte month",
                ),
                _gcp_row(
                    service_name_original="Cloud Storage",
                    usage_type="Archive Storage Sao Paulo",
                    usage_quantity=300,
                    usage_unit="gibibyte month",
                ),
                _gcp_row(
                    service_name_original="Cloud Storage",
                    usage_type="Regional Standard Class A Operations",
                    usage_quantity=20000,
                    usage_unit="count",
                ),
                _gcp_row(
                    service_name_original="Cloud Filestore",
                    usage_type="Filestore Capacity Regional and Enterprise Sao Paulo",
                    usage_quantity=50,
                    usage_unit="gibibyte month",
                ),
                _gcp_row(
                    usage_type="Network Internet Data Transfer Out from Sao Paulo to Americas",
                    usage_quantity=12,
                    usage_unit="gibibyte",
                ),
                _gcp_row(
                    service_name_original="Networking",
                    usage_type="Cloud Load Balancer Forwarding Rule Additional for Sao Paulo",
                    usage_quantity=24,
                    usage_unit="hour",
                ),
            ]
        )

        artifacts = build_oci_bom(raw)

        expected = {
            "B91961": 100.0,
            "B91962": 1000.0,
            "B91628": 200.0,
            "B91633": 300.0,
            "B91627": 2.0,
            "B89057": 50.0,
            "B93455": 12.0,
            "B93030": 24.0,
        }
        for part_number, quantity in expected.items():
            self.assertAlmostEqual(
                _quantity(artifacts.bom_df, part_number),
                quantity,
                msg=part_number,
            )

    def test_gcp_managed_service_without_safe_quantity_stays_in_review(self):
        raw = pd.DataFrame(
            [
                _gcp_row(
                    service_name_original="Cloud SQL",
                    usage_type="Cloud SQL for PostgreSQL: Zonal - vCPU in Sao Paulo",
                    usage_quantity=100,
                    cost=200,
                )
            ]
        )

        artifacts = build_oci_bom(raw)

        self.assertTrue(artifacts.bom_df.empty)
        self.assertEqual(
            artifacts.review_df.iloc[0]["source_sku"],
            "Cloud SQL for PostgreSQL: Zonal - vCPU in Sao Paulo",
        )
        self.assertIn("Managed GCP service", artifacts.review_df.iloc[0]["review_reason"])

if __name__ == "__main__":
    unittest.main()
