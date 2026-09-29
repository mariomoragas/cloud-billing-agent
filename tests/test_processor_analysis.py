from __future__ import annotations

import unittest

import pandas as pd

from app.processor_analysis import build_processor_analysis


class ProcessorAnalysisTests(unittest.TestCase):
    def test_maps_representative_shapes_from_three_clouds(self) -> None:
        raw = pd.DataFrame(
            [
                {
                    "cloud": "aws",
                    "service_name_original": "Amazon EC2",
                    "sku": "SAE1-BoxUsage:c6a.4xlarge",
                    "item_description": "c6a.4xlarge instance hour",
                    "usage_quantity": 744,
                    "cost": 500,
                },
                {
                    "cloud": "azure",
                    "service_name_original": "Virtual Machines",
                    "sku": "Standard_D8as_v5",
                    "usage_quantity": 744,
                    "cost": 400,
                },
                {
                    "cloud": "gcp",
                    "service_name_original": "Compute Engine",
                    "sku": "T2D Instance Core running in Sao Paulo",
                    "usage_quantity": 2000,
                    "cost": 300,
                },
            ]
        )

        result = build_processor_analysis(raw)

        self.assertEqual(
            set(result.detail_df["source_processor_model"]),
            {"AMD EPYC 7R13", "AMD EPYC 7763", "AMD EPYC 7B13"},
        )
        self.assertEqual(len(result.comparison_df), 3)
        self.assertTrue(
            result.comparison_df["source_benchmark_coverage_pct"].eq(100).all()
        )

    def test_excludes_arm_and_gpu_from_numeric_baseline(self) -> None:
        raw = pd.DataFrame(
            [
                {
                    "cloud": "aws",
                    "service_name_original": "Amazon EC2",
                    "sku": "BoxUsage:m8g.xlarge",
                    "usage_quantity": 10,
                    "cost": 100,
                },
                {
                    "cloud": "aws",
                    "service_name_original": "Amazon EC2",
                    "sku": "BoxUsage:g6e.xlarge",
                    "usage_quantity": 10,
                    "cost": 1000,
                },
            ]
        )

        result = build_processor_analysis(raw)

        self.assertEqual(len(result.detail_df), 2)
        self.assertTrue(result.comparison_df.empty)
        statuses = set(result.detail_df["comparison_status"])
        self.assertIn("Architecture differs; no x86 performance estimate", statuses)
        self.assertIn("GPU workload; CPU-only comparison excluded", statuses)

    def test_keeps_ambiguous_gcp_family_unresolved(self) -> None:
        raw = pd.DataFrame(
            [
                {
                    "cloud": "gcp",
                    "service_name_original": "Compute Engine",
                    "sku": "N2D Instance Core",
                    "usage_quantity": 100,
                    "cost": 50,
                }
            ]
        )

        result = build_processor_analysis(raw)

        self.assertEqual(result.detail_df.iloc[0]["source_processor_model"], "UNRESOLVED")
        self.assertTrue(result.comparison_df.empty)

    def test_gcp_n2_and_e2_generate_low_confidence_oci_comparison(self) -> None:
        raw = pd.DataFrame(
            [
                {
                    "cloud": "gcp",
                    "service_name_original": "Compute Engine",
                    "sku": "N2 Instance Core running in Sao Paulo",
                    "usage_quantity": 1000,
                    "cost": 900,
                },
                {
                    "cloud": "gcp",
                    "service_name_original": "Compute Engine",
                    "sku": "E2 Instance Core running in Sao Paulo",
                    "usage_quantity": 100,
                    "cost": 100,
                },
            ]
        )

        result = build_processor_analysis(raw)

        self.assertEqual(
            set(result.detail_df["source_processor_model"]),
            {"Intel Xeon Gold 6268CL @ 2.80GHz", "AMD EPYC 7B13"},
        )
        self.assertEqual(len(result.comparison_df), 3)
        self.assertTrue(result.comparison_df["source_cloud"].eq("gcp").all())
        self.assertTrue(result.comparison_df["confidence"].eq("Low").all())
        self.assertTrue(
            result.comparison_df["source_processor_models"]
            .str.contains("Intel Xeon Gold 6268CL")
            .all()
        )
        gains = result.comparison_df.set_index("target_shape_family")[
            "estimated_gain_pct"
        ]
        self.assertGreater(gains["E5"], 20)
        self.assertGreater(gains["E6"], 45)
    def test_suppresses_gain_when_only_comparable_shapes_have_zero_cost(self) -> None:
        raw = pd.DataFrame(
            [
                {
                    "cloud": "aws",
                    "service_name_original": "Amazon EC2",
                    "sku": "BoxUsage:m6i.large",
                    "usage_quantity": 744,
                    "cost": 0,
                },
                {
                    "cloud": "aws",
                    "service_name_original": "Amazon EC2",
                    "sku": "BoxUsage:m8g.large",
                    "usage_quantity": 744,
                    "cost": 200,
                },
            ]
        )

        result = build_processor_analysis(raw)

        self.assertTrue(result.comparison_df.empty)



if __name__ == "__main__":
    unittest.main()
