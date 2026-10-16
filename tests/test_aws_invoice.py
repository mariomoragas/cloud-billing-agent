from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import pandas as pd

from app.parsers.aws_invoice import (
    _deduplicate_payer_linked_rows,
    load_aws_invoice_csv,
)


class AwsInvoiceClientRefTests(unittest.TestCase):
    def test_client_ref_is_used_when_linked_account_columns_are_absent(self) -> None:
        rows = [
            {
                "InvoiceID": "invoice-1",
                "ClientRef": "CLIENT_01",
                "RecordType": "LinkedLineItem",
                "BillingPeriodStartDate": "2026/08/01 00:00:00",
                "ProductCode": "AmazonEC2",
                "ProductName": "Amazon Elastic Compute Cloud",
                "UsageType": "SAE1-BoxUsage:m6i.large",
                "Operation": "",
                "UsageQuantity": "10",
                "CurrencyCode": "USD",
                "CostBeforeTax": "5",
                "TaxAmount": "0",
                "TotalCost": "5",
            },
            {
                "InvoiceID": "invoice-1",
                "ClientRef": "CLIENT_01",
                "RecordType": "LinkedLineItem",
                "BillingPeriodStartDate": "2026/08/01 00:00:00",
                "ProductCode": "AmazonS3",
                "ProductName": "Amazon Simple Storage Service",
                "UsageType": "SAE1-TimedStorage-ByteHrs",
                "Operation": "",
                "UsageQuantity": "20",
                "CurrencyCode": "USD",
                "CostBeforeTax": "2",
                "TaxAmount": "0",
                "TotalCost": "2",
            },
        ]
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "invoice.csv"
            pd.DataFrame(rows).to_csv(path, index=False)
            result = load_aws_invoice_csv(path)

        self.assertEqual(len(result.dataframe), 2)
        self.assertEqual(set(result.dataframe["client_ref"]), {"CLIENT_01"})
        self.assertEqual(
            set(result.dataframe["linked_account_name"]), {"CLIENT_01"}
        )
        self.assertEqual(
            set(result.dataframe["linked_account_id"]), {"CLIENT_01"}
        )

    def test_empty_deduplication_input_is_returned_without_error(self) -> None:
        result = _deduplicate_payer_linked_rows(pd.DataFrame())
        self.assertTrue(result.empty)


if __name__ == "__main__":
    unittest.main()
