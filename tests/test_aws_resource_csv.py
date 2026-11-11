from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from app.parsers.aws_resource_csv import load_aws_resource_csv


class AwsResourceCsvTests(unittest.TestCase):
    def test_parses_semicolon_csv_and_brazilian_numbers(self) -> None:
        content = (
            "invoice;emissor;conta_id;conta_nome;servico;regiao;tipo_item;usage_type;operacao;tipo_instancia;resource_id;tag_SETOR;tag_createdBy;quantidade_uso;unidade;custo_usd;custo_preco_tabela_usd;descricao\n"
            "1;AWS;123;Conta;Amazon Elastic Compute Cloud;us-east-1;Usage;BoxUsage:r6a.8xlarge;RunInstances;r6a.8xlarge;i-1;;;404,1014;Hrs;2877,20;3000,00;Compute\n"
        )
        with TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "aws-fatura-2026-10-por-recurso.csv"
            path.write_text(content, encoding="utf-8")
            result = load_aws_resource_csv(path)

        row = result.dataframe.iloc[0]
        self.assertEqual(row["product_code"], "AmazonEC2")
        self.assertEqual(row["period"], "2026-10")
        self.assertEqual(row["usage_unit"], "Hours")
        self.assertAlmostEqual(float(row["usage_quantity"]), 404.1014)
        self.assertAlmostEqual(float(row["cost"]), 2877.20)
        self.assertAlmostEqual(float(row["credits"]), -122.80)
        self.assertEqual(row["instance_type"], "r6a.8xlarge")


if __name__ == "__main__":
    unittest.main()

