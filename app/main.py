from __future__ import annotations

import argparse
import os
from pathlib import Path

from app.config import load_local_config
from app.pipeline import process_billing_file


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Processa billing (CSV/PDF) e gera Excel com graficos e mapeamento OCI."
    )
    parser.add_argument("--input", required=True, help="Caminho do arquivo de billing (CSV/PDF).")
    parser.add_argument(
        "--cloud",
        default="aws",
        choices=["aws", "azure", "gcp"],
        help="Cloud de origem quando o CSV nao informa explicitamente.",
    )
    parser.add_argument(
        "--format",
        default="generic",
        choices=["generic", "aws-invoice", "aws-billing-pdf", "gcp-cost-table"],
        help="Formato do arquivo de entrada.",
    )
    parser.add_argument(
        "--mapping",
        default="app/mappings/service_mapping.csv",
        help="Caminho do CSV com regras de mapeamento para OCI.",
    )
    parser.add_argument(
        "--output",
        default="output/billing_report.xlsx",
        help="Caminho do arquivo Excel de saida.",
    )
    parser.add_argument(
        "--company-name",
        default="",
        help="Nome da empresa para exibir na capa do PowerPoint.",
    )
    parser.add_argument(
        "--project-name",
        default="",
        help="Nome do projeto ou assessment para exibir na capa do PowerPoint.",
    )
    parser.add_argument(
        "--llm-model",
        default=os.getenv("OPENAI_MODEL", "gpt-4o-mini"),
        help="Modelo OpenAI usado para analise FinOps via LLM.",
    )
    return parser


def main() -> None:
    load_local_config()
    parser = build_parser()
    args = parser.parse_args()

    input_path = Path(args.input).resolve()
    mapping_path = Path(args.mapping).resolve()
    output_path = Path(args.output).resolve()
    presentation_path = output_path.with_suffix(".pptx")

    result = process_billing_file(
        input_path=input_path,
        output_path=output_path,
        presentation_path=presentation_path,
        file_format=args.format,
        cloud=args.cloud,
        mapping_path=mapping_path,
        company_name=args.company_name,
        project_name=args.project_name,
        llm_model=args.llm_model,
    )

    print(f"Relatorio gerado com sucesso em: {result.output_path}")
    if result.presentation_path is not None:
        print(f"Apresentacao gerada com sucesso em: {result.presentation_path}")


if __name__ == "__main__":
    main()
