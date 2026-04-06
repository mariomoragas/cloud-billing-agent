# Cloud Billing Agent

Esqueleto funcional em Python para:

- ler um CSV de billing de AWS, Azure ou GCP
- disponibilizar uma interface web local com upload de arquivo
- normalizar colunas para um modelo comum
- consolidar custo e quantidade por servico e por regiao
- gerar um Excel com abas de dados, resumo, graficos e mapeamento OCI
- gerar um PowerPoint executivo com slides e graficos principais
- apontar servicos sem de-para para revisao manual

## Requisitos

- Python 3.11 ou superior

## Instalacao

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Uso

```powershell
python -m app.main --input .\seu_billing.csv --cloud aws
```

## Interface web local

Para abrir a interface web com upload do arquivo:

```powershell
python -m app.web
```

Ou no Windows:

```powershell
.\start_web_ui.bat
```

Depois abra:

```text
http://127.0.0.1:8501
```

A interface permite:

- enviar o CSV
- escolher o formato do arquivo
- baixar o Excel pronto no navegador
- baixar o PowerPoint executivo no navegador

## Empacotar como executavel no Windows

Depois de instalar o `PyInstaller`, voce pode gerar um `.exe` com:

```powershell
.\build_web_exe.bat
```

O launcher usado para empacotar fica em `web_launcher.pyw`.

Exemplo com AWS Invoice CSV:

```powershell
python -m app.main `
  --input .\ecsv_4_2024.csv `
  --format aws-invoice `
  --output .\output\billing_report_aws_invoice.xlsx
```

Exemplo com caminhos explicitos:

```powershell
python -m app.main `
  --input C:\caminho\fatura_aws.csv `
  --cloud aws `
  --mapping .\app\mappings\service_mapping.csv `
  --output .\output\billing_report.xlsx
```

## Formatos suportados

- `generic`: CSV com aliases de colunas comuns de billing multicloud
- `aws-invoice`: CSV de invoice consolidada da AWS

## Colunas aceitas

O parser tenta reconhecer aliases comuns para as colunas abaixo:

- `cloud`
- `service_name_original`
- `sku`
- `region`
- `usage_quantity`
- `usage_unit`
- `cost`
- `currency`
- `period`

Obrigatorias:

- `service_name_original`
- `usage_quantity`
- `cost`

## AWS Invoice CSV

O parser dedicado de invoice AWS foi ajustado para colunas reais como:

- `InvoiceID`
- `PayerAccountId`
- `LinkedAccountId`
- `RecordType`
- `PayerAccountName`
- `LinkedAccountName`
- `ProductCode`
- `ProductName`
- `UsageType`
- `Operation`
- `ItemDescription`
- `UsageQuantity`
- `CurrencyCode`
- `CostBeforeTax`
- `TaxAmount`
- `TotalCost`

Regras aplicadas:

- usa `TotalCost` como custo principal para refletir o valor faturado
- preserva `CostBeforeTax`, `TaxAmount` e `Credits` para analise financeira
- filtra para linhas detalhadas `LinkedLineItem` e `PayerLineItem`
- remove por padrao qualquer linha sem `LinkedAccountName`
- infere regiao a partir do prefixo de `UsageType` quando possivel
- infere unidade de uso com base em `UsageType` e `ItemDescription`

## PowerPoint executivo

Toda execucao agora gera tambem um `.pptx` ao lado do arquivo Excel.

Slides incluidos:

- capa
- indicadores principais
- top servicos por custo
- participacao percentual por servico
- top regioes por custo
- servicos sem mapeamento OCI

## Abas geradas no Excel

- `Raw_Data`
- `Resumo_Servicos`
- `Resumo_Regioes`
- `Mapeamento_OCI`
- `Pendencias`
- `Charts`

## Proximos passos sugeridos

- adicionar parser dedicado para Azure Cost Export
- adicionar parser dedicado para GCP Billing Export
- adicionar leitura de PDF com `pdfplumber`
- integrar um LLM apenas para sugerir mapeamentos OCI em casos nao cobertos
