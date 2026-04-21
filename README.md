# Cloud Billing Agent

Esqueleto funcional em Python para:

- ler billing em CSV (multicloud) e PDF (AWS Billing and Cost Management)
- ler CSV real de Azure Cost exportado em tabela de metricas
- ler CSV real de GCP Cost table exportado pelo console de billing
- disponibilizar uma interface web local com upload de arquivo
- normalizar colunas para um modelo comum
- consolidar custo e quantidade por servico e por regiao
- gerar um Excel com abas de dados, resumo, graficos e mapeamento OCI
- gerar um PowerPoint executivo com slides e graficos principais
- apontar servicos sem de-para para revisao manual
- gerar analise FinOps/migracao OCI via Gemini API do Google AI Studio (com fallback local quando API indisponivel)

## Requisitos

- Python 3.11 ou superior

## Instalacao

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

No Windows, se voce nao ativar a virtualenv, use sempre o Python da `.venv`:

```powershell
.venv\Scripts\python.exe -m app.main --help
```

## Uso

```powershell
.venv\Scripts\python.exe -m app.main --input .\seu_billing.csv --cloud aws
```

Para incluir analise FinOps via Gemini API no relatorio:

```powershell
$env:GEMINI_API_KEY="sua_chave_google_ai_studio"
.venv\Scripts\python.exe -m app.main --input .\seu_billing.csv --cloud aws --llm-model gemini-2.5-flash
```

Se `GEMINI_API_KEY` nao estiver definida, a aplicacao gera um fallback local (estimativo)
e ainda preenche as abas/slides de analise LLM.

### Configuracao via arquivo `.env` (CLI e Web)

Voce pode criar um arquivo `.env` na raiz do projeto com:

```text
GEMINI_API_KEY=sua_chave_google_ai_studio
GEMINI_MODEL=gemini-2.5-flash
```

A aplicacao carrega esse arquivo automaticamente ao iniciar (`python -m app.main` e `python -m app.web`).
O arquivo `.env` configura apenas variaveis como `GEMINI_API_KEY` e `GEMINI_MODEL`; dependencias
Python como `python-pptx` devem estar instaladas na `.venv` via `requirements.txt`.

## Interface web local

Para abrir a interface web com upload do arquivo:

```powershell
.venv\Scripts\python.exe -m app.web
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

- enviar CSV ou PDF
- limitar upload de CSV/PDF a 300MB por arquivo
- escolher o formato do arquivo
- escolher AWS Invoice CSV, AWS Billing PDF, Azure Cost CSV, GCP Cost table CSV ou CSV generico
- baixar o Excel pronto no navegador
- baixar o PowerPoint executivo no navegador
- marcar opcao para apagar os arquivos gerados apos o download (Excel/PPT)
- manter relatorios em memoria por ate 5 dias (TTL), com limpeza automatica de expirados
- exibir estado `Processando...` e bloquear envio duplo do formulario
- visualizar previa com custo total, top servicos e servicos sem mapeamento OCI
- validar conteudo real do arquivo (assinatura PDF e estrutura CSV), nao apenas extensao

## Empacotar como executavel no Windows

Depois de instalar o `PyInstaller`, voce pode gerar um `.exe` com:

```powershell
.\build_web_exe.bat
```

O launcher usado para empacotar fica em `web_launcher.pyw`.

Exemplo com AWS Invoice CSV:

```powershell
.venv\Scripts\python.exe -m app.main `
  --input .\ecsv_4_2024.csv `
  --format aws-invoice `
  --output .\output\billing_report_aws_invoice.xlsx
```

Exemplo com AWS Billing PDF:

```powershell
.venv\Scripts\python.exe -m app.main `
  --input "C:\caminho\billing_aws.pdf" `
  --format aws-billing-pdf `
  --output .\output\billing_report_aws_pdf.xlsx
```

Exemplo com GCP Cost table CSV:

```powershell
.venv\Scripts\python.exe -m app.main `
  --input "C:\caminho\Minha conta de faturamento_Cost table, 2025-02-01 — 2025-02-28.csv" `
  --format gcp-cost-table `
  --cloud gcp `
  --output .\output\billing_report_gcp_cost_table.xlsx `
  --company-name "Minha conta de faturamento" `
  --project-name "GCP to OCI Assessment"
```

Exemplo com Azure Cost CSV:

```powershell
.venv\Scripts\python.exe -m app.main `
  --input "C:\caminho\tabela_metricas.csv" `
  --format azure-cost-csv `
  --cloud azure `
  --output .\output\billing_report_azure_cost.xlsx `
  --company-name "Maida Health" `
  --project-name "Azure to OCI Assessment"
```

Exemplo com caminhos explicitos:

```powershell
.venv\Scripts\python.exe -m app.main `
  --input C:\caminho\fatura_aws.csv `
  --cloud aws `
  --mapping .\app\mappings\service_mapping.csv `
  --output .\output\billing_report.xlsx
```

## Formatos suportados

- `generic`: CSV com aliases de colunas comuns de billing multicloud
- `aws-invoice`: CSV de invoice consolidada da AWS
- `aws-billing-pdf`: PDF de billing consolidado da AWS Billing and Cost Management
- `azure-cost-csv`: CSV Azure em formato de tabela de metricas com colunas em portugues
- `gcp-cost-table`: CSV de GCP Cost table exportado pelo console de billing

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

## GCP Cost table CSV

O parser dedicado de GCP foi criado a partir de um CSV real exportado do Cost table.
Ele reconhece o layout com metadados no topo, por exemplo:

- `Invoice number`
- `Invoice date`
- `Due date`
- `Billing ID`
- `Billing account ID`
- `Currency`
- `Currency exchange rate`
- `Total amount due`

Depois localiza automaticamente o cabecalho real iniciado por `Billing account name` e usa colunas como:

- `Billing account name`
- `Billing account ID`
- `Project name`
- `Project ID`
- `Project hierarchy`
- `Service description`
- `Service ID`
- `SKU description`
- `SKU ID`
- `Credit type`
- `Cost type`
- `Usage start date`
- `Usage end date`
- `Usage amount`
- `Usage unit`
- `Unrounded Cost (R$)`
- `Cost (R$)`

Regras aplicadas:

- ignora linhas de metadados antes do cabecalho real
- converte numeros com formato brasileiro, como `45.625,35` e `0,003`
- identifica a moeda pelos metadados (`Currency`) ou pelo nome da coluna de custo
- usa apenas linhas `Cost type = Usage` com `Service description` preenchido na analise por servico
- remove linhas de `Tax`, `Total` e `Rounding error` da analise operacional
- preserva projeto, SKU, tipo de credito, datas de uso e custo arredondado/nao arredondado
- cria reconciliacao na aba `Data_Quality` com delta entre a linha `Total` e o custo analisado

## Azure Cost CSV

O parser dedicado de Azure foi criado a partir de um CSV real de tabela de metricas.
Ele reconhece colunas como:

- `Instancia`
- `Grupo`
- `Categoria`
- `Subcategoria`
- `Nome do Produto`
- `Nome do recurso`
- `Local`
- `QTD`
- `UN`
- `Data`
- `Consumo`

Regras aplicadas:

- detecta delimitador `,` ou `;` e encoding (`utf-8-sig`/`latin-1`)
- normaliza nomes de colunas com acentos para o modelo interno
- usa `Categoria` como servico principal para mapeamento OCI
- usa `Nome do Produto` como SKU
- converte `Data` para periodo mensal (`YYYY-MM`)
- preserva colunas auxiliares como grupo, recurso, instancia e subcategoria

## PowerPoint executivo

Toda execucao agora gera tambem um `.pptx` ao lado do arquivo Excel.

Slides incluidos:

- capa
- indicadores principais
- top servicos por custo
- participacao percentual por servico
- top regioes por custo
- mapeamento consolidado AWS -> OCI
- (quando habilitado) secao LLM FinOps com baseline/projecao/ROI e plano de migracao
- servicos sem mapeamento OCI

## Abas geradas no Excel

- `Raw_Data`
- `Resumo_Servicos`
- `Resumo_Regioes`
- `Mapeamento_OCI`
- `Data_Quality`
- `LLM_Resumo`
- `LLM_Migracao`
- `LLM_Recomendacoes`
- `LLM_Confianca`
- `Pendencias`
- `Charts`

Para GCP Cost table, tambem podem ser geradas abas auxiliares:

- `gcp_project_name`
- `gcp_sku`
- `gcp_credit_type`

## Analise LLM no PowerPoint

Quando habilitada, o deck inclui slides adicionais com:

- baseline e projecoes OCI (conservador/base/agressivo)
- economia, ROI e payback
- plano de migracao em fases
- recomendacoes de arquitetura e achados FinOps

## Proximos passos sugeridos

- adicionar suporte ao GCP BigQuery Billing Export detalhado
- permitir escolha de provedor LLM e endpoint via UI
- enriquecer parser PDF AWS para mais variacoes de layout/idioma
