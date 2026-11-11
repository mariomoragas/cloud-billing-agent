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
- incluir grafico/matriz de complexidade de migracao (top servicos por custo)
- identificar shapes de compute e estimar ganho de CPU para OCI X9, E5 e E6
- gerar uma BoM OCI com part numbers e quantidades tecnicas equivalentes quando o billing tiver granularidade suficiente
- calcular baseline comparavel, tarifas efetivas, cobertura dos dados e drivers de custo
- precificar a BoM com catalogo OCI PAYG versionado e calcular saving por workload, cenarios e sensibilidade
- apontar servicos sem de-para para revisao manual
- gerar analise FinOps/migracao OCI via Gemini API do Google AI Studio (com fallback local quando API indisponivel)

## Requisitos

- Python 3.11

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

Se `GEMINI_API_KEY` nao estiver definida, a aplicacao gera um fallback local sem presumir percentuais de economia, ROI ou payback. O business case deterministico continua disponivel nas abas e slides FinOps.
### Anonimizacao obrigatoria do payload LLM

A aplicacao nunca envia o arquivo de billing bruto nem suas linhas individuais ao Gemini. Antes da
chamada externa, o payload e reduzido a agregados de custo, uso, servicos, regioes e mapeamentos OCI.
Campos de conta, projeto, subscription, tenant, invoice, recurso, tags e nome do arquivo recebem
aliases SHA-256 estaveis. E-mails, UUIDs, ARNs, caminhos de recursos AWS/Azure/GCP, enderecos IP,
billing IDs e sequencias numericas longas sao redigidos dos textos.

A aba `LLM_Resumo` registra `payload_anonymized=true` e `payload_scope=aggregated_only` para
auditoria. A anonimizacao e obrigatoria e nao possui opcao de desativacao pela CLI ou pela Web.

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
- escolher AWS Invoice CSV, AWS Resource CSV, AWS Billing PDF, Azure Cost CSV, GCP Cost table CSV ou CSV generico
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
  --company-name "xxx" `
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
- `aws-resource-csv`: CSV AWS detalhado por recurso, separado por ponto e virgula e com numeros em formato brasileiro
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
- `ClientRef` (fallback para identificar e separar clientes)
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
- usa `ClientRef` como identificador e nome da conta vinculada quando `LinkedAccountId`/`LinkedAccountName` nao estao presentes
- remove por padrao linhas sem identificacao de conta vinculada nem `ClientRef`
- infere regiao a partir do prefixo de `UsageType` quando possivel
- infere unidade de uso com base em `UsageType` e `ItemDescription`

## AWS Resource CSV

O parser `aws-resource-csv` suporta exports detalhados por recurso com colunas como `servico`, `regiao`, `usage_type`, `tipo_instancia`, `resource_id`, `quantidade_uso`, `unidade` e `custo_usd`.

Regras aplicadas:

- detecta valores decimais no formato brasileiro
- preserva conta, recurso, shape, operacao e tags no dado normalizado
- infere `product_code` para os principais servicos AWS usados na BoM OCI
- registra desconto apenas quando existe preco de tabela positivo e superior ao custo faturado
- infere o periodo `AAAA-MM` a partir do nome do arquivo


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

### Template de origem

O PowerPoint e criado a partir do template corporativo versionado em
`app/assets/ppt_template/source_template.pptx`. A aplicacao preserva dimensoes, tema,
slide masters e layouts do arquivo-base. A copia incorporada ao projeto e sanitizada sem
slides de conteudo; o carregador tambem remove slides de qualquer template personalizado antes de
gerar o relatorio. A geracao usa `Light - Blank` como layout padrao, inclusive para capa e divisorias. O slide de Safe Harbor usa o layout corporativo `Light - Safe Harbor`.
O deck final organiza os slides nas secoes `Default Section`, `Comercial` e `Final`; nao e criada uma secao `Anexos`.

Para usar outro template sem alterar o codigo, defina no `.env` um caminho absoluto:

```text
PPT_TEMPLATE_PATH=C:\caminho\para\template.pptx
```

O arquivo configurado deve ser um `.pptx` valido. Se `PPT_TEMPLATE_PATH` nao for
definida, a aplicacao usa o template versionado. Caso o template nao contenha os
layouts nomeados, o gerador procura outro layout com `Blank` no nome e, por fim, usa
o layout de fallback.

Slides incluidos:

- capa
- Safe Harbor statement
- indicadores principais
- top servicos por custo consolidado com baseline comparavel, impostos e maior categoria
- participacao percentual por servico
- top regioes por custo
- matriz de complexidade de migracao dos top servicos por custo (score 1-5)
  - em GCP Cost table, os rotulos usam `Service Description`; o `Service ID` permanece disponivel nos dados tecnicos
  - em AWS Billing PDF, os rotulos continuam usando `product_code`
- performance estimada de CPU para OCI X9, E5 e E6, quando houver shape comparavel
- dimensionamento e part numbers OCI, com quantidades, confianca e premissas
- mapeamento consolidado AWS -> OCI
- saving por tipo de servico
- contribuicao por workload com impactos negativos a esquerda e positivos a direita
- sensibilidade de capacidade GPU e recomendacao de ondas de migracao

- (quando habilitado) secao LLM FinOps com baseline/projecao/ROI e plano de migracao
- servicos sem mapeamento OCI
- encerramento com o layout `Light - Closing Slide`

## Abas geradas no Excel

- `Raw_Data`
- `Resumo_Servicos`
- `Resumo_Regioes`
- `Mapeamento_OCI`
- `Processor_Analysis` (quando shapes de compute forem identificados)
- `Processor_Comparison` (quando houver benchmark x86 comparavel)
- `OCI_BoM` (part numbers OCI e quantidades consolidadas)
- `OCI_BoM_Review` (linhas que exigem dimensionamento ou validacao manual)
- `OCI_BoM_Summary` (cenario, cobertura e versao do catalogo)
- `FinOps_Baseline` (custo faturado, pre-tax, impostos e creditos)
- `FinOps_Categories` (baseline por categoria)
- `Effective_Rates` (tarifa efetiva por SKU e unidade)
- `Analysis_Coverage` (cobertura e confianca dos campos)
- `OCI_Priced_BoM` (quantidades, preco PAYG, free tier e custo OCI estimado)
- `Savings_by_Service` (origem versus OCI por workload)
- `Business_Case` (cenarios e escopo sem preco)
- `Savings_Waterfall` (contribuicao por workload)
- `Sensitivity` (sensibilidade de reducao de capacidade GPU)

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

## Analise de processadores

A aplicacao procura shapes/familias de compute nos campos de SKU e descricao dos parsers AWS, Azure e GCP. As regras ficam em:

- `app/mappings/cloud_shape_processors.csv`: de-para de shape/familia para modelo de processador
- `app/mappings/processor_benchmarks.csv`: catalogo versionado de CPU Mark, Single Thread, samples, URL e data de consulta

A comparacao usa o `CPU Mark Single Thread` do PassMark. O baseline de origem e ponderado pelo custo positivo das linhas de compute com benchmark valido. O ganho e calculado por:

```text
ganho_estimado_pct = (score_single_thread_OCI / score_single_thread_origem - 1) * 100
```

Regras de integridade:

- o resultado e sempre rotulado como estimativa, nao como benchmark da aplicacao
- shapes ambiguos ficam como `UNRESOLVED` e nao recebem ganho numerico
- workloads ARM e GPU nao sao convertidos automaticamente para comparacao x86
- a cobertura de custo e a confianca acompanham o resultado
- poucos samples no PassMark reduzem a confianca
- CPU nao representa memoria, NUMA, rede, storage, GPU ou escalabilidade multicore
- para GCP N2, o baseline usa Cascade Lake como plataforma representativa; N2 tambem pode operar em Ice Lake
- para GCP E2, o baseline conservador usa AMD Milan; E2 tambem pode operar em Intel Broadwell
- inferencias N2/E2 recebem confianca baixa e exigem confirmacao do processador real no inventario ou guest OS

Os processadores OCI de referencia sao confirmados na documentacao de Compute Shapes: X9 (`Intel Xeon Platinum 8358`), E5 (`AMD EPYC 9J14`) e E6 (`AMD EPYC 9J45`). Os benchmarks devem ser revisados periodicamente porque os valores do PassMark mudam com novos samples.

Fontes:

- https://www.cpubenchmark.net/CPU_mega_page.html
- https://docs.oracle.com/en-us/iaas/Content/Compute/References/computeshapes.htm
- https://aws.amazon.com/ec2/instance-types/
- https://learn.microsoft.com/en-us/azure/virtual-machines/sizes/overview
- https://cloud.google.com/compute/docs/cpu-platforms

## Dimensionamento e part numbers OCI

A aplicacao gera uma Bill of Materials tecnica quando o billing possui metricas auditaveis de consumo.

Para AWS Invoice CSV, utiliza:

- horas e capacidade dos shapes EC2 para OCPU-hora, GB-hora e GPU-hora
- EBS gp2/gp3 para Block Volume em GB-mes e VPU-GB-mes
- S3 Standard/Glacier e requisicoes para Object Storage/Archive Storage
- EFS para File Storage
- trafego de saida em GB e horas de Load Balancer

Para GCP Cost table CSV, utiliza SKU description, Usage amount e Usage unit para converter:

- core-hours x86 de familias N1/N2/N2D/N4/E2/C2/C3/C4/M/H/Z em OCPU-hora E5
- memoria provisionada em GB-hora E5
- Balanced Persistent Disk em Block Volume e 10 VPU por GB-mes
- Standard/Archive Cloud Storage e operacoes Class A/B
- Cloud Filestore em File Storage
- Internet Data Transfer Out e horas de Load Balancer Forwarding Rule

Regras de integridade:

- a BoM representa o cenario `Capacity equivalent`; nao aplica rightsizing por utilizacao
- para x86, 1 OCPU OCI equivale a 2 vCPUs de origem; para Arm/Ampere A1, 1 OCPU equivale a 1 vCPU de origem
- cobrancas de reserva, Savings Plans, `HeavyUsage` e GCP Commitments nao viram capacidade para evitar dupla contagem
- a aba `OCI_BoM` permanece tecnica; a precificacao ocorre separadamente em `OCI_Priced_BoM` usando catalogo PAYG versionado
- servicos gerenciados, tiers proprietarios e linhas sem metrica suficiente ficam em `OCI_BoM_Review`
- cada linha inclui confianca, premissa, origem do consumo e versao do catalogo OCI

Catalogos versionados:

- `app/mappings/oci_sku_catalog.csv`: part number, nome, unidade e categoria OCI
- `app/mappings/cloud_shape_capacity.csv`: vCPU, memoria, arquitetura e GPU dos shapes de origem suportados

A fonte dos part numbers e a API publica do OCI Price List. O snapshot incluido nesta versao informa atualizacao em `2026-08-14T09:09:55.872Z`:

- https://apexapps.oracle.com/pls/apex/cetools/api/v1/products/?currencyCode=USD
- https://www.oracle.com/cloud/price-list/

Azure continua gerando a fila de revisao quando o export nao fornece uma regra quantitativa auditavel. No GCP, commitments, servicos gerenciados, snapshots, image storage, tiers proprietarios e SKUs sem equivalencia segura permanecem em OCI_BoM_Review. Novos shapes e servicos devem ser adicionados somente com especificacoes verificadas.
## Analise FinOps deterministica - fases 1 e 2

A fase 1 usa apenas o billing normalizado para gerar baseline auditavel, categorias, tarifas efetivas e cobertura dos dados. Impostos, creditos e custos sem granularidade permanecem separados.

A fase 2 combina somente quantidades auditaveis da BoM com `app/mappings/oci_price_catalog.csv`. O calculo aplica preco PAYG publico, free tiers registrados no catalogo e carry-forward do custo de origem para o escopo sem preco OCI. Nenhuma economia e presumida para linhas sem cobertura.

Cenarios incluidos:

- capacidade equivalente sem rightsizing ou desconto contratual
- retencao temporaria de GPU com migracao dos demais workloads precificados
- sensibilidade de 20% sobre preco OCI, explicitamente marcada como dependente de cotacao
- sensibilidade de reducao de capacidade GPU entre 0% e 50%

O resultado nao representa proposta comercial. Precos, regioes, moeda, tenancy, suporte, licencas e descontos contratuais devem ser validados antes do compromisso com o cliente.

Catalogo de precos:

- `app/mappings/oci_price_catalog.csv`: part number, preco unitario USD, free tier, data e fonte oficial
- https://apexapps.oracle.com/pls/apex/cetools/api/v1/products/?currencyCode=USD
- https://www.oracle.com/cloud/price-list/

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

