from __future__ import annotations

import json
import os
import re
import urllib.error
import urllib.request
from dataclasses import dataclass

import pandas as pd

SYSTEM_PROMPT = """
Você é um Especialista em FinOps e Migração Cloud focado em transformar exports de billing/costs de AWS/Azure/GCP em uma projeção auditável de custo na Oracle Cloud Infrastructure (OCI), com plano de migração, economia projetada, ROI e recomendações de arquitetura.

1) OBJETIVO (o que você deve entregar)
Gerar um relatório para vendas/arquitetura demonstrando:
- Custo atual (baseline) baseado no billing fornecido
- Projeção de custo em OCI por serviço/categoria e total
- Economia projetada (meta: identificar caminhos para 30%+, quando suportado pelos dados)
- ROI, payback e business case
- Plano de migração (fases, riscos, dependências, quick wins)
- Recomendações de arquitetura OCI alinhadas aos workloads e ao padrão de consumo observado

2) ENTRADAS (o que você receberá)
Você pode receber um ou mais dos itens abaixo:
- Export de billing/costs em PDF (fatura), CSV, XLSX, ou "Cost by service"
- Metadados do período (mês, moeda), contas/subscrições, regiões
- (Opcional) inventário/arquitetura atual, tags, contas, dimensionamento, SLAs, ambientes (prod/dev)

3) REGRAS CRÍTICAS (anti-hallucination / integridade)

*REGRA 1 — Se houver arquivo de billing (PDF/CSV/XLSX):*
Use APENAS valores exatos presentes no arquivo para o baseline atual.
Extraia categorias reais (serviço, região, usage type, subscription, resource group, etc.) conforme disponível.
Sempre que possível, cite a origem do número: "conforme billing fornecido" + referência contextual (ex.: seção/linha/aba/categoria).
Não invente breakdowns que não existam no export.

*REGRA 2 — Se NÃO houver arquivo (apenas descrição manual):*
- Você pode estimar, mas deve marcar claramente como ESTIMADO e listar assunções.
- Não misture estimativa com dado real sem rotular.

*REGRA 3 — Preços OCI e "atualização":*
Você não pode afirmar que usou "preços OCI atualizados" sem:
a) o usuário fornecer uma tabela/print/CSV de preços OCI, ou
b) haver uma fonte explícita disponibilizada nas entradas.

Se preços OCI não forem fornecidos, faça projeções com:
- assunções explícitas (ex.: modelo on-demand, regiões comparáveis, sem promoções), ou
- placeholders do tipo [PREÇO OCI A VALIDAR] quando a precisão exigir validação.

Nunca "chute" valores unitários sem declarar assunção e nível de confiança.

*REGRA 4 — Comparabilidade e escopo:*
Separe claramente:
- Compute, Storage, Network/Egress, Database, Managed services, Support, Taxes, Credits/Discounts.

Impostos, taxas e créditos devem ser tratados à parte, pois variam (não assuma equivalência automática entre clouds).

4) NORMALIZAÇÃO (como tratar os dados antes de projetar)
Ao receber o export, normalize:
- Período (mensal, diário), moeda, contas/subscrições, região
- Custos recorrentes vs. one-off
- Descontos e programas:
  - AWS: Savings Plans / Reserved Instances / EDP (se aparecer)
  - Azure: Reservations / Savings Plan / EA (se aparecer)
  - GCP: Committed Use / Sustained Use (se aparecer)
- Créditos (promocionais, marketplace) e taxas (suporte, impostos) em linhas separadas
- Identifique custos "difíceis de migrar" (ex.: SaaS/Marketplace, licenças, serviços muito proprietários)

5) MÉTODO DE PROJEÇÃO PARA OCI (obrigatório)
Para cada grande categoria do billing atual:
- Identifique o "equivalente funcional" em OCI (ou alternativa recomendada)
- Defina o modelo de destino: Rehost / Replatform / Refactor (com justificativa)
- Escolha o método de custo:
  - Se houver dados suficientes de consumo (vCPU/hora, GB/mês, requests, GB egress): projeção quantitativa
  - Se houver apenas custo agregado por serviço: projeção por mapeamento + faixas com confiança e assunções
- Aponte alavancas de economia (ex.: rightsizing, eliminação de desperdícios, ajuste de storage tier, otimização de egress, mudança de DB)
- Gere 3 cenários: Conservador / Base / Agressivo (com o que muda em cada)

6) IDENTIFICAÇÃO DE DESPERDÍCIOS (FinOps)
Com base no billing, procure e destaque:
- Serviços com custo alto e baixa evidência de valor (ex.: snapshots excessivos, storage frio/quente mal posicionado, IPs/volumes órfãos, serviços duplicados)
- Padrões de custo por ambiente (prod vs dev) se houver tags/contas
- Custo de egress e inter-region (normalmente grande vilão de ROI)
- Custos de suporte/planos (se presentes)

7) PADRÕES DE CONFIANÇA (obrigatório)
Rotule cada bloco crítico como:
- Confiança Alta: quando há consumo detalhado e mapeamento direto
- Confiança Média: quando há custo por serviço sem métricas granulares
- Confiança Baixa: quando só existe total agregado e muitas variáveis (ex.: egress sem volume)

8) LINGUAGEM E POSTURA
- Seja direto, numérico e auditável.
- Não prometa 30%+ se os dados não suportarem; em vez disso: mostre quais alavancas poderiam levar a 30% e o que precisa ser validado.
- Evite termos vagos ("otimizar bastante"); sempre diga o que e onde.
""".strip()

JSON_INSTRUCTION = """
Responda somente em JSON válido com este schema:
{
  "executive_summary": "string",
  "baseline": {
    "current_total_cost": number,
    "currency": "string",
    "evidence": "string"
  },
  "oci_projection": {
    "conservative_total": number,
    "base_total": number,
    "aggressive_total": number,
    "method_notes": "string"
  },
  "savings": {
    "base_savings_pct": number,
    "base_savings_value": number,
    "path_to_30_plus": "string"
  },
  "business_case": {
    "roi_pct": number,
    "payback_months": number,
    "migration_investment": number,
    "assumptions": "string"
  },
  "migration_plan": [
    {
      "phase": "string",
      "duration": "string",
      "activities": "string",
      "risks": "string",
      "dependencies": "string",
      "quick_wins": "string"
    }
  ],
  "architecture_recommendations": ["string"],
  "finops_findings": ["string"],
  "confidence": [
    {
      "topic": "string",
      "level": "Alta|Média|Baixa",
      "reason": "string"
    }
  ],
  "assumptions": ["string"],
  "citations": ["string"]
}
""".strip()


@dataclass
class LLMReportArtifacts:
    summary_df: pd.DataFrame
    migration_df: pd.DataFrame
    recommendations_df: pd.DataFrame
    confidence_df: pd.DataFrame
    status: str
    executive_summary: str
    error_message: str = ""


def build_llm_report_artifacts(
    *,
    raw_df: pd.DataFrame,
    service_summary_df: pd.DataFrame,
    region_summary_df: pd.DataFrame,
    oci_mapping_df: pd.DataFrame,
    source_name: str,
    llm_model: str,
) -> LLMReportArtifacts:
    baseline_total = float(raw_df["cost"].sum()) if not raw_df.empty else 0.0
    currency = "USD"
    if not raw_df.empty and "currency" in raw_df.columns:
        currency = str(raw_df["currency"].fillna("USD").astype(str).mode().iloc[0])

    payload = _build_payload(
        raw_df=raw_df,
        service_summary_df=service_summary_df,
        region_summary_df=region_summary_df,
        oci_mapping_df=oci_mapping_df,
        source_name=source_name,
        baseline_total=baseline_total,
        currency=currency,
    )

    report, request_error = _request_llm_report(payload=payload, llm_model=llm_model)
    error_message = ""
    if report is None:
        report = _fallback_report(payload=payload)
        if not os.getenv("GEMINI_API_KEY", "").strip():
            status = "fallback_missing_api_key"
            error_message = "GEMINI_API_KEY ausente."
        else:
            status = "fallback_api_error"
            error_message = request_error or "Falha ao chamar Gemini API."
    else:
        status = "llm_gemini"

    summary_df = _build_summary_df(report, status=status, error_message=error_message)
    migration_df = _build_migration_df(report)
    recommendations_df = _build_recommendations_df(report)
    confidence_df = _build_confidence_df(report)
    executive_summary = str(report.get("executive_summary", "")).strip()

    return LLMReportArtifacts(
        summary_df=summary_df,
        migration_df=migration_df,
        recommendations_df=recommendations_df,
        confidence_df=confidence_df,
        status=status,
        executive_summary=executive_summary,
        error_message=error_message,
    )


def _build_payload(
    *,
    raw_df: pd.DataFrame,
    service_summary_df: pd.DataFrame,
    region_summary_df: pd.DataFrame,
    oci_mapping_df: pd.DataFrame,
    source_name: str,
    baseline_total: float,
    currency: str,
) -> dict[str, object]:
    mapped_df = oci_mapping_df.copy()
    if not mapped_df.empty and "oci_service" in mapped_df.columns:
        mapped_df["oci_service"] = mapped_df["oci_service"].fillna("").astype(str)
        mapped_df = mapped_df[mapped_df["oci_service"] != "REVIEW_REQUIRED"]

    oci_by_service = []
    if not mapped_df.empty:
        grouped = (
            mapped_df.groupby("oci_service", as_index=False)["total_cost"]
            .sum()
            .sort_values("total_cost", ascending=False)
            .head(20)
        )
        for _, row in grouped.iterrows():
            oci_by_service.append(
                {"oci_service": str(row["oci_service"]), "total_cost": float(row["total_cost"])}
            )

    unmapped = []
    if not oci_mapping_df.empty:
        unmapped_df = oci_mapping_df[oci_mapping_df["oci_service"] == "REVIEW_REQUIRED"].copy()
        if not unmapped_df.empty:
            unmapped_df = unmapped_df.sort_values("total_cost", ascending=False).head(20)
            for _, row in unmapped_df.iterrows():
                unmapped.append(
                    {
                        "service_name_original": str(row["service_name_original"]),
                        "total_cost": float(row["total_cost"]),
                    }
                )

    service_top = []
    for _, row in service_summary_df.head(25).iterrows():
        service_top.append(
            {
                "service_name_original": str(row["service_name_original"]),
                "total_cost": float(row["total_cost"]),
                "total_usage_quantity": float(row.get("total_usage_quantity", 0.0)),
                "primary_unit": str(row.get("primary_unit", "")),
            }
        )

    region_top = []
    for _, row in region_summary_df.head(15).iterrows():
        region_top.append(
            {
                "region": str(row.get("region", "")),
                "total_cost": float(row["total_cost"]),
            }
        )

    return {
        "source_name": source_name,
        "baseline_total_cost": baseline_total,
        "currency": currency,
        "row_count": int(len(raw_df)),
        "service_count": int(raw_df["service_name_original"].nunique()) if not raw_df.empty else 0,
        "period_mode": _mode_or_default(raw_df, "period", ""),
        "cloud_mode": _mode_or_default(raw_df, "cloud", ""),
        "top_services": service_top,
        "top_regions": region_top,
        "oci_mapped_cost_by_service": oci_by_service,
        "unmapped_services": unmapped,
    }


def _request_llm_report(
    payload: dict[str, object],
    llm_model: str,
) -> tuple[dict[str, object] | None, str]:
    api_key = os.getenv("GEMINI_API_KEY", "").strip()
    if not api_key:
        return None, "GEMINI_API_KEY ausente."

    model = llm_model.strip() or os.getenv("GEMINI_MODEL", "gemini-2.5-flash")
    endpoint = _resolve_gemini_endpoint(model)

    user_message = (
        "Dados de billing normalizados para análise FinOps/OCI:\n"
        + json.dumps(payload, ensure_ascii=False, indent=2)
        + "\n\n"
        + JSON_INSTRUCTION
    )
    body = {
        "systemInstruction": {
            "parts": [{"text": SYSTEM_PROMPT}],
        },
        "contents": [
            {
                "role": "user",
                "parts": [{"text": user_message}],
            },
        ],
        "generationConfig": {
            "temperature": 0.2,
            "responseMimeType": "application/json",
        },
    }

    data = json.dumps(body).encode("utf-8")
    request = urllib.request.Request(
        endpoint,
        data=data,
        method="POST",
        headers={
            "x-goog-api-key": api_key,
            "Content-Type": "application/json",
        },
    )
    try:
        with urllib.request.urlopen(request, timeout=90) as response:
            response_json = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        details = ""
        try:
            details = exc.read().decode("utf-8")
        except Exception:
            details = ""
        reason = _extract_gemini_error(details) or f"HTTP {exc.code}"
        return None, reason
    except urllib.error.URLError as exc:
        return None, f"URLError: {exc.reason}"
    except TimeoutError:
        return None, "Timeout na chamada Gemini API."
    except json.JSONDecodeError:
        return None, "Resposta da Gemini API em formato invalido."

    try:
        parts = response_json["candidates"][0]["content"]["parts"]
        content = "".join(str(part.get("text", "")) for part in parts)
    except (KeyError, IndexError, TypeError):
        return None, "Resposta da Gemini API sem candidates[0].content.parts[].text."

    parsed = _parse_json_text(content)
    if parsed is None:
        return None, "Resposta da LLM nao esta em JSON valido."
    return parsed, ""


def _resolve_gemini_endpoint(model: str) -> str:
    endpoint_template = os.getenv("GEMINI_API_ENDPOINT", "").strip()
    if endpoint_template:
        return endpoint_template.replace("{model}", model)
    return (
        "https://generativelanguage.googleapis.com/v1beta/models/"
        f"{model}:generateContent"
    )


def _parse_json_text(content: str) -> dict[str, object] | None:
    content = content.strip()
    if not content:
        return None
    if content.startswith("```"):
        content = re.sub(r"^```(?:json)?\s*", "", content)
        content = re.sub(r"\s*```$", "", content)
    try:
        return json.loads(content)
    except json.JSONDecodeError:
        return None


def _extract_gemini_error(details: str) -> str:
    details = (details or "").strip()
    if not details:
        return ""
    try:
        payload = json.loads(details)
    except json.JSONDecodeError:
        return details[:200]
    error_obj = payload.get("error", {}) if isinstance(payload, dict) else {}
    message = error_obj.get("message", "") if isinstance(error_obj, dict) else ""
    status = error_obj.get("status", "") if isinstance(error_obj, dict) else ""
    code = error_obj.get("code", "") if isinstance(error_obj, dict) else ""
    if message and status:
        return f"{message} (status={status})"
    if message and code:
        return f"{message} (code={code})"
    if message:
        return str(message)
    return details[:200]


def _fallback_report(payload: dict[str, object]) -> dict[str, object]:
    baseline = float(payload.get("baseline_total_cost", 0.0))
    conservative_total = baseline * 0.9
    base_total = baseline * 0.8
    aggressive_total = baseline * 0.7
    base_savings_value = baseline - base_total
    migration_investment = baseline * 0.15
    roi_pct = ((base_savings_value - migration_investment) / migration_investment * 100.0) if migration_investment > 0 else 0.0
    payback_months = (migration_investment / max(base_savings_value, 1e-9)) if base_savings_value > 0 else 0.0

    return {
        "executive_summary": (
            "Relatório gerado sem chamada externa de LLM (fallback local). "
            "Os valores de baseline são exatos do billing fornecido; projeções OCI estão marcadas como estimadas."
        ),
        "baseline": {
            "current_total_cost": baseline,
            "currency": str(payload.get("currency", "USD")),
            "evidence": "Conforme billing fornecido (abas Raw_Data/Resumo_Servicos).",
        },
        "oci_projection": {
            "conservative_total": conservative_total,
            "base_total": base_total,
            "aggressive_total": aggressive_total,
            "method_notes": "Estimativa por faixas sem tabela de preços OCI fornecida.",
        },
        "savings": {
            "base_savings_pct": 20.0,
            "base_savings_value": base_savings_value,
            "path_to_30_plus": "Para 30%+, validar rightsizing, storage tiering, egress e refatoração de serviços gerenciados.",
        },
        "business_case": {
            "roi_pct": roi_pct,
            "payback_months": payback_months,
            "migration_investment": migration_investment,
            "assumptions": "Capex/opex de migração estimado em 15% do baseline mensal.",
        },
        "migration_plan": [
            {
                "phase": "Fase 1 - Descoberta e landing zone",
                "duration": "2-4 semanas",
                "activities": "Validar inventário, tagging, segurança, conectividade e guardrails em OCI.",
                "risks": "Inconsistência de inventário e dependências ocultas.",
                "dependencies": "Acesso aos tenants, contas e owners técnicos.",
                "quick_wins": "Mapear top 20 serviços por custo e priorizar equivalentes OCI.",
            },
            {
                "phase": "Fase 2 - Migração de workloads prioritários",
                "duration": "4-8 semanas",
                "activities": "Migrar workloads com maior custo e menor complexidade.",
                "risks": "Janelas de mudança e regressão de performance.",
                "dependencies": "Runbooks, testes e observabilidade.",
                "quick_wins": "Compute rehost com rightsizing e storage tier adequado.",
            },
            {
                "phase": "Fase 3 - Otimização contínua FinOps",
                "duration": "Contínuo",
                "activities": "Ajustar sizing, políticas de lifecycle e governança de custos.",
                "risks": "Recaída de desperdício sem disciplina de operação.",
                "dependencies": "KPIs e cadência mensal de revisão.",
                "quick_wins": "Ações em egress, snapshots e recursos ociosos.",
            },
        ],
        "architecture_recommendations": [
            "Priorizar rehost de compute e bancos mais aderentes ao mapeamento OCI já identificado.",
            "Aplicar políticas de lifecycle em Object Storage/Archive e revisão de snapshots.",
            "Revisar padrões de tráfego inter-região para reduzir egress.",
        ],
        "finops_findings": [
            "Top serviços concentram a maior parte do custo mensal; priorizar ondas de migração por impacto.",
            "Custos de rede/egress podem impactar ROI se não forem tratados no desenho alvo.",
        ],
        "confidence": [
            {
                "topic": "Baseline",
                "level": "Alta",
                "reason": "Valores extraídos diretamente do billing processado.",
            },
            {
                "topic": "Projeção OCI",
                "level": "Média",
                "reason": "Sem tabela explícita de preços OCI nas entradas.",
            },
        ],
        "assumptions": [
            "Sem preços OCI fornecidos, projeção feita por faixas conservador/base/agressivo.",
            "Impostos, créditos e descontos específicos tratados separadamente quando visíveis no billing.",
        ],
        "citations": [
            "Conforme billing fornecido: abas Raw_Data, Resumo_Servicos, Resumo_Regioes, Mapeamento_OCI.",
        ],
    }


def _build_summary_df(
    report: dict[str, object],
    *,
    status: str,
    error_message: str,
) -> pd.DataFrame:
    baseline = report.get("baseline", {}) if isinstance(report.get("baseline"), dict) else {}
    projection = report.get("oci_projection", {}) if isinstance(report.get("oci_projection"), dict) else {}
    savings = report.get("savings", {}) if isinstance(report.get("savings"), dict) else {}
    business = report.get("business_case", {}) if isinstance(report.get("business_case"), dict) else {}

    rows = [
        {"section": "meta", "metric": "analysis_mode", "value": status},
        {"section": "meta", "metric": "analysis_error", "value": error_message},
        {"section": "summary", "metric": "executive_summary", "value": report.get("executive_summary", "")},
        {"section": "baseline", "metric": "current_total_cost", "value": baseline.get("current_total_cost", 0.0)},
        {"section": "baseline", "metric": "currency", "value": baseline.get("currency", "USD")},
        {"section": "baseline", "metric": "evidence", "value": baseline.get("evidence", "")},
        {"section": "projection", "metric": "conservative_total", "value": projection.get("conservative_total", 0.0)},
        {"section": "projection", "metric": "base_total", "value": projection.get("base_total", 0.0)},
        {"section": "projection", "metric": "aggressive_total", "value": projection.get("aggressive_total", 0.0)},
        {"section": "projection", "metric": "method_notes", "value": projection.get("method_notes", "")},
        {"section": "savings", "metric": "base_savings_pct", "value": savings.get("base_savings_pct", 0.0)},
        {"section": "savings", "metric": "base_savings_value", "value": savings.get("base_savings_value", 0.0)},
        {"section": "savings", "metric": "path_to_30_plus", "value": savings.get("path_to_30_plus", "")},
        {"section": "business_case", "metric": "roi_pct", "value": business.get("roi_pct", 0.0)},
        {"section": "business_case", "metric": "payback_months", "value": business.get("payback_months", 0.0)},
        {"section": "business_case", "metric": "migration_investment", "value": business.get("migration_investment", 0.0)},
        {"section": "business_case", "metric": "assumptions", "value": business.get("assumptions", "")},
    ]
    assumptions = report.get("assumptions", [])
    if isinstance(assumptions, list):
        for item in assumptions:
            rows.append({"section": "assumptions", "metric": "item", "value": str(item)})
    citations = report.get("citations", [])
    if isinstance(citations, list):
        for item in citations:
            rows.append({"section": "citations", "metric": "item", "value": str(item)})
    findings = report.get("finops_findings", [])
    if isinstance(findings, list):
        for item in findings:
            rows.append({"section": "finops_findings", "metric": "item", "value": str(item)})
    return pd.DataFrame(rows)


def _build_migration_df(report: dict[str, object]) -> pd.DataFrame:
    plan = report.get("migration_plan", [])
    if not isinstance(plan, list) or not plan:
        return pd.DataFrame([{"phase": "N/A", "duration": "", "activities": "", "risks": "", "dependencies": "", "quick_wins": ""}])
    rows = []
    for item in plan:
        if not isinstance(item, dict):
            continue
        rows.append(
            {
                "phase": item.get("phase", ""),
                "duration": item.get("duration", ""),
                "activities": item.get("activities", ""),
                "risks": item.get("risks", ""),
                "dependencies": item.get("dependencies", ""),
                "quick_wins": item.get("quick_wins", ""),
            }
        )
    if not rows:
        rows = [{"phase": "N/A", "duration": "", "activities": "", "risks": "", "dependencies": "", "quick_wins": ""}]
    return pd.DataFrame(rows)


def _build_recommendations_df(report: dict[str, object]) -> pd.DataFrame:
    rows: list[dict[str, str]] = []
    for item in report.get("architecture_recommendations", []) if isinstance(report.get("architecture_recommendations"), list) else []:
        rows.append({"type": "architecture_recommendation", "item": str(item)})
    for item in report.get("finops_findings", []) if isinstance(report.get("finops_findings"), list) else []:
        rows.append({"type": "finops_finding", "item": str(item)})
    if not rows:
        rows = [{"type": "note", "item": "Sem recomendações retornadas pelo modelo."}]
    return pd.DataFrame(rows)


def _build_confidence_df(report: dict[str, object]) -> pd.DataFrame:
    confidence = report.get("confidence", [])
    if not isinstance(confidence, list) or not confidence:
        return pd.DataFrame([{"topic": "Overall", "level": "Baixa", "reason": "Sem classificação retornada."}])
    rows = []
    for item in confidence:
        if not isinstance(item, dict):
            continue
        rows.append(
            {
                "topic": item.get("topic", ""),
                "level": item.get("level", ""),
                "reason": item.get("reason", ""),
            }
        )
    if not rows:
        rows = [{"topic": "Overall", "level": "Baixa", "reason": "Sem classificação retornada."}]
    return pd.DataFrame(rows)


def _mode_or_default(df: pd.DataFrame, column: str, default: str) -> str:
    if df.empty or column not in df.columns:
        return default
    series = df[column].fillna("").astype(str).str.strip()
    series = series[series != ""]
    if series.empty:
        return default
    return str(series.mode().iloc[0])
