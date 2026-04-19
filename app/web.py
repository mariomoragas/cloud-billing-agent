from __future__ import annotations

import cgi
import html
import os
import tempfile
import traceback
import uuid
import webbrowser
from pathlib import Path
from typing import Iterable
from wsgiref.simple_server import make_server

from app.config import load_local_config
from app.pipeline import process_billing_file
from app.static import guess_content_type, resolve_font_asset

HOST = "127.0.0.1"
PORT = 8501
APP_ROOT = Path(__file__).resolve().parent.parent
MAPPING_PATH = APP_ROOT / "app" / "mappings" / "service_mapping.csv"
WEB_OUTPUT_DIR = APP_ROOT / "output" / "web_downloads"
WEB_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
REPORTS: dict[str, dict[str, object]] = {}


def application(environ, start_response):
    method = environ.get("REQUEST_METHOD", "GET").upper()
    path = environ.get("PATH_INFO", "/")

    try:
        if method == "GET" and path == "/":
            return _html_response(start_response, _render_home())
        if method == "POST" and path == "/process":
            return _handle_process(environ, start_response)
        if method == "GET" and path.startswith("/download/"):
            report_id = path.removeprefix("/download/")
            return _handle_download(report_id, start_response)
        if method == "GET" and path.startswith("/download-ppt/"):
            report_id = path.removeprefix("/download-ppt/")
            return _handle_download_ppt(report_id, start_response)
        if method == "GET" and path.startswith("/assets/fonts/"):
            font_name = path.removeprefix("/assets/fonts/")
            return _handle_font(font_name, start_response)
        if method == "GET" and path == "/health":
            return _text_response(start_response, "ok")
        return _html_response(
            start_response,
            _render_error("Pagina nao encontrada."),
            status="404 Not Found",
        )
    except Exception as exc:  # pragma: no cover - defensive UI path
        details = html.escape("".join(traceback.format_exception_only(type(exc), exc)).strip())
        return _html_response(
            start_response,
            _render_error(f"Falha ao processar arquivo: {details}"),
            status="500 Internal Server Error",
        )


def _handle_process(environ, start_response):
    form = cgi.FieldStorage(fp=environ["wsgi.input"], environ=environ, keep_blank_values=True)
    uploaded_file = form["billing_file"] if "billing_file" in form else None
    file_format = form.getfirst("format", "aws-invoice")
    cloud = form.getfirst("cloud", "aws")
    company_name = form.getfirst("company_name", "").strip()
    project_name = form.getfirst("project_name", "").strip()
    llm_model = form.getfirst("llm_model", os.getenv("GEMINI_MODEL", "gemini-2.5-flash")).strip()

    if uploaded_file is None or not getattr(uploaded_file, "filename", ""):
        return _html_response(
            start_response,
            _render_error("Selecione um arquivo de billing (CSV ou PDF) antes de continuar."),
            status="400 Bad Request",
        )

    file_name = Path(uploaded_file.filename).name
    suffix = Path(file_name).suffix or ".csv"

    with tempfile.TemporaryDirectory() as temp_dir:
        temp_dir_path = Path(temp_dir)
        input_path = temp_dir_path / f"upload{suffix}"
        report_id = uuid.uuid4().hex
        output_name = f"{Path(file_name).stem}_report.xlsx"
        output_path = WEB_OUTPUT_DIR / f"{report_id}_{output_name}"
        presentation_name = f"{Path(file_name).stem}_report.pptx"
        presentation_path = WEB_OUTPUT_DIR / f"{report_id}_{presentation_name}"

        file_bytes = uploaded_file.file.read()
        input_path.write_bytes(file_bytes)

        result = process_billing_file(
            input_path=input_path,
            output_path=output_path,
            presentation_path=presentation_path,
            file_format=file_format,
            cloud=cloud,
            mapping_path=MAPPING_PATH,
            company_name=company_name,
            project_name=project_name,
            llm_model=llm_model,
        )

    REPORTS[report_id] = {
        "path": result.output_path,
        "download_name": output_name,
        "presentation_path": result.presentation_path,
        "presentation_name": presentation_name,
        "preview": _build_preview_context(result),
    }
    return _html_response(
        start_response,
        _render_result(
            report_id,
            REPORTS[report_id],
            company_name=company_name,
            project_name=project_name,
        ),
    )


def _handle_download(report_id: str, start_response):
    report = REPORTS.get(report_id)
    if report is None:
        return _html_response(
            start_response,
            _render_error("Arquivo nao encontrado ou expirado."),
            status="404 Not Found",
        )

    output_path = Path(report["path"])
    if not output_path.exists():
        return _html_response(
            start_response,
            _render_error("Arquivo de relatorio nao esta mais disponivel."),
            status="404 Not Found",
        )

    response_body = output_path.read_bytes()
    headers = [
        ("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        ("Content-Disposition", f'attachment; filename="{report["download_name"]}"'),
        ("Content-Length", str(len(response_body))),
    ]
    start_response("200 OK", headers)
    return [response_body]


def _handle_download_ppt(report_id: str, start_response):
    report = REPORTS.get(report_id)
    if report is None:
        return _html_response(
            start_response,
            _render_error("Arquivo nao encontrado ou expirado."),
            status="404 Not Found",
        )

    presentation_path = Path(report["presentation_path"])
    if not presentation_path.exists():
        return _html_response(
            start_response,
            _render_error("Arquivo de apresentacao nao esta mais disponivel."),
            status="404 Not Found",
        )

    response_body = presentation_path.read_bytes()
    headers = [
        (
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        ),
        ("Content-Disposition", f'attachment; filename="{report["presentation_name"]}"'),
        ("Content-Length", str(len(response_body))),
    ]
    start_response("200 OK", headers)
    return [response_body]


def _handle_font(font_name: str, start_response):
    font_path = resolve_font_asset(font_name)
    if font_path is None:
        return _text_response(start_response, "", status="404 Not Found")

    body = font_path.read_bytes()
    headers = [
        ("Content-Type", guess_content_type(font_path)),
        ("Content-Length", str(len(body))),
        ("Cache-Control", "public, max-age=3600"),
    ]
    start_response("200 OK", headers)
    return [body]


def _build_preview_context(result) -> dict[str, object]:
    total_cost = float(result.raw_df["cost"].sum()) if not result.raw_df.empty else 0.0
    currency = "USD"
    if not result.raw_df.empty and "currency" in result.raw_df.columns:
        currency = str(result.raw_df["currency"].mode().iloc[0])

    top_services = []
    for _, row in result.service_summary_df.head(5).iterrows():
        top_services.append(
            {
                "service": str(row["service_name_original"]),
                "cost": float(row["total_cost"]),
            }
        )

    unmapped_services = []
    unmapped_df = result.oci_mapping_df[
        result.oci_mapping_df["oci_service"] == "REVIEW_REQUIRED"
    ]
    for _, row in unmapped_df.head(10).iterrows():
        unmapped_services.append(str(row["service_name_original"]))

    return {
        "currency": currency,
        "total_cost": total_cost,
        "top_services": top_services,
        "unmapped_services": unmapped_services,
        "raw_rows": len(result.raw_df),
    }


def _render_home() -> str:
    return """
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Cloud Billing Agent</title>
  <style>
    :root {
      --bg: #0b0f14;
      --panel: rgba(17, 24, 39, 0.92);
      --panel-2: rgba(10, 15, 24, 0.88);
      --ink: #f3f4f6;
      --muted: #9ca3af;
      --accent: #ea4335;
      --accent-dark: #b91c1c;
      --line: rgba(255, 255, 255, 0.09);
      --glow: rgba(234, 67, 53, 0.22);
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: Verdana, "Segoe UI", Tahoma, sans-serif;
      background:
        radial-gradient(circle at top left, rgba(234, 67, 53, 0.16) 0, transparent 28%),
        radial-gradient(circle at top right, rgba(59, 130, 246, 0.12) 0, transparent 24%),
        linear-gradient(135deg, #05070b 0%, #0b0f14 45%, #101826 100%);
      color: var(--ink);
      min-height: 100vh;
    }
    .wrap {
      max-width: 980px;
      margin: 0 auto;
      padding: 48px 24px 64px;
    }
    .brand {
      display: inline-flex;
      align-items: center;
      gap: 16px;
      margin-bottom: 22px;
      padding: 10px 14px;
      border: 1px solid var(--line);
      border-radius: 999px;
      background: rgba(255,255,255,0.03);
      box-shadow: 0 0 0 1px rgba(255,255,255,0.02), 0 10px 30px rgba(0,0,0,0.35);
    }
    .logo-oracle {
      font-family: Verdana, "Segoe UI", Tahoma, sans-serif;
      font-weight: 700;
      letter-spacing: 0.28em;
      color: #f43f33;
      font-size: 14px;
      text-transform: uppercase;
    }
    .brand-note {
      color: var(--muted);
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: 0.12em;
    }
    .hero {
      display: grid;
      gap: 12px;
      margin-bottom: 28px;
    }
    .eyebrow {
      text-transform: uppercase;
      letter-spacing: 0.12em;
      color: var(--accent-dark);
      font-size: 12px;
      font-weight: 700;
    }
    h1 {
      margin: 0;
      font-size: clamp(34px, 5vw, 54px);
      line-height: 1;
      max-width: 12ch;
    }
    .sub {
      margin: 0;
      max-width: 60ch;
      color: var(--muted);
      font-size: 18px;
      line-height: 1.5;
    }
    .grid {
      display: grid;
      grid-template-columns: 1.1fr 0.9fr;
      gap: 20px;
    }
    .panel {
      background: linear-gradient(180deg, var(--panel) 0%, var(--panel-2) 100%);
      border: 1px solid var(--line);
      border-radius: 22px;
      padding: 24px;
      box-shadow: 0 18px 45px rgba(0, 0, 0, 0.34), 0 0 0 1px rgba(255,255,255,0.02);
    }
    label {
      display: block;
      margin-bottom: 8px;
      font-weight: 700;
    }
    input[type=file], select {
      width: 100%;
      margin-bottom: 18px;
      padding: 14px 16px;
      border-radius: 14px;
      border: 1px solid rgba(255,255,255,0.10);
      background: rgba(255,255,255,0.04);
      color: #f9fafb;
      font-size: 15px;
    }
    button {
      appearance: none;
      border: none;
      background: linear-gradient(135deg, var(--accent) 0%, #c62828 100%);
      color: white;
      font-size: 16px;
      font-weight: 700;
      padding: 14px 20px;
      border-radius: 999px;
      cursor: pointer;
      width: 100%;
      box-shadow: 0 16px 34px var(--glow);
    }
    ul {
      margin: 0;
      padding-left: 20px;
      color: var(--muted);
      line-height: 1.6;
    }
    .card-title {
      margin-top: 0;
      margin-bottom: 12px;
      font-size: 20px;
    }
    .note {
      margin-top: 18px;
      color: var(--muted);
      font-size: 14px;
      line-height: 1.5;
    }
    .preview-grid {
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 14px;
      margin-top: 18px;
    }
    .stat {
      padding: 16px;
      border: 1px solid var(--line);
      border-radius: 16px;
      background: rgba(255,255,255,0.04);
    }
    .stat strong {
      display: block;
      font-size: 24px;
      margin-top: 6px;
    }
    .result-actions {
      margin-top: 18px;
      display: flex;
      gap: 12px;
      flex-wrap: wrap;
    }
    .ghost {
      display: inline-block;
      text-decoration: none;
      padding: 12px 18px;
      border-radius: 999px;
      border: 1px solid rgba(255,255,255,0.10);
      color: var(--ink);
      font-weight: 700;
      background: rgba(255,255,255,0.04);
    }
    @media (max-width: 840px) {
      .grid { grid-template-columns: 1fr; }
      .wrap { padding: 28px 16px 40px; }
      .preview-grid { grid-template-columns: 1fr; }
    }
  </style>
</head>
<body>
  <main class="wrap">
    <div class="brand">
      <div class="logo-oracle">Oracle</div>
      <div class="brand-note">Cloud Billing Agent</div>
    </div>
    <section class="hero">
      <div class="eyebrow">Cloud Billing Agent</div>
      <h1>Upload de billing em csv e pdf</h1>
      <p class="sub">
        Envie o CSV ou PDF da fatura, escolha o CSP correto e gere Excel e PowerPoint
        com resumos, graficos, mapeamento OCI e validacoes de qualidade.
      </p>
    </section>

    <section class="grid">
      <div class="panel">
        <h2 class="card-title">Gerar relatorio</h2>
        <form action="/process" method="post" enctype="multipart/form-data">
          <label for="billing_file">Arquivo CSV</label>
          <input id="billing_file" type="file" name="billing_file" accept=".csv,.pdf" required>

          <label for="company_name">Nome da empresa</label>
          <input id="company_name" type="text" name="company_name" placeholder="Ex.: Tractian">

          <label for="project_name">Projeto / Assessment</label>
          <input id="project_name" type="text" name="project_name" placeholder="Ex.: OCI Conversion Assessment">

          <label for="format">Formato</label>
          <select id="format" name="format">
            <option value="aws-invoice" selected>AWS Invoice CSV</option>
            <option value="aws-billing-pdf">AWS Billing PDF</option>
            <option value="gcp-cost-table">GCP Cost table CSV</option>
            <option value="generic">CSV generico</option>
          </select>

          <label for="cloud">Cloud de origem</label>
          <select id="cloud" name="cloud">
            <option value="aws" selected>AWS</option>
            <option value="azure">Azure</option>
            <option value="gcp">GCP</option>
          </select>

          <label for="llm_model">Modelo LLM (Gemini / Google AI Studio)</label>
          <input id="llm_model" type="text" name="llm_model" value="gemini-2.5-flash" placeholder="Ex.: gemini-2.5-flash">

          <button type="submit">Processar relatorios</button>
        </form>
        <p class="note">
          Para AWS Invoice, a regra padrao remove linhas sem <code>LinkedAccountName</code>
          antes da analise. Para AWS Billing PDF, o parser extrai linhas de uso/custo do layout
          de fatura consolidada da AWS Billing and Cost Management. Para GCP Cost table, o parser
          ignora os metadados iniciais do export, normaliza valores em BRL/USD e remove linhas de
          totalizacao/impostos da analise por servico. A analise LLM usa <code>GEMINI_API_KEY</code>
          quando disponivel.
        </p>
      </div>

      <aside class="panel">
        <h2 class="card-title">O que voce recebe</h2>
        <ul>
          <li>Resumo por servico e por regiao.</li>
          <li>Graficos de custo total e participacao percentual.</li>
          <li>Mapeamento inicial de servicos para OCI.</li>
          <li>Aba de qualidade mostrando filtros e reconciliacao.</li>
          <li>Deck executivo em PowerPoint pronto para apresentar.</li>
        </ul>
      </aside>
    </section>
  </main>
</body>
</html>
"""


def _render_result(
    report_id: str,
    report: dict[str, object],
    company_name: str = "",
    project_name: str = "",
) -> str:
    preview = report["preview"]
    currency = html.escape(str(preview["currency"]))
    total_cost = f"{preview['total_cost']:,.2f}"
    raw_rows = int(preview["raw_rows"])
    company_badge = (
        f"<p><strong>Empresa na capa:</strong> {html.escape(company_name)}</p>"
        if company_name
        else ""
    )
    project_badge = (
        f"<p><strong>Projeto na capa:</strong> {html.escape(project_name)}</p>"
        if project_name
        else ""
    )

    top_services_markup = "".join(
        f"<li>{html.escape(item['service'])} <strong>{item['cost']:,.2f} {currency}</strong></li>"
        for item in preview["top_services"]
    )
    if not top_services_markup:
        top_services_markup = "<li>Nenhum servico encontrado.</li>"

    unmapped_markup = "".join(
        f"<li>{html.escape(service)}</li>" for service in preview["unmapped_services"]
    )
    if not unmapped_markup:
        unmapped_markup = "<li>Todos os servicos desta amostra possuem mapeamento inicial.</li>"

    return f"""
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Relatorio pronto</title>
  <style>
    :root {{
      --bg: #0b0f14;
      --panel: rgba(17, 24, 39, 0.92);
      --panel-2: rgba(10, 15, 24, 0.88);
      --ink: #f3f4f6;
      --muted: #9ca3af;
      --accent: #ea4335;
      --line: rgba(255, 255, 255, 0.09);
      --glow: rgba(234, 67, 53, 0.22);
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: Verdana, "Segoe UI", Tahoma, sans-serif;
      background:
        radial-gradient(circle at top left, rgba(234, 67, 53, 0.16) 0, transparent 28%),
        radial-gradient(circle at top right, rgba(59, 130, 246, 0.12) 0, transparent 24%),
        linear-gradient(135deg, #05070b 0%, #0b0f14 45%, #101826 100%);
      color: var(--ink);
      min-height: 100vh;
    }}
    .wrap {{ max-width: 1040px; margin: 0 auto; padding: 40px 24px 64px; }}
    .brand {{
      display: inline-flex;
      align-items: center;
      gap: 16px;
      margin-bottom: 22px;
      padding: 10px 14px;
      border: 1px solid var(--line);
      border-radius: 999px;
      background: rgba(255,255,255,0.03);
      box-shadow: 0 0 0 1px rgba(255,255,255,0.02), 0 10px 30px rgba(0,0,0,0.35);
    }}
    .logo-oracle {{
      font-family: Verdana, "Segoe UI", Tahoma, sans-serif;
      font-weight: 700;
      letter-spacing: 0.28em;
      color: #f43f33;
      font-size: 14px;
      text-transform: uppercase;
    }}
    .brand-note {{
      color: var(--muted);
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: 0.12em;
    }}
    .panel {{
      background: linear-gradient(180deg, var(--panel) 0%, var(--panel-2) 100%);
      border: 1px solid var(--line);
      border-radius: 22px;
      padding: 24px;
      box-shadow: 0 18px 45px rgba(0, 0, 0, 0.34), 0 0 0 1px rgba(255,255,255,0.02);
      margin-bottom: 20px;
    }}
    h1, h2 {{ margin-top: 0; }}
    .preview-grid {{
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 14px;
      margin-top: 18px;
    }}
    .stat {{
      padding: 16px;
      border: 1px solid var(--line);
      border-radius: 16px;
      background: rgba(255,255,255,0.04);
    }}
    .stat strong {{
      display: block;
      font-size: 24px;
      margin-top: 6px;
    }}
    .grid {{
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 20px;
    }}
    ul {{
      margin: 0;
      padding-left: 20px;
      line-height: 1.7;
    }}
    .actions {{
      display: flex;
      gap: 12px;
      flex-wrap: wrap;
      margin-top: 18px;
    }}
    .button {{
      appearance: none;
      text-decoration: none;
      border: none;
      background: linear-gradient(135deg, var(--accent) 0%, #c62828 100%);
      color: white;
      font-size: 16px;
      font-weight: 700;
      padding: 14px 20px;
      border-radius: 999px;
      box-shadow: 0 16px 34px var(--glow);
    }}
    .ghost {{
      text-decoration: none;
      padding: 13px 18px;
      border-radius: 999px;
      border: 1px solid rgba(255,255,255,0.10);
      color: var(--ink);
      font-weight: 700;
      background: rgba(255,255,255,0.04);
    }}
    @media (max-width: 840px) {{
      .grid, .preview-grid {{ grid-template-columns: 1fr; }}
      .wrap {{ padding: 24px 16px 40px; }}
    }}
  </style>
</head>
<body>
  <main class="wrap">
    <div class="brand">
      <div class="logo-oracle">Oracle</div>
      <div class="brand-note">Cloud Billing Agent</div>
    </div>
    <section class="panel">
      <h1>Relatorio pronto para download</h1>
      <p>O arquivo foi processado com sucesso. Abaixo esta uma previa rapida antes de baixar o Excel ou o PowerPoint.</p>
      {company_badge}
      {project_badge}

      <div class="preview-grid">
        <div class="stat">
          Custo total analisado
          <strong>{total_cost} {currency}</strong>
        </div>
        <div class="stat">
          Linhas analisadas
          <strong>{raw_rows}</strong>
        </div>
        <div class="stat">
          Servicos nao mapeados
          <strong>{len(preview["unmapped_services"])}</strong>
        </div>
      </div>

      <div class="actions">
        <a class="button" href="/download/{report_id}">Baixar Excel</a>
        <a class="ghost" href="/download-ppt/{report_id}">Baixar PowerPoint</a>
        <a class="ghost" href="/">Processar outro arquivo</a>
      </div>
    </section>

    <section class="grid">
      <div class="panel">
        <h2>Top servicos por custo</h2>
        <ul>{top_services_markup}</ul>
      </div>
      <div class="panel">
        <h2>Servicos sem mapeamento OCI</h2>
        <ul>{unmapped_markup}</ul>
      </div>
    </section>
  </main>
</body>
</html>
"""


def _render_error(message: str) -> str:
    safe_message = html.escape(message)
    return f"""
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <title>Erro</title>
  <style>
    body {{ font-family: "Segoe UI", Tahoma, sans-serif; background: #f7f3ee; color: #2d3748; padding: 32px; }}
    .box {{ max-width: 720px; margin: 0 auto; background: white; border: 1px solid #eadfcf; border-radius: 18px; padding: 24px; }}
    a {{ color: #b45309; }}
  </style>
</head>
<body>
  <div class="box">
    <h1>Falha no processamento</h1>
    <p>{safe_message}</p>
    <p><a href="/">Voltar</a></p>
  </div>
</body>
</html>
"""


def _html_response(start_response, content: str, status: str = "200 OK") -> Iterable[bytes]:
    body = content.encode("utf-8")
    headers = [
        ("Content-Type", "text/html; charset=utf-8"),
        ("Content-Length", str(len(body))),
    ]
    start_response(status, headers)
    return [body]


def _text_response(start_response, content: str, status: str = "200 OK") -> Iterable[bytes]:
    body = content.encode("utf-8")
    headers = [
        ("Content-Type", "text/plain; charset=utf-8"),
        ("Content-Length", str(len(body))),
    ]
    start_response(status, headers)
    return [body]


def main() -> None:
    load_local_config()
    url = f"http://{HOST}:{PORT}"
    print(f"Cloud Billing Agent Web UI disponivel em {url}")
    print("Pressione Ctrl+C para encerrar.")
    try:
        webbrowser.open(url)
    except Exception:
        pass

    with make_server(HOST, PORT, application) as httpd:
        httpd.serve_forever()


if __name__ == "__main__":
    main()
