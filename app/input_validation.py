from __future__ import annotations

from pathlib import Path


CSV_FORMATS = {"generic", "aws-invoice", "azure-cost-csv", "gcp-cost-table"}
PDF_FORMATS = {"aws-billing-pdf"}
DISALLOWED_MAGIC_PREFIXES = (
    b"MZ",  # Windows executables
    b"PK\x03\x04",  # Zip/Office archives
    b"\x7fELF",  # Linux executables
)
SNIFF_BYTES = 16384


def validate_billing_input_file(input_path: Path, file_format: str) -> None:
    if not input_path.exists() or not input_path.is_file():
        raise ValueError("Arquivo de entrada invalido ou inexistente.")

    if input_path.stat().st_size <= 0:
        raise ValueError("Arquivo de entrada vazio.")

    file_head = input_path.read_bytes()[:SNIFF_BYTES]
    if not file_head:
        raise ValueError("Arquivo de entrada vazio.")

    if file_format in PDF_FORMATS:
        _validate_pdf_content(input_path, file_head)
        return

    if file_format in CSV_FORMATS:
        _validate_csv_content(file_head)
        return

    raise ValueError(f"Formato nao suportado para validacao: {file_format}")


def _validate_pdf_content(input_path: Path, file_head: bytes) -> None:
    if not file_head.startswith(b"%PDF-"):
        raise ValueError("Conteudo invalido: arquivo nao possui assinatura PDF valida.")

    tail_size = min(2048, input_path.stat().st_size)
    with input_path.open("rb") as handle:
        handle.seek(-tail_size, 2)
        file_tail = handle.read()
    if b"%%EOF" not in file_tail:
        raise ValueError("Conteudo invalido: PDF sem marcador de encerramento esperado.")


def _validate_csv_content(file_head: bytes) -> None:
    for signature in DISALLOWED_MAGIC_PREFIXES:
        if file_head.startswith(signature):
            raise ValueError("Conteudo invalido: arquivo nao e um CSV textual.")

    if file_head.startswith(b"%PDF-"):
        raise ValueError("Conteudo invalido: arquivo PDF enviado para parser CSV.")

    if b"\x00" in file_head:
        raise ValueError("Conteudo invalido: bytes nulos detectados no arquivo.")

    sample_text = _decode_with_fallback(file_head)
    if not _looks_like_delimited_text(sample_text):
        raise ValueError(
            "Conteudo invalido: cabecalho/linhas nao correspondem a CSV delimitado."
        )


def _decode_with_fallback(file_head: bytes) -> str:
    for encoding in ("utf-8-sig", "latin-1"):
        try:
            return file_head.decode(encoding, errors="strict")
        except UnicodeDecodeError:
            continue
    raise ValueError("Conteudo invalido: nao foi possivel decodificar como texto CSV.")


def _looks_like_delimited_text(sample_text: str) -> bool:
    lines = [line.strip() for line in sample_text.splitlines() if line.strip()]
    if not lines:
        return False

    candidate_lines = lines[:8]
    delimiters = [",", ";", "\t"]
    for delimiter in delimiters:
        lines_with_delimiter = sum(1 for line in candidate_lines if delimiter in line)
        if lines_with_delimiter >= 1:
            return True
    return False
