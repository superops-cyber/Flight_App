#!/usr/bin/env python3
"""
Publish local Markdown template pack to Google Drive as Google Docs,
and generate a master index CSV + Google Sheet with clickable links.

Usage examples:
  python3 publish_document_pack_to_drive.py --dry-run
  python3 publish_document_pack_to_drive.py --credentials credentials.json --root-folder-name "Asas - Document Pack"

Requirements:
  pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib
"""

from __future__ import annotations

import argparse
import csv
import io
import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional
import html

try:
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseUpload
except Exception:
    Credentials = None

SCOPES = ["https://www.googleapis.com/auth/drive"]


@dataclass
class TemplateDoc:
    title: str
    category: str
    priority: str
    owner: str
    source: str
    local_path: Path
    drive_doc_id: str = ""
    drive_doc_url: str = ""


def parse_template(md_path: Path) -> TemplateDoc:
    text = md_path.read_text(encoding="utf-8", errors="ignore")
    lines = text.splitlines()

    title = ""
    category = ""
    priority = ""
    owner = ""
    source = ""

    if lines and lines[0].startswith("# "):
        title = lines[0][2:].strip()

    for line in lines:
        if line.startswith("**Categoria:**"):
            category = line.replace("**Categoria:**", "").strip()
        elif line.startswith("**Prioridade:**"):
            priority = line.replace("**Prioridade:**", "").strip()
        elif line.startswith("**Responsável sugerido:**"):
            owner = line.replace("**Responsável sugerido:**", "").strip()
        elif line.startswith("**Fonte (manuais):**"):
            source = line.replace("**Fonte (manuais):**", "").strip()

    if not title:
        title = md_path.stem

    return TemplateDoc(
        title=title,
        category=category,
        priority=priority,
        owner=owner,
        source=source,
        local_path=md_path,
    )


def slugify(text: str) -> str:
    text = text.strip().lower()
    text = re.sub(r"[^a-z0-9\s_-]", "", text)
    text = re.sub(r"\s+", "-", text)
    return text


def extract_folder_id(value: str) -> str:
    value = value.strip()
    if "/folders/" in value:
        return value.split("/folders/")[1].split("?")[0].split("/")[0]
    if "id=" in value:
        return value.split("id=")[1].split("&")[0]
    return value


def normalize_text(text: str) -> str:
    normalized = unicodedata.normalize("NFKD", text)
    ascii_text = "".join(ch for ch in normalized if not unicodedata.combining(ch))
    return ascii_text.lower().strip()


def extract_sequence_from_filename(path: Path) -> int:
    match = re.match(r"^(\d{1,3})_", path.name)
    if not match:
        return 0
    return int(match.group(1))


def derive_governance_fields(doc: TemplateDoc) -> Dict[str, str]:
    title_n = normalize_text(doc.title)
    category_n = normalize_text(doc.category)

    sequence = extract_sequence_from_filename(doc.local_path)
    code = f"FRM-EXT-{sequence:03d}" if sequence else f"FRM-EXT-{slugify(doc.title)[:12].upper()}"

    usage_by_category = {
        "governanca": "Controle de governança e conformidade documental",
        "regulatorio": "Comprovação regulatória e rastreabilidade ANAC",
        "operacoes": "Suporte à liberação, execução e registro de missão",
        "aeronave": "Controle técnico-operacional da aeronave",
        "manutencao": "Rastreabilidade de manutenção e aeronavegabilidade",
        "operacoes especiais": "Gestão de cenários operacionais especiais",
        "programas": "Gestão de programas e declarações aplicáveis",
    }
    usage = usage_by_category.get(category_n, "Controle documental operacional")

    pilot_direct_keywords = [
        "lista de passageiros",
        "cartao de instrucoes ao passageiro",
        "peso e balanceamento",
        "manifesto de carga",
        "biblioteca tecnica afm/poh e suplementos",
    ]
    if any(key in title_n for key in pilot_direct_keywords):
        access = "PILOTO + DISPATCH + BASE"
    elif "liberacao/despacho" in title_n or "decisoes operacionais" in title_n:
        access = "DISPATCH + PILOTO (consulta)"
    elif category_n in {"governanca", "regulatorio", "programas"}:
        access = "GESTÃO + QUALIDADE"
    elif category_n == "manutencao":
        access = "MANUTENÇÃO + GESTÃO"
    elif category_n == "aeronave":
        access = "DISPATCH + MANUTENÇÃO"
    elif category_n == "operacoes especiais":
        access = "DISPATCH + PILOTO (quando aplicável)"
    else:
        access = "DISPATCH + BASE"

    if "PILOTO" in access:
        active_storage = "Drive/OPS/Formularios_Vigentes + copia a bordo (quando aplicável)"
    elif category_n == "manutencao":
        active_storage = "Drive/MNT/Controle_Tecnico/Formularios_Vigentes"
    elif category_n == "aeronave":
        active_storage = "Drive/OPS/Aeronaves/Formularios_Vigentes"
    else:
        active_storage = "Drive/SGQ/Documentos_Controlados/Formularios_Vigentes"

    if category_n == "manutencao":
        archive_storage = "Drive/MNT/Arquivo_Morto + backup NAS_MNT + pasta física OM145"
    elif "PILOTO" in access:
        archive_storage = "Drive/OPS/Arquivo_Morto + Base (via física, se aplicável)"
    else:
        archive_storage = "Drive/SGQ/Arquivo_Morto + backup NAS_SGQ"

    retention = "36 meses"
    if any(key in title_n for key in ["jornada", "fadiga", "treinamento", "manifesto", "peso e balanceamento", "lista de passageiros", "liberacao/despacho"]):
        retention = "60 meses"
    if "contrato" in title_n:
        retention = "60 meses após término contratual"

    return {
        "codigo_interno": code,
        "uso_do_formulario": usage,
        "acesso_requerido": access,
        "local_vigente": active_storage,
        "local_arquivo_morto": archive_storage,
        "retencao_minima": retention,
    }


def load_templates(pack_dir: Path) -> List[TemplateDoc]:
    files = sorted(
        [
            path
            for path in pack_dir.glob("*.md")
            if path.name.lower() != "index.md"
        ]
    )
    return [parse_template(path) for path in files]


def get_drive_service(credentials_path: Path, token_path: Path):
    creds = None
    if token_path.exists():
        creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)

    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())

    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file(str(credentials_path), SCOPES)
        creds = flow.run_local_server(port=0)
        token_path.write_text(creds.to_json(), encoding="utf-8")

    return build("drive", "v3", credentials=creds)


def create_folder(drive, name: str, parent_id: Optional[str] = None) -> str:
    body = {
        "name": name,
        "mimeType": "application/vnd.google-apps.folder",
    }
    if parent_id:
        body["parents"] = [parent_id]

    folder = drive.files().create(
        body=body,
        fields="id,name",
        supportsAllDrives=True,
    ).execute()
    return folder["id"]


def upload_markdown_as_google_doc(drive, doc: TemplateDoc, parent_id: str) -> str:
    content = doc.local_path.read_text(encoding="utf-8", errors="ignore")
    html_content = markdown_to_html(content)
    bio = io.BytesIO(html_content.encode("utf-8"))
    media = MediaIoBaseUpload(bio, mimetype="text/html", resumable=False)

    body = {
        "name": doc.title,
        "mimeType": "application/vnd.google-apps.document",
        "parents": [parent_id],
    }
    result = drive.files().create(
        body=body,
        media_body=media,
        fields="id,name",
        supportsAllDrives=True,
    ).execute()
    return result["id"]


def format_inline(text: str) -> str:
    text = html.escape(text)
    text = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", text)
    return text


def table_to_html(lines: List[str]) -> str:
    rows = []
    for line in lines:
        stripped = line.strip()
        if not stripped.startswith("|"):
            continue
        cells = [format_inline(cell.strip()) for cell in stripped.strip("|").split("|")]
        rows.append(cells)

    if len(rows) < 2:
        return ""

    header = rows[0]
    body_rows = rows[2:] if len(rows) > 2 else []

    out = ['<table border="1" cellspacing="0" cellpadding="4">', '<thead><tr>']
    out.extend(f'<th>{cell}</th>' for cell in header)
    out.append('</tr></thead><tbody>')
    for row in body_rows:
        out.append('<tr>')
        out.extend(f'<td>{cell}</td>' for cell in row)
        out.append('</tr>')
    out.append('</tbody></table>')
    return ''.join(out)


def markdown_to_html(markdown: str) -> str:
    lines = markdown.splitlines()
    out = [
        '<html><head><meta charset="utf-8"></head><body>',
        '<style>body{font-family:Arial,sans-serif;font-size:11pt;} table{border-collapse:collapse;width:100%;margin:8px 0 14px;} th{background:#f0f0f0;} h1,h2,h3{color:#222;} p{margin:6px 0;} ul{margin:6px 0 10px 18px;}</style>',
    ]

    i = 0
    while i < len(lines):
        stripped = lines[i].strip()

        if not stripped:
            i += 1
            continue

        if stripped.startswith('|'):
            block = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                block.append(lines[i])
                i += 1
            out.append(table_to_html(block))
            continue

        if stripped.startswith('# '):
            out.append(f'<h1>{format_inline(stripped[2:].strip())}</h1>')
            i += 1
            continue
        if stripped.startswith('## '):
            out.append(f'<h2>{format_inline(stripped[3:].strip())}</h2>')
            i += 1
            continue
        if stripped.startswith('### '):
            out.append(f'<h3>{format_inline(stripped[4:].strip())}</h3>')
            i += 1
            continue

        if stripped.startswith('- '):
            items = []
            while i < len(lines) and lines[i].strip().startswith('- '):
                items.append(lines[i].strip()[2:])
                i += 1
            out.append('<ul>' + ''.join(f'<li>{format_inline(item)}</li>' for item in items) + '</ul>')
            continue

        paragraph_lines = [stripped]
        i += 1
        while i < len(lines):
            candidate = lines[i].strip()
            if not candidate or candidate.startswith('#') or candidate.startswith('- ') or candidate.startswith('|'):
                break
            paragraph_lines.append(candidate)
            i += 1
        out.append(f'<p>{format_inline(" ".join(paragraph_lines))}</p>')

    out.append('</body></html>')
    return '\n'.join(out)


def upload_csv_as_sheet(drive, csv_path: Path, parent_id: str, sheet_name: str) -> str:
    content = csv_path.read_bytes()
    bio = io.BytesIO(content)
    media = MediaIoBaseUpload(bio, mimetype="text/csv", resumable=False)

    body = {
        "name": sheet_name,
        "mimeType": "application/vnd.google-apps.spreadsheet",
        "parents": [parent_id],
    }
    result = drive.files().create(
        body=body,
        media_body=media,
        fields="id,name",
        supportsAllDrives=True,
    ).execute()
    return result["id"]


def write_csv(output_csv: Path, docs: List[TemplateDoc]) -> None:
    output_csv.parent.mkdir(parents=True, exist_ok=True)
    with output_csv.open("w", encoding="utf-8", newline="") as file:
        writer = csv.writer(file)
        writer.writerow(
            [
                "documento",
                "categoria",
                "prioridade",
                "responsavel",
                "fonte_manuais",
                "arquivo_local",
                "codigo_interno",
                "uso_do_formulario",
                "acesso_requerido",
                "local_vigente",
                "local_arquivo_morto",
                "retencao_minima",
                "drive_doc_id",
                "drive_doc_url",
                "hyperlink_formula",
                "status",
                "revisao",
                "vigencia",
                "ultima_atualizacao",
                "observacoes",
            ]
        )

        for doc in docs:
            governance = derive_governance_fields(doc)
            formula = f'=HYPERLINK("{doc.drive_doc_url}","Abrir")' if doc.drive_doc_url else ""
            writer.writerow(
                [
                    doc.title,
                    doc.category,
                    doc.priority,
                    doc.owner,
                    doc.source,
                    str(doc.local_path),
                    governance["codigo_interno"],
                    governance["uso_do_formulario"],
                    governance["acesso_requerido"],
                    governance["local_vigente"],
                    governance["local_arquivo_morto"],
                    governance["retencao_minima"],
                    doc.drive_doc_id,
                    doc.drive_doc_url,
                    formula,
                    "Não iniciado",
                    "",
                    "",
                    "",
                ]
            )


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--pack-dir", default="135/document_pack", help="Directory with template .md files")
    parser.add_argument("--output-csv", default="135/document_pack_master_index.csv", help="Path of generated master CSV")
    parser.add_argument("--credentials", default="credentials.json", help="Google OAuth client credentials JSON")
    parser.add_argument("--token", default="token.json", help="OAuth token cache file")
    parser.add_argument("--root-folder-name", default="Asas - Document Pack", help="Root folder name to create in Drive")
    parser.add_argument("--parent-folder-id", default="", help="Optional parent Drive folder ID where the root folder will be created")
    parser.add_argument("--parent-folder-url", default="", help="Optional parent Drive folder URL where the root folder will be created")
    parser.add_argument("--use-parent-as-root", action="store_true", help="Use the parent folder itself as root (do not create a new root folder)")
    parser.add_argument("--sheet-name", default="Asas - Master Index", help="Google Sheet name for converted CSV")
    parser.add_argument("--dry-run", action="store_true", help="Only generate local CSV without touching Drive")
    args = parser.parse_args()

    pack_dir = Path(args.pack_dir)
    output_csv = Path(args.output_csv)

    if not pack_dir.exists():
        raise SystemExit(f"Pack directory not found: {pack_dir}")

    docs = load_templates(pack_dir)
    if not docs:
        raise SystemExit("No template files found in pack directory.")

    if args.dry_run:
        write_csv(output_csv, docs)
        print(f"Dry run complete. CSV generated at: {output_csv}")
        print(f"Templates discovered: {len(docs)}")
        return

    if Credentials is None:
        raise SystemExit(
            "Google API libraries not installed. Run: pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib"
        )

    credentials_path = Path(args.credentials)
    token_path = Path(args.token)

    if not credentials_path.exists():
        raise SystemExit(
            f"Credentials file not found: {credentials_path}. Download OAuth client JSON from Google Cloud Console."
        )

    drive = get_drive_service(credentials_path, token_path)

    parent_id = ""
    if args.parent_folder_url:
        parent_id = extract_folder_id(args.parent_folder_url)
    elif args.parent_folder_id:
        parent_id = extract_folder_id(args.parent_folder_id)

    if args.use_parent_as_root and not parent_id:
        raise SystemExit("--use-parent-as-root requires --parent-folder-id or --parent-folder-url")

    if args.use_parent_as_root:
        root_id = parent_id
    else:
        root_id = create_folder(drive, args.root_folder_name, parent_id or None)

    docs_folder_id = create_folder(drive, "Google Docs", root_id)

    for doc in docs:
        doc_id = upload_markdown_as_google_doc(drive, doc, docs_folder_id)
        doc.drive_doc_id = doc_id
        doc.drive_doc_url = f"https://docs.google.com/document/d/{doc_id}/edit"

    write_csv(output_csv, docs)

    sheet_id = upload_csv_as_sheet(drive, output_csv, root_id, args.sheet_name)
    sheet_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit"

    print(f"Drive root folder: https://drive.google.com/drive/folders/{root_id}")
    print(f"Master sheet: {sheet_url}")
    print(f"Local CSV: {output_csv}")
    print(f"Docs created: {len(docs)}")


if __name__ == "__main__":
    main()
