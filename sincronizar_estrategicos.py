#!/usr/bin/env python3
"""
sincronizar_estrategicos.py
Atualiza o painel-estrategicos.html com os dados mais recentes da planilha
Analise de Tribunais - aba "Informacoes Institucionais".

Uso:
    python sincronizar_estrategicos.py

Dependencias:
    pip install -r requirements.txt
"""

import os, sys, re, json, csv, io
from pathlib import Path
from datetime import datetime

# -- Configuracao ---------------------------------------------------------------
FILE_ID     = "1o3hYd5euWdgNHms-qL6ejlA3gdiLJ2K0UHefvcaVgrE"
SHEET_NAME  = "Informações Institucionais"
SCOPES      = ["https://www.googleapis.com/auth/spreadsheets.readonly",
               "https://www.googleapis.com/auth/drive.readonly"]
CREDENTIALS          = "credentials.json"
TOKEN_FILE           = "token.json"
SERVICE_ACCOUNT_FILE = "service_account.json"
SERVICE_ACCOUNT_ENV  = "GOOGLE_SERVICE_ACCOUNT_JSON"
OUTPUT_HTML          = "painel-estrategicos.html"
TEMPLATE_BEFORE_FILE = "template_before.txt"
TEMPLATE_AFTER_FILE  = "template_after.txt"

# -- Auth -----------------------------------------------------------------------
def get_credentials():
    try:
        from google.oauth2 import service_account as gsa
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
        from google.auth.transport.requests import Request
    except ImportError:
        print("\nDependencias ausentes. Execute: pip install -r requirements.txt")
        sys.exit(1)

    sa_json = os.environ.get(SERVICE_ACCOUNT_ENV)
    if sa_json:
        return gsa.Credentials.from_service_account_info(json.loads(sa_json), scopes=SCOPES)
    if Path(SERVICE_ACCOUNT_FILE).exists():
        return gsa.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

    creds = None
    if Path(TOKEN_FILE).exists():
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not Path(CREDENTIALS).exists():
                print(f"\nNenhuma credencial encontrada.")
                sys.exit(1)
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "w") as f:
            f.write(creds.to_json())
    return creds

# -- Download -------------------------------------------------------------------
def download_sheet(creds, file_id, sheet_name):
    from googleapiclient.discovery import build
    service = build("sheets", "v4", credentials=creds)
    result = service.spreadsheets().values().get(
        spreadsheetId=file_id,
        range=sheet_name
    ).execute()
    return result.get("values", [])

# -- Parser ---------------------------------------------------------------------
def clean(v): return (v or "").strip()

def parse_valor_br(s):
    s = clean(s)
    if not s or "-" in s: return 0.0
    nums = re.sub(r"[^\d.,]", "", s)
    if not nums: return 0.0
    nums = nums.replace(".", "").replace(",", ".")
    try: return float(nums)
    except: return 0.0

def classTrib(s):
    if not s or s == "-": return "N/D"
    up = s.upper()
    for t in ["STJ","STF","TRF1","TRF5","TRF3","TRF2","TJPE","TJCE","TJAL","TJMA","TJBA","TJPI","TJSP","TJPB","TJRJ","TJRS","TSE"]:
        if t in up: return t
    return s.split(" ")[0][:6]

def classGrau(loc, trib):
    l, t = (loc or "").upper(), (trib or "").upper()
    if re.search(r"1[°º]\s*GRAU", l): return "1º Grau"
    if re.search(r"2[°º]\s*GRAU", l): return "2º Grau"
    if "SUPERIOR" in l or "TRIBUNAIS SUPERIORES" in l or t in ("STJ","STF","TSE"): return "Superior"
    if "TRF" in t: return "2º Grau"
    return "N/D"

def parse_rows(rows):
    """
    Mapeamento de colunas (planilha Analise de Tribunais):
      0  Processo              1  Nome das Partes
      2  Tribunal Responsavel  3  Localizacao do Processo
      4  Vara de Origem        5  Magistrado de Piso
      6  Turma                 7  Magistrado Responsavel
      8  Composicao            9  Processo Pautado?
     10  Tipo da Sessao       11  Data da Sessao
     12  Resultado            13  Data de conclusao
     14  Tipo da Acao         15  Materia
     16  Motivo               17  Objetivo
     18  Envolvimento         19  Responsavel Matriz
     20  Responsavel Filial   21  Setor Responsavel
     22  Filial Responsavel   23  Valores Envolvidos
     24  Valores Contingencia 25  Observacoes
    """
    if not rows or len(rows) < 2:
        return []

    def norm(s): return (s or "").strip().lower()
    headers = [norm(h) for h in rows[0]]

    def idx(*terms):
        for t in terms:
            nt = norm(t)
            i = next((i for i,h in enumerate(headers) if nt in h), -1)
            if i >= 0: return i
        return -1

    iProc    = idx("processo")
    iPartes  = idx("nome das partes", "partes")
    iTrib    = idx("tribunal responsavel", "tribunal")
    iLoc     = idx("localizacao", "localiza")
    iVara    = idx("vara de origem", "vara")
    iMagPiso = idx("magistrado de piso")
    iTurma   = idx("turma")
    iMagResp = idx("magistrado responsavel", "relator")
    iPautado = idx("processo pautado", "pautado")
    iTipoS   = idx("tipo da sessao")
    iDataS   = idx("data da sessao")
    iTipoA   = idx("tipo da acao", "tipo de acao")
    iMat     = idx("materia")
    iSetor   = idx("setor responsavel", "setor")
    iFilial  = idx("filial responsavel", "filial")
    iValor   = idx("valores envolvidos", "valor envolvido")

    def g(row, i): return clean(row[i]) if i >= 0 and i < len(row) else ""

    records = []
    for row in rows[1:]:
        proc   = g(row, iProc)
        partes = g(row, iPartes)
        if not proc and not partes:
            continue

        trib_raw = g(row, iTrib)
        loc_raw  = g(row, iLoc)
        trib = classTrib(trib_raw)
        grau = classGrau(loc_raw, trib)

        loc = loc_raw or ""
        if not loc or "TRIBUNAIS SUPERIORES" in loc.upper(): loc = trib
        elif re.search(r"1[°º]\s*GRAU", loc): loc = trib + " 1º Grau"
        elif re.search(r"2[°º]\s*GRAU", loc): loc = trib + " 2º Grau"

        data_sessao = g(row, iDataS)
        if data_sessao and re.match(r"^\d{5}$", data_sessao):
            from datetime import date
            d = date.fromordinal(date(1899,12,30).toordinal() + int(data_sessao))
            data_sessao = d.strftime("%d/%m/%Y")

        records.append({
            "proc":        proc,
            "partes":      partes,
            "trib_raw":    trib_raw,
            "loc_raw":     loc_raw,
            "loc":         loc,
            "trib":        trib,
            "grau":        grau,
            "vara":        g(row, iVara),
            "mag_piso":    g(row, iMagPiso),
            "turma":       g(row, iTurma),
            "mag_resp":    g(row, iMagResp),    # col 7
            "pautado":     g(row, iPautado) or "Não",
            "tipo_sessao": g(row, iTipoS),
            "data_sessao": data_sessao,          # col 11
            "tipo_acao":   g(row, iTipoA),       # col 14
            "materia":     g(row, iMat),         # col 15
            "setor":       g(row, iSetor),       # col 21
            "filial":      g(row, iFilial),      # col 22
            "valor":       g(row, iValor),
        })
    return records

# -- Geracao do HTML ------------------------------------------------------------
def build_html(records):
    """Le os templates de arquivos separados e injeta os dados."""
    before_path = Path(TEMPLATE_BEFORE_FILE)
    after_path  = Path(TEMPLATE_AFTER_FILE)

    if not before_path.exists() or not after_path.exists():
        print(f"\nERRO: Arquivos de template nao encontrados.")
        print(f"  Necessario: {TEMPLATE_BEFORE_FILE} e {TEMPLATE_AFTER_FILE}")
        sys.exit(1)

    tmpl_before = before_path.read_text(encoding="utf-8")
    tmpl_after  = after_path.read_text(encoding="utf-8")
    data_json   = json.dumps(records, ensure_ascii=False, separators=(",", ":"))
    return tmpl_before + data_json + tmpl_after

# -- Main -----------------------------------------------------------------------
def main():
    t0 = datetime.now()
    print("=" * 56)
    print("  Sincronizador - Painel Processos Estrategicos")
    print("=" * 56)

    print("\n[1/3] Autenticando...")
    creds = get_credentials()
    print("      OK - Autenticado")

    print(f"\n[2/3] Baixando aba '{SHEET_NAME}'...")
    rows = download_sheet(creds, FILE_ID, SHEET_NAME)
    print(f"      OK - {len(rows)} linhas recebidas (incluindo cabecalho)")

    records = parse_rows(rows)
    print(f"      OK - {len(records)} processos parseados")

    if not records:
        print("\nNenhum registro encontrado.")
        sys.exit(1)

    print("\n[3/3] Gerando HTML...")
    html = build_html(records)
    out = Path(OUTPUT_HTML)
    out.write_text(html, encoding="utf-8")

    elapsed = (datetime.now() - t0).total_seconds()
    size_kb = out.stat().st_size / 1024

    print(f"      OK - {out.resolve()}")
    print(f"      OK - {size_kb:.0f} KB  ({len(records)} processos  |  {elapsed:.1f}s)")

    tribs = len({r["trib"] for r in records if r["trib"] and r["trib"] != "N/D"})
    paut  = sum(1 for r in records if (r["pautado"] or "").lower() == "sim")
    val   = sum(parse_valor_br(r["valor"]) for r in records)

    print()
    print("  +- Resumo -------------------------------------------")
    print(f"  |  Processos : {len(records):,}")
    print(f"  |  Tribunais : {tribs}")
    print(f"  |  Pautados  : {paut}")
    print(f"  |  Valor     : R$ {val/1e6:.1f} M")
    print("  +-----------------------------------------------------")
    print()

if __name__ == "__main__":
    main()
