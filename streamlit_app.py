# ────────────────────────────────────────────────────────────────────────────────
# app.py – Selecionar amostras no Google Sheets e gerar Excel
# Execute com:  streamlit run app.py
# ────────────────────────────────────────────────────────────────────────────────
from __future__ import annotations

import io
import os
import json
from datetime import datetime
from typing import List, Dict

import pandas as pd
import streamlit as st
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow

# ╭─────────────────────────── CONFIGURAÇÕES ─────────────────────────────╮
SCOPES         = ["https://www.googleapis.com/auth/spreadsheets"]
SPREADSHEET_ID = "1VLDQUCO3Aw4ClAvhjkUsnBxG44BTjz-MjHK04OqPxYM"
SHEET_NAME     = "Geral"

STATUS_COL = "AF"                     # coluna que recebe "Analisando Amostra"
DATE_COL   = "AG"                     # coluna que recebe a data de hoje
STATUS_VAL = "Analisando Amostra"
DATE_FMT   = "%d/%m/%Y"
# ╰────────────────────────────────────────────────────────────────────────╯

# ───────────────────────── Google Sheets helpers ────────────────────────
@st.cache_resource
def _authorize_google() -> Credentials:
    token_path = "token.json"
    creds = None
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            try:
                client_config = json.loads(st.secrets["GOOGLE_CLIENT_SECRET"])
            except Exception:
                st.error("❌ Credenciais do Google não encontradas.")
                st.stop()
            flow = InstalledAppFlow.from_client_config(client_config, SCOPES)
            creds = flow.run_console()
        try:
            with open(token_path, "w", encoding="utf-8") as fp:
                fp.write(creds.to_json())
        except Exception:
            pass
    return creds


@st.cache_resource
def _get_service():
    return build("sheets", "v4", credentials=_authorize_google(), cache_discovery=False)


def _col_to_index(col: str) -> int:
    """Converte 'A' → 0, 'B' → 1, …"""
    idx = 0
    for c in col:
        idx = idx * 26 + (ord(c.upper()) - 64)
    return idx - 1


def fetch_sheet() -> List[List[str]]:
    """Lê a aba inteira, forçando tudo como texto."""
    res = (
        _get_service()
        .spreadsheets()
        .values()
        .get(
            spreadsheetId=SPREADSHEET_ID,
            range=SHEET_NAME,
            valueRenderOption="FORMATTED_VALUE",  # <- devolve sempre string
        )
        .execute()
    )
    return res.get("values", [])


def update_status(rows_idx: List[int], today: str) -> None:
    """Atualiza colunas de STATUS e DATA nas linhas indicadas (1-based)."""
    svc = _get_service()
    status_range = f"{SHEET_NAME}!{STATUS_COL}{rows_idx[0]}:{STATUS_COL}{rows_idx[-1]}"
    date_range   = f"{SHEET_NAME}!{DATE_COL}{rows_idx[0]}:{DATE_COL}{rows_idx[-1]}"
    body_status = {"values": [[STATUS_VAL]] * len(rows_idx)}
    body_date   = {"values": [[today]]      * len(rows_idx)}

    try:
        svc.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=status_range,
            valueInputOption="RAW",
            body=body_status,
        ).execute()
        svc.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=date_range,
            valueInputOption="RAW",
            body=body_date,
        ).execute()
    except HttpError as exc:
        st.error("❌ Falha ao atualizar status no Google Sheets.")
        st.stop()

# ───────────────────────────── UI Streamlit ─────────────────────────────
st.set_page_config(page_title="Selecionar Amostras", page_icon="🛢️", layout="centered")
st.title("Selecionar Amostras 🛢️")

if "samples" not in st.session_state:
    st.session_state["samples"]: List[str] = []
if "current_input" not in st.session_state:
    st.session_state["current_input"] = ""

def _add_sample():
    code = st.session_state["current_input"].strip()
    if code and code not in st.session_state["samples"]:
        st.session_state["samples"].append(code)
    st.session_state["current_input"] = ""   # limpa campo após Enter

st.text_input(
    "📷 Escaneie o código de barras e pressione Enter",
    key="current_input",
    on_change=_add_sample,
)

st.write("### Amostras pré-selecionadas")
st.write(", ".join(st.session_state["samples"]) or "Nenhuma.")

col1, col2 = st.columns(2)
with col1:
    if st.button("🗑️ Limpar lista"):
        st.session_state["samples"].clear()
with col2:
    gerar = st.button("📥 Gerar planilha")

# ───────────────────── Geração da planilha Excel ────────────────────────
if gerar and st.session_state["samples"]:
    with st.spinner("Consultando Google Sheets…"):
        sheet_rows = fetch_sheet()
        if not sheet_rows:
            st.error("Aba vazia ou não encontrada.")
            st.stop()

        header, *data = sheet_rows
        g_idx = _col_to_index("G")
        selected_rows, lines_idx = [], []
        samples_set = {s.strip() for s in st.session_state["samples"]}

        for i, row in enumerate(data, start=2):              # 1-based (linha 1 = header)
            sample_text = str(row[g_idx]).strip() if g_idx < len(row) else ""
            if sample_text in samples_set:
                selected_rows.append(row)
                lines_idx.append(i)

        if not selected_rows:
            st.warning("Nenhuma amostra da lista foi encontrada na planilha.")
            st.stop()

    today = datetime.now().strftime(DATE_FMT)
    with st.spinner("Atualizando status no Sheets…"):
        update_status(lines_idx, today)

    with st.spinner("Gerando arquivo Excel…"):
        # Preenche linhas curtas até o tamanho do cabeçalho
        norm_rows = [
            r + [""] * (len(header) - len(r))
            for r in selected_rows
        ]
        df = pd.DataFrame(norm_rows, columns=header)
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Amostras")
        buf.seek(0)

    st.success(f"✔️ {len(df)} amostra(s) exportada(s).")
    st.download_button(
        "⬇️ Baixar Excel",
        data=buf,
        file_name=f"amostras_{today}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

elif gerar:
    st.error("📋 A lista de amostras está vazia.")
