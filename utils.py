# ────────────────────────────────────────────────────────────────────────────────
# utils.py – funções auxiliares (Google Sheets, PDF e UI Streamlit)
# ────────────────────────────────────────────────────────────────────────────────
"""Utilitários para:
1. Persistir respostas em Google Sheets
2. Gerar PDF A4 (uma página) com QR‑Code da amostra
3. Construir o formulário em Streamlit com duas caixas (Sim/Não) para questões
   binárias.

Requisitos (pip install):
    streamlit google-auth google-auth-oauthlib google-api-python-client
    fpdf2 qrcode[pil] pillow
"""
from __future__ import annotations

import io
import os
from datetime import datetime
from typing import Dict, List, Tuple, Any

import qrcode
from fpdf import FPDF
from fpdf.errors import FPDFException
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.auth.transport.requests import Request

from math import ceil
from barcode import Code128
from barcode.writer import ImageWriter

import streamlit as st
import json

# ░░░ Config Google Sheets ░░░───────────────────────────────────────────────────
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SPREADSHEET_ID = "1VLDQUCO3Aw4ClAvhjkUsnBxG44BTjz-MjHK04OqPxYM"
SHEET_NAME = "Geral"


# ░░░ Autenticação ░░░───────────────────────────────────────────────────────────

@st.cache_resource
def _authorize_google_sheets() -> Credentials:
    from google_auth_oauthlib.flow import InstalledAppFlow

    token_path = "token.json"
    creds = None

    # Em ambiente local, tenta reutilizar token salvo
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # Em produção (ex: Streamlit Cloud), usa client_config do secrets
            try:
                client_config = json.loads(st.secrets["GOOGLE_CLIENT_SECRET"])
            except Exception:
                st.error("❌ Não foi possível carregar as credenciais do Google.")
                st.stop()
            flow = InstalledAppFlow.from_client_config(client_config, SCOPES)
            creds = flow.run_console()

        # Salva token localmente apenas se possível
        try:
            with open(token_path, "w", encoding="utf-8") as fp:
                fp.write(creds.to_json())
        except Exception:
            pass  # Ignora falha de escrita no Streamlit Cloud

    return creds


@st.cache_resource
def _get_sheets_service():
    return build("sheets", "v4", credentials=_authorize_google_sheets(), cache_discovery=False)



# ░░░ Estrutura do formulário ░░░────────────────────────────────────────────────
# Cada tupla interna contém: (rótulo, valor padrão)
# Perguntas binárias usam bool para o valor padrão → serão renderizadas
# como duas caixas de seleção "Sim" e "Não".
FORM_SECTIONS: List[Tuple[str, List[Tuple[str, Any]]]] = [
    (
        "Geral",
        [
            ("Estado de Origem", "AM"),
            ("Cliente", "Pie - Oliveira Energia"),
            ("Data da coleta", datetime.today().strftime("%d/%m/%Y")),
            ("Local de operação", ""),
            ("UGD", ""),
            ("Responsável Pela Coleta", ""),
            ("n.º da Amostra", ""),  # obrigatório
        ],
    ),
    (
        "Equipamento",
        [
            ("n.º de série", ""),
            ("Frota", ""),
            ("Horímetro do Óleo", ""),
            ("Houve troca de óleo após coleta?", False),
            ("Troca de Filtro após coleta", False),
            ("Houve mudança do local de operação?", False),
            ("Fabricante", "Scania"),
            ("Modelo", "DC13"),
            ("Horímetro do Motor", ""),
        ],
    ),
    (
        "Óleo",
        [
            ("Houve complemento de óleo", False),
            ("Se sim, quantos litros", ""),
            ("Amostra coletada", "Motor"),
            ("Fabricante", "Mobil"),
            ("Grau de viscosidade", "15W40"),
            ("Nome", "Mobil Delvac"),
            ("Apresentou limalha no filtro ou na tela?", False),
            ("Apresentou limalhas no bujão magnético?", False),
            ("Equipamento apresentou ruído anormal?", False),
            ("Existem vazamentos no sistema", False),
            ("A temperatura de operação está normal?", False),
            ("O desempenho do sistema está normal?", False),
            ("Detalhes das anormalidades (caso Haja)", ""),
        ],
    ),
    (
        "Contato",
        [
            ("Pessoa de contato", "Francisco Sampaio"),
            ("Telefone", "(92) 99437-6579"),
        ],
    ),
]

SHEET_COLUMNS: List[str] = [label for _, block in FORM_SECTIONS for (label, _) in block]


# ░░░ UI Streamlit ░░░───────────────────────────────────────────────────────────
# Funções para construir o formulário e garantir duas caixas para perguntas Sim/Não


try:
    import streamlit as st
except ModuleNotFoundError:
    st = None  # Permite importar utils.py em scripts que não usem Streamlit


def _two_checkboxes(label: str, default: bool | None = None) -> bool:
    """Renderiza a pergunta e, abaixo, duas caixas mutuamente exclusivas (Sim/Não).

    Retorna **True** para Sim, **False** para Não.
    """
    if st is None:
        raise RuntimeError("Streamlit não instalado – UI indisponível.")

    # Apresenta o texto da pergunta
    st.markdown(f"**{label}**")

    # Chaves únicas para cada pergunta
    key_yes = f"{label}_yes"
    key_no = f"{label}_no"

    # Estado inicial coerente
    if key_yes not in st.session_state and key_no not in st.session_state:
        if default is True:
            st.session_state[key_yes] = True
            st.session_state[key_no] = False
        elif default is False:
            st.session_state[key_yes] = False
            st.session_state[key_no] = True
        else:
            st.session_state[key_yes] = False
            st.session_state[key_no] = False

    col_yes, col_no = st.columns(2)

    def _sync_yes():
        if st.session_state[key_yes]:
            st.session_state[key_no] = False

    def _sync_no():
        if st.session_state[key_no]:
            st.session_state[key_yes] = False

    with col_yes:
        st.checkbox("Sim", key=key_yes, on_change=_sync_yes)
    with col_no:
        st.checkbox("Não", key=key_no, on_change=_sync_no)

    # Retorno lógico final
    return bool(st.session_state[key_yes])


def build_form_and_get_responses() -> Dict[str, Any]:
    """Constrói UI e coleta respostas em um dicionário compatível com save_to_sheets."""
    if st is None:
        raise RuntimeError("Streamlit não instalado – UI indisponível.")

    st.header("Formulário de Coleta de Óleo")
    responses: Dict[str, Any] = {}

    for section, questions in FORM_SECTIONS:
        st.subheader(section)
        for label, default in questions:
            # Questões booleanas → duas checkboxes Sim/Não
            if isinstance(default, bool):
                responses[label] = _two_checkboxes(label, default=default)
            else:
                # Campo de texto ou número (uso simples de text_input)
                responses[label] = st.text_input(label, value=str(default))

    return responses


# ░░░ Google Sheets ░░░──────────────────────────────────────────────────────────

def save_to_sheets(responses: Dict[str, Any]) -> None:
    row = []
    for col in SHEET_COLUMNS:
        val = responses.get(col, "")
        row.append("Sim" if val is True else "Não" if val is False else str(val))

    body = {"values": [row]}
    try:
        service = _get_sheets_service()
        service.spreadsheets().values().append(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{SHEET_NAME}!A1",
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body=body,
        ).execute()
    except HttpError as exc:
        st.error("❌ Erro ao gravar no Google Sheets.")
        raise RuntimeError(f"Erro ao gravar → {exc}") from exc



# ░░░ Sanitização de texto ░░░───────────────────────────────────────────────────
_REPL = {
    "\u2013": "-",  # en dash
    "\u2014": "-",  # em dash
    "\u2011": "-",  # non‑breaking hyphen
    "\u00A0": " ",  # non‑breaking space
    "\n": " ",
    "\r": " ",
}


def _safe(txt: object) -> str:
    if txt is None:
        return ""
    if not isinstance(txt, str):
        txt = str(txt)
    for bad, good in _REPL.items():
        txt = txt.replace(bad, good)
    return txt.encode("latin-1", "replace").decode("latin-1")


# ░░░ PDF + QR Code (compacto em 1 página) ░░░───────────────────────────────────

# utils.py  ────────────────────────────────────────────────────────────────────
def generate_pdf(responses: Dict[str, Any]) -> bytes:
    """Gera um PDF A4 (retrato) em 1 página, usando tabelas com bordas
    e duas colunas de perguntas por linha.
    """
    sample_no = str(responses.get("n.º da Amostra", "SEM_NUMERO")).strip() or "SEM_NUMERO"

    # ░░░ Cria QR-Code em memória ░░░
    qr_img = qrcode.make(sample_no)
    buf = io.BytesIO()
    qr_img.save(buf, format="PNG")
    buf.seek(0)

    # ░░░ Configura página A4 ░░░
    pdf = FPDF(unit="mm", format="A4")
    pdf.set_auto_page_break(auto=False)
    pdf.set_left_margin(10)
    pdf.set_top_margin(10)
    pdf.set_right_margin(10)

    pdf.add_page()

    # ── Cabeçalho com QR à esquerda, título central e código de barras à direita ──

    # Gera QR Code
    qr_img = qrcode.make(sample_no)
    buf_qr = io.BytesIO()
    qr_img.save(buf_qr, format="PNG")
    buf_qr.seek(0)

    # Gera Código de Barras
    buf_bar = io.BytesIO()
    barcode = Code128(sample_no, writer=ImageWriter())
    barcode.write(buf_bar, options={
        "module_width": 0.3,
        "module_height": 15,
        "font_size": 8,
    })
    buf_bar.seek(0)

    # Tamanhos
    qr_w = 25
    bar_w = 30
    header_h = 20  # altura da faixa do cabeçalho
    y_start = pdf.get_y()

    # Posições
    x_qr  = pdf.l_margin
    x_bar = pdf.w - pdf.r_margin - bar_w

    # QR no canto esquerdo
    pdf.image(buf_qr, x=x_qr, y=y_start, w=qr_w)

    # Código de barras no canto direito
    pdf.image(buf_bar, x=x_bar, y=y_start + 5, w=bar_w)

    # Texto centralizado no topo
    pdf.set_font("Helvetica", size=16)
    pdf.set_y(y_start + 8)
    pdf.set_x(0)
    pdf.cell(w=0, h=10, txt="Oliveira Energia - Amostra de óleo", align="C", ln=True)

    # Espaço após o cabeçalho
    pdf.ln(8)


    # ── Corpo em tabelas ─────────────────────────────────────────────────────
    inner_width = pdf.w - pdf.l_margin - pdf.r_margin
    LABEL_RATIO = 0.655                 # <<< AQUI: rótulo maior
    group_width = inner_width / 2
    label_w  = group_width * LABEL_RATIO
    value_w  = group_width - label_w
    row_h    = 4.5

    pdf.set_font("Helvetica", size=7)

    for section, qs in FORM_SECTIONS:
        # ░░ Cabeçalho de seção: célula cheia ░░
        pdf.set_font("Helvetica", style="B", size=11)
        pdf.set_fill_color(240)            # cinza-claro
        pdf.cell(0, 7, _safe(section), ln=True, border=1, fill=True)
        pdf.set_font("Helvetica", size=9)

        # Constrói pares (label, value) já formatados
        pairs = []
        for label, _ in qs:
            val = responses.get(label, "")
            if val is True:
                val = "Sim"
            elif val is False:
                val = "Não"
            pairs.append((_safe(label), _safe(str(val))))

        # Escreve sempre 2 perguntas por linha
        n_rows = ceil(len(pairs) / 2)
        idx = 0
        for _ in range(n_rows):
            for _in_group in range(2):     # duas “perguntas” por linha
                if idx < len(pairs):
                    lab, val = pairs[idx]
                    pdf.cell(label_w,  row_h, lab, border=1)
                    pdf.cell(value_w,  row_h, val, border=1)
                    idx += 1
                else:
                    # preenche células vazias se a quantidade for ímpar
                    pdf.cell(label_w, row_h, "", border=1)
                    pdf.cell(value_w, row_h, "", border=1)
            pdf.ln(row_h)                  # próxima linha
        pdf.ln(1)                          # espaço extra entre seções

    # ░░░ Salva em bytes ░░░
    raw = pdf.output(dest="S")
    return bytes(raw) if isinstance(raw, (bytes, bytearray)) else str(raw).encode("latin-1")



# ░░░ Execução direta ░░░────────────────────────────────────────────────────────
if __name__ == "__main__":
    # Pequeno demo local – executa somente se estiver rodando `streamlit run utils.py`
    if st is None:
        raise SystemExit("Execute via `streamlit run utils.py` para visualizar o formulário.")

    st.title("Coleta de Óleo – Demo Utilitário")

    resps = build_form_and_get_responses()

    if st.button("Salvar e Gerar PDF"):
        if not resps.get("n.º da Amostra"):
            st.error("Por favor, preencha o número da amostra!")
        else:
            save_to_sheets(resps)
            pdf_bytes = generate_pdf(resps)
            st.success("Dados salvos! Faça o download do PDF abaixo.")
            st.download_button(
                "Baixar PDF",
                data=pdf_bytes,
                file_name=f"amostra_{resps.get('n.º da Amostra')}.pdf",
                mime="application/pdf",
            )
