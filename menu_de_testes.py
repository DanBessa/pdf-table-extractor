import base64
import contextlib
import csv
import ctypes
import hashlib
import importlib
import inspect
import io
import json
import os
import re
import subprocess
import sys
import threading
import time
import tkinter as tk
from datetime import datetime as dt_cls, timedelta
from decimal import Decimal, InvalidOperation
from tkinter import filedialog, messagebox

from cryptography.exceptions import InvalidSignature
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import padding
import customtkinter as ctk
from customtkinter import CTkImage
import openpyxl
from openpyxl import Workbook
import pandas as pd
from packaging import version
import pdfplumber
from PIL import Image, ImageTk, ImageFilter
import requests

# ==========================================
# 🔒 CONTROLE DE ACESSO E REGRAS
# ==========================================
CANAL_BUILD = "escritorio"

EXIGIR_LOGIN_NA_INICIALIZACAO = False
CURRENT_VERSION = "5.6"
GITHUB_REPO = "DanBessa/pdf-table-extractor"
MODO_DESENVOLVIMENTO = False

URL_USUARIOS_ONLINE = "https://raw.githubusercontent.com/DanBessa/licencas-conversor/main/usuarios.json"

GITHUB_REPO_RUBRICAS = "DanBessa/rubricaupdate"
GITHUB_FILE_RUBRICAS = "rubricas_folha.json"
GITHUB_BRANCH_RUBRICAS = "main"

URLS_GITHUB_RUBRICAS = [
    f"https://raw.githubusercontent.com/{GITHUB_REPO_RUBRICAS}/refs/heads/{GITHUB_BRANCH_RUBRICAS}/{GITHUB_FILE_RUBRICAS}",
    f"https://raw.githubusercontent.com/{GITHUB_REPO_RUBRICAS}/{GITHUB_BRANCH_RUBRICAS}/{GITHUB_FILE_RUBRICAS}"
]

NOME_ARQUIVO_RUBRICAS = "rubricas_folha.json"

# ==========================================
# 🎨 DESIGN SYSTEM & PALETA PROFISSIONAL
# ==========================================
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

THEME = {
    "bg_root": "#0F1115",
    "bg_card": "#181B22",
    "bg_card_hover": "#1F232C",
    "border": "#272D3B",
    "input_bg": "#0D0F13",
    "text_main": "#F8FAFC",
    "text_muted": "#94A3B8",
    "text_dim": "#64748B",
    "primary": "#2563EB",
    "primary_hover": "#1D4ED8",
    "success": "#2563EB",
    "success_hover": "#1D4ED8",
    "warning": "#F59E0B",
    "danger": "#EF4444",
    "danger_hover": "#DC2626",
    "neutral_btn": "#242A38",
    "neutral_hover": "#31394D"
}

FONTS = {
    "brand": ("Segoe UI", 20, "bold"),
    "sub_brand": ("Segoe UI", 11),
    "section": ("Segoe UI", 13, "bold"),
    "body": ("Segoe UI", 11),
    "body_bold": ("Segoe UI", 11, "bold"),
    "code": ("Consolas", 11),
    "status": ("Segoe UI", 10)
}

# ==========================================
# 🛠️ GESTÃO DE DIRETÓRIOS E ARQUIVOS
# ==========================================
def is_frozen():
    return getattr(sys, 'frozen', False)

def get_app_dir():
    if is_frozen():
        return os.path.dirname(os.path.realpath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))

def get_bundle_dir():
    return getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))

def get_storage_dir():
    if sys.platform.startswith('win'):
        base_dir = os.getenv('APPDATA') or os.path.expanduser('~')
        storage_dir = os.path.join(base_dir, 'FolhaBank')
    else:
        storage_dir = os.path.join(os.path.expanduser('~'), '.folhabank')
    
    os.makedirs(storage_dir, exist_ok=True)
    return storage_dir

def desocultar_arquivo_windows(caminho_arquivo):
    if sys.platform.startswith('win') and os.path.exists(caminho_arquivo):
        try:
            FILE_ATTRIBUTE_NORMAL = 0x80
            ctypes.windll.kernel32.SetFileAttributesW(caminho_arquivo, FILE_ATTRIBUTE_NORMAL)
        except Exception:
            pass

def ocultar_arquivo_windows(caminho_arquivo):
    if sys.platform.startswith('win') and os.path.exists(caminho_arquivo):
        try:
            FILE_ATTRIBUTE_HIDDEN = 0x02
            ctypes.windll.kernel32.SetFileAttributesW(caminho_arquivo, FILE_ATTRIBUTE_HIDDEN)
        except Exception:
            pass

conversores_path = os.path.join(get_bundle_dir(), 'conversores')
if conversores_path not in sys.path:
    sys.path.insert(0, conversores_path)

ARQUIVO_LICENCA = os.path.join(get_storage_dir(), 'licenca.dat')
ARQUIVO_SESSAO = os.path.join(get_storage_dir(), 'session.json')
ARQUIVO_RUBRICAS = os.path.join(get_storage_dir(), NOME_ARQUIVO_RUBRICAS)
ARQUIVO_TOKEN = os.path.join(get_storage_dir(), 'github_token.txt')
ARQUIVO_CONFIG_OFX = os.path.join(get_storage_dir(), 'config_contas.json')

CHAVE_PUBLICA_PEM = """
-----BEGIN PUBLIC KEY-----
MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAn38GZoY4+//oDoA7djrd
z5IuRswvGgcJWY6TIIn5Z5cMP7EKphxHUjf20xVGlJQRU6X5KenLNWCLUn/tG3ge
GEdvLyuHI4rw6Lkgua3mzopSdtlP9e29FI3wv5ZHeo08giOT1DvWKG3idRNctAoZ
gC4WIaY3a172ZF0Wig1GdXho/LlvZJutxe10d1mW/wYHYSbpGvFZ2AL4NpylU6ni
noGXyRy+gH5rB/t8htcT2/qroTyh3JB1gH7mXjnEVnt8M2e61cT62DCi/YM44HwB
m2JNNffvxZx4wU1lwEPp9eqgqEYF0I+EVuX4ERN20k5XpFeQ+q1Q4Q3FD2V36ncj
1wIDAQAB
-----END PUBLIC KEY-----
"""

RUBRICAS_BASE = {
    "1": "00001", "100": "00020", "9179": "00014", "3": "00004", "28": "00004",
    "29": "00004", "64": "00016", "8169": "00016", "805": "00037", "806": "00037",
    "815": "00037", "816": "00037", "818": "00037", "819": "00037", "20": "01220",
    "206": "01220", "246": "00074", "208": "00074", "288": "01221", "305": "00003",
    "209": "00003", "995": "00019", "992": "00134", "9754": "00178", "220": "00103",
    "207": "00103", "937": "00140", "8111": "00139", "201": "00139", "262": "00139",
    "202": "00139", "317": "00139", "213": "00139", "320": "00176", "321": "00176",
    "322": "00176", "323": "00176", "324": "00176", "326": "00176", "328": "00176",
    "331": "00176", "9750": "00176", "9752": "00176", "998": "00105", "812": "00106",
    "826": "00108", "989": "00108", "843": "00113", "836": "00041", "821": "00041",
    "999": "00111", "942": "00111", "988": "00170"
}

CONVERTERS = {
    "itau": {
        "nome": "Banco Itaú", "icons": "itaulogo.png", "aba": "pdf", "type": "model_choice",
        "preview": "itau_mod1.png",
        "model_config": {
            "titulo": "Modelo Banco Itaú",
            "label": "Selecione o modelo do extrato do Banco Itaú:",
            "opcoes": {
                "modelo1": "Modelo 1 (Extrato Mensal Consolidado)",
                "modelo2": "Modelo 2 (Itaú BBA / Empresas com CNPJ)",
            },
            "previews": {
                "modelo1": "itau_mod1.png",
                "modelo2": "itau_mod2.png",
                "modelo3": "itau_mod3.png"
            }
        }
    },
    "bb": {
        "nome": "Banco do Brasil", "icons": "bblogo.png", "aba": "pdf", "type": "model_choice",
        "preview": "bb_mod1.png",
        "model_config": {
            "titulo": "Modelo Banco do Brasil",
            "label": "Selecione o modelo do extrato do Banco do Brasil:",
            "opcoes": {"modelo1": "Modelo 1 (Com Cabeçalho)", "modelo2": "Modelo 2 (Extrato Simples)"},
            "previews": {
                "modelo1": "bb_mod1.png",
                "modelo2": "bb_mod2.png"
            }
        }
    },
    "inter": {"nome": "Banco Inter", "icons": "inter.png", "aba": "pdf", "type": "simple_run", "module": "conversor_inter", "function": "iniciar_processamento", "preview": "inter_preview.png"},
    "sicoob": {
        "nome": "Sicoob", "icons": "sicoob2.png", "aba": "pdf", "type": "model_choice",
        "preview": "sicoob_mod1.png",
        "model_config": {
            "titulo": "Modelo Sicoob",
            "label": "Selecione o modelo do extrato do Sicoob:",
            "opcoes": {"modelo1": "Modelo 1", "modelo2": "Modelo 2 (Quebras)", "modelo3": "Modelo 3 (Novo Layout)"},
            "previews": {
                "modelo1": "sicoob_mod1.png",
                "modelo2": "sicoob_mod2.png",
                "modelo3": "sicoob_mod3.png"
            }
        }
    },
    "santander": {
        "nome": "Santander", "icons": "santander.png", "aba": "pdf", "type": "model_choice",
        "preview": "santander_mod1.png",
        "model_config": {
            "titulo": "Modelo Santander",
            "label": "Selecione o modelo do extrato do Santander:",
            "opcoes": {"modelo1": "Modelo 1 (Consolidado)", "modelo2": "Modelo 2"},
            "previews": {
                "modelo1": "santander_mod1.png",
                "modelo2": "santander_mod2.png"
            }
        }
    },
    "safra": {
        "nome": "Banco Safra", "icons": "safra.png", "aba": "pdf", "type": "model_choice",
        "preview": "safra_mod1.png",
        "model_config": {
            "titulo": "Modelo Banco Safra",
            "label": "Selecione o modelo do extrato do Safra:",
            "opcoes": {"modelo1": "Modelo 1 (Simples)", "modelo2": "Modelo 2"},
            "previews": {
                "modelo1": "safra_mod1.png",
                "modelo2": "safra_mod2.png"
            }
        }
    },
    "cef": {
        "nome": "Caixa Econômica", "icons": "cef.png", "aba": "pdf", "type": "model_choice",
        "preview": "cef_mod1.png",
        "model_config": {
            "titulo": "Modelo Caixa Econômica",
            "label": "Selecione o modelo do extrato da Caixa Econômica:",
            "opcoes": {
                "modelo1": "Modelo 1 (Gerenciador Web / Completo)",
                "modelo2": "Modelo 2 (Internet Banking Tradicional)"
            },
            "previews": {
                "modelo1": "cef_mod1.png",
                "modelo2": "cef_mod2.png"
            }
        }
    },
    "bradesco": {"nome": "Bradesco", "icons": "bradesco.png", "aba": "pdf", "type": "simple_run", "module": "conversor_bradesco", "function": "iniciar_processamento", "preview": "bradesco_preview.png"},
    "pagbank": {"nome": "PagBank", "icons": "pagbank.png", "aba": "pdf", "type": "multi_file", "module": "conversor_pagbank", "function": "main", "enabled": True, "preview": "pagbank_preview.png"},
    "c6": {"nome": "C6 Bank", "icons": "c6logo.png", "aba": "pdf", "type": "simple_run", "module": "conversor_c6", "function": "iniciar_processamento", "preview": "c6_preview.png"},
    "banestes": {"nome": "Banestes", "icons": "banestes.png", "aba": "pdf", "type": "simple_run", "module": "conversor_banestes", "function": "iniciar_processamento", "preview": "banestes_preview.png"},
    "paycash": {"nome": "PayCash", "icons": "paycash.png", "aba": "pdf", "type": "simple_run", "module": "conversor_paycash", "function": "iniciar_processamento", "preview": "paycash_preview.png"},
    "sicredi": {"nome": "Sicredi", "icons": "sicredi.png", "aba": "pdf", "type": "simple_run", "module": "conversor_sicredi", "function": "iniciar_processamento", "preview": "sicredi_preview.png"},
    "stone": {"nome": "Stone", "icons": "stone.png", "aba": "pdf", "type": "simple_run", "module": "conversor_stone", "function": "iniciar_processamento", "preview": "stone_preview.png"},
    "mercadopago": {"nome": "Mercado Pago", "icons": "mp3.png", "aba": "pdf", "type": "simple_run", "module": "conversor_mp", "function": "main", "enabled": True, "preview": "mercadopago_preview.png"},
    "ofx": {"nome": "Conversor de Arquivos OFX", "icons": None, "aba": "ofx", "type": "ofx"}
}

# ==========================================
# 📊 MOTOR DE CONVERSÃO OFX
# ==========================================
def extrair_extrato_ofx(caminho_ofx: str) -> str:
    """Lê arquivos OFX/QFX (SGML e XML) e converte para Excel estruturado."""
    conteudo = ""
    for enc in ["utf-8", "latin1", "cp1252", "iso-8859-1"]:
        try:
            with open(caminho_ofx, "r", encoding=enc) as f:
                conteudo = f.read()
            if conteudo: break
        except Exception:
            continue

    if not conteudo:
        raise ValueError(f"Não foi possível abrir o arquivo OFX: {os.path.basename(caminho_ofx)}")

    transactions = []
    blocks = re.findall(r'<STMTTRN>(.*?)(?=<STMTTRN>|</BANKTRANLIST>|\Z)', conteudo, re.DOTALL | re.IGNORECASE)

    for b in blocks:
        def get_tag(tag):
            m = re.search(rf'<{tag}>([^<\r\n]+)', b, re.IGNORECASE)
            return m.group(1).strip() if m else ""

        trntype = get_tag('TRNTYPE').upper()
        dtposted = get_tag('DTPOSTED')
        trnamt = get_tag('TRNAMT')
        fitid = get_tag('FITID')
        checknum = get_tag('CHECKNUM')
        memo = get_tag('MEMO')
        name = get_tag('NAME')

        date_fmt = ""
        m_d = re.match(r'^(\d{4})(\d{2})(\d{2})', dtposted)
        if m_d:
            ano, mes, dia = m_d.group(1), m_d.group(2), m_d.group(3)
            date_fmt = f"{dia}/{mes}/{ano}"

        val_float = limpar_para_float(trnamt)
        
        tipos_debito = ["DEBIT", "PAYMENT", "SRVCHG", "FEE", "CHECK", "XFER"]
        if val_float > 0 and trntype in tipos_debito:
            val_float = -val_float

        tipo = "D" if val_float < 0 else "C"
        desc = memo or name or trntype
        doc = checknum or fitid

        if date_fmt and val_float != 0.0:
            transactions.append({
                'Data': date_fmt,
                'Histórico': re.sub(r'\s+', ' ', desc).strip(),
                'Documento': doc.strip(),
                'Valor': val_float,
                'Tipo': tipo,
                'Saldo': 0.0
            })

    if not transactions:
        raise ValueError("Nenhuma transação financeira localizada dentro do arquivo OFX.")

    df = pd.DataFrame(transactions)
    caminho_saida = os.path.splitext(caminho_ofx)[0] + ".xlsx"
    df.to_excel(caminho_saida, index=False)
    return caminho_saida

# ==========================================
# 📄 FORMATADORES & GERADORES DE TXT
# ==========================================
def limpar_para_float(val) -> float:
    """Converte com segurança qualquer valor para float, preservando negativos."""
    if val is None or val == "" or pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
        
    v_str = str(val).strip()
    v_clean = re.sub(r'[^\d,\.\-−–—]', '', v_str)
    v_clean = v_clean.replace('−', '-').replace('–', '-').replace('—', '-')
    
    is_negative = False
    if '-' in v_clean:
        is_negative = True
        v_clean = v_clean.replace('-', '')

    if ',' in v_clean and '.' in v_clean:
        v_clean = v_clean.replace('.', '').replace(',', '.')
    elif ',' in v_clean:
        v_clean = v_clean.replace(',', '.')
        
    try:
        resultado = float(v_clean)
        return -resultado if is_negative else resultado
    except Exception:
        return 0.0
    
def formatar_moeda_txt(val, apenas_positivo=False) -> str:
    """Formata número para o padrão monetário brasileiro (0,00)."""
    f = limpar_para_float(val)
    if apenas_positivo:
        f = abs(f)
    return f"{f:.2f}".replace('.', ',')

def carregar_dados_extraidos(caminho_arquivo: str) -> pd.DataFrame:
    """Lê arquivos de forma resiliente contra CSVs/HTML salvos como .xlsx ou delimitadores variados."""
    if not os.path.exists(caminho_arquivo) or os.path.getsize(caminho_arquivo) == 0:
        raise ValueError(f"O arquivo gerado está vazio ou inacessível: {os.path.basename(caminho_arquivo)}")

    try:
        return pd.read_excel(caminho_arquivo, engine='openpyxl')
    except Exception:
        pass

    for enc in ['utf-8-sig', 'utf-8', 'latin1', 'cp1252', 'iso-8859-1']:
        for sep in [';', ',', '\t', '|']:
            try:
                df = pd.read_csv(caminho_arquivo, sep=sep, encoding=enc)
                if len(df.columns) >= 2 and len(df) > 0:
                    return df
            except Exception:
                continue

    try:
        tabelas = pd.read_html(caminho_arquivo)
        if tabelas:
            return tabelas[0]
    except Exception:
        pass

    try:
        return pd.read_csv(caminho_arquivo, engine='python', on_bad_lines='skip', encoding='latin1')
    except Exception as e:
        raise ValueError(f"Formato ilegível para o arquivo '{os.path.basename(caminho_arquivo)}': {e}")

def normalizar_dataframe_extrato(df: pd.DataFrame) -> pd.DataFrame:
    """Padroniza qualquer saída de conversor para as 4 colunas canônicas com leitura inteligente."""
    if df is None or df.empty:
        return pd.DataFrame(columns=["Data", "Histórico", "Valor", "Tipo"])

    df_norm = df.copy()

    # 1. Recuperar cabeçalho perdido se pandas leu colunas como "Unnamed: 0"
    if any("UNNAMED" in str(c).upper() for c in df_norm.columns):
        for idx in range(min(5, len(df_norm))):
            row_vals = [str(x).upper() for x in df_norm.iloc[idx].values]
            if (any("DATA" in v or "DT" in v for v in row_vals) and 
                any("VALOR" in v or "VR" in v for v in row_vals)):
                df_norm.columns = df_norm.iloc[idx]
                df_norm = df_norm[idx+1:].reset_index(drop=True)
                break

    mapa_colunas = {}
    for col in df_norm.columns:
        c_upper = str(col).strip().upper()
        
        if "DATA" in c_upper or "DT" in c_upper or "LANÇAMENTO" in c_upper:
            if "Data" not in mapa_colunas.values():
                mapa_colunas[col] = "Data"
        elif any(k in c_upper for k in ["HIST", "LANÇ", "LANC", "DESC", "MEMO", "DOCTO", "ORIGEM", "DETALHE"]):
            if "Histórico" not in mapa_colunas.values():
                mapa_colunas[col] = "Histórico"
        elif any(k in c_upper for k in ["VALOR", "VR", "SAIDA", "ENTRADA", "SAÍDA", "IMPORTANCIA", "QUANTIA"]):
            if "Valor" not in mapa_colunas.values():
                mapa_colunas[col] = "Valor"
        elif any(k in c_upper for k in ["TIPO", "D/C", "C/D", "DC", "NATUREZA", "OPERAÇÃO"]):
            if "Tipo" not in mapa_colunas.values():
                mapa_colunas[col] = "Tipo"

    # Fallback cego caso nenhuma coluna tenha sido mapeada, mas existam colunas suficientes
    if "Data" not in mapa_colunas.values() and len(df_norm.columns) >= 3:
        cols = list(df_norm.columns)
        mapa_colunas[cols[0]] = "Data"
        mapa_colunas[cols[1]] = "Histórico"
        mapa_colunas[cols[2]] = "Valor"
        if len(cols) > 3:
            mapa_colunas[cols[3]] = "Tipo"

    df_norm.rename(columns=mapa_colunas, inplace=True)

    if "Data" not in df_norm.columns:
        df_norm["Data"] = ""
    if "Histórico" not in df_norm.columns:
        df_norm["Histórico"] = "LANCAMENTO BANCARIO"
    if "Valor" not in df_norm.columns:
        df_norm["Valor"] = 0.0

    df_norm["Valor"] = df_norm["Valor"].apply(limpar_para_float)

    if "Tipo" not in df_norm.columns:
        df_norm["Tipo"] = df_norm["Valor"].apply(lambda v: "D" if v < 0 else "C")
    else:
        df_norm["Tipo"] = df_norm["Tipo"].fillna("").astype(str).str.strip().str.upper()
        df_norm["Tipo"] = df_norm.apply(
            lambda r: ("D" if r["Valor"] < 0 else "C") if r["Tipo"] not in ["D", "C"] else r["Tipo"],
            axis=1
        )

    df_norm["Histórico"] = df_norm["Histórico"].astype(str).apply(
        lambda h: re.sub(r'\s+', ' ', h.replace('|', ' ').replace('nan', '')).strip()
    )
    
    # Formata data garantindo que não tenha strings inválidas do pandas
    df_norm["Data"] = df_norm["Data"].astype(str).apply(
        lambda d: "" if d.lower() in ["nan", "nat", "none", ""] else d.strip()
    )

    return df_norm[["Data", "Histórico", "Valor", "Tipo"]]

def gerar_txt_alterdata(df: pd.DataFrame, caminho_txt: str, config_contas: dict = None):
    config = config_contas or {}
    conta_banco = str(config.get("banco", "") or config.get("conta_banco", "")).strip()
    conta_caixa = str(config.get("caixa", "") or config.get("conta_contrapartida", "")).strip()
    
    conta_aplicacao = str(config.get("aplicacao", "")).strip()
    conta_tarifa = str(config.get("tarifa", "")).strip()
    conta_rendimento = str(config.get("rendimento", "")).strip()
    conta_cartao = str(config.get("cartao", "")).strip()
    
    palavras_custom = str(config.get("palavras_cartao", "")).strip().upper()
    custom_regex = ""
    if palavras_custom:
        termos = [re.escape(t.strip()) for t in palavras_custom.split(',') if t.strip()]
        if termos: custom_regex = "|" + "|".join(termos)
    
    cod_hist_padrao = str(config.get("cod_historico", "") or "25").strip()

    with open(caminho_txt, mode='w', newline='', encoding='latin1', errors='replace') as f:
        writer = csv.writer(f, delimiter=',', quoting=csv.QUOTE_ALL)
        
        for _, row in df.iterrows():
            d = str(row.get("Data", "") or "").strip()
            h = str(row.get("Histórico", "") or "").strip()
            
            f_val = limpar_para_float(row.get("Valor", 0.0))
            tipo_raw = str(row.get("Tipo", "") or "").strip().upper()
            if not tipo_raw or tipo_raw not in ["D", "C"]:
                tipo_raw = "D" if f_val < 0 else "C"
                
            v_str = formatar_moeda_txt(f_val, apenas_positivo=True)
            h_upper = h.upper()
            
            # 1. CORREÇÃO DE POLARIDADE (Entrada vs Saída)
            if any(p in h_upper for p in ["PAGAMENTO", "PAGO", "TARIFA", "TAXA", "COMPRA", "ENVIADO", "PIX_DEB", "DEB AUT", "DÉBITO AUT", "LIQUIDACAO BOLETO", "LIQUIDAÇÃO BOLETO", "SICREDI DEBITO"]):
                tipo_raw = "D"
            else:
                regex_entradas = r'\b(RECEBIMENTO|RECEBIDO|CREDITO|CRÉDITO|RESGATE|DEPOSITO|PIX_CRED|REND PAGO|RENDIMENTO|REND|SICREDI CREDITO' + custom_regex + r')\b'
                if re.search(regex_entradas, h_upper):
                    tipo_raw = "C"

            # 2. INTELIGÊNCIA DE CONTAS
            conta_contra = conta_caixa
            is_rendimento = bool(re.search(r'\b(RENDIMENTO|REND PAGO|REND)\b', h_upper))
            is_tarifa = bool(re.search(r'\b(TARIFA|TAXA|IOF|ENCARGO|JUROS|MANUTEN|CUSTO|MENSALIDADE|CESTA|TAR PLANO|PLANO ADAPT|TAR)\b', h_upper))
            is_aplic_resg = bool(re.search(r'\b(APLICACAO|APLICAÇÃO|APLIC|APL|RESGATE|RESG|RES|INVESTIMENTO|CDB|POUPANCA)\b', h_upper))
            regex_cartoes = r'\b(REDE|ALELO|AMEX|VISA|MASTERCARD|MASTER|CREDICARD|ELO|HIPERCARD|BANESCARD|GETNET|CR ANTECIPAÇÃO|RECEBIMENTO VENDAS|SICREDI CREDITO|SICREDI DEBITO' + custom_regex + r')\b'
            is_cartao = bool(re.search(regex_cartoes, h_upper))

            if is_tarifa and conta_tarifa: conta_contra = conta_tarifa
            elif is_rendimento and conta_rendimento: conta_contra = conta_rendimento
            elif is_aplic_resg and conta_aplicacao: conta_contra = conta_aplicacao
            elif is_cartao and conta_cartao: conta_contra = conta_cartao

            # 3. DEFINIÇÃO CONTÁBIL
            if tipo_raw == "C":
                debito = conta_banco
                credito = conta_contra
            else:
                debito = conta_contra
                credito = conta_banco

            writer.writerow(["", debito, credito, d, v_str, cod_hist_padrao, h, ""])

def gerar_txt_dominio(df: pd.DataFrame, caminho_txt: str, config_contas: dict = None):
    """
    Gera o TXT formatado especificamente para importação no Domínio Sistemas.
    Formato: Data;Debito;Credito;Valor;CodHistorico;Historico_Sem_Espaco;;CodEmpresa;;
    """
    config = config_contas or {}
    conta_banco = str(config.get("banco", "") or config.get("conta_banco", "")).strip()
    conta_caixa = str(config.get("caixa", "") or config.get("conta_contrapartida", "")).strip()
    
    conta_aplicacao = str(config.get("aplicacao", "")).strip()
    conta_tarifa = str(config.get("tarifa", "")).strip()
    conta_rendimento = str(config.get("rendimento", "")).strip()
    conta_cartao = str(config.get("cartao", "")).strip()
    
    palavras_custom = str(config.get("palavras_cartao", "")).strip().upper()
    custom_regex = ""
    if palavras_custom:
        termos = [re.escape(t.strip()) for t in palavras_custom.split(',') if t.strip()]
        if termos: custom_regex = "|" + "|".join(termos)
    
    cod_empresa = str(config.get("codigo_empresa", "")).strip()
    
    cod_hist_padrao_entrada = "108"
    cod_hist_padrao_saida = "109"

    with open(caminho_txt, mode='w', newline='', encoding='latin1', errors='replace') as f:
        for _, row in df.iterrows():
            d = str(row.get("Data", "") or "").strip()
            h_original = str(row.get("Histórico", "") or "").strip()
            
            f_val = limpar_para_float(row.get("Valor", 0.0))
            tipo_raw = str(row.get("Tipo", "") or "").strip().upper()
            if not tipo_raw or tipo_raw not in ["D", "C"]:
                tipo_raw = "D" if f_val < 0 else "C"
                
            v_str = formatar_moeda_txt(f_val, apenas_positivo=True)
            h_upper = h_original.upper()
            
            # 1. CORREÇÃO DE POLARIDADE
            if any(p in h_upper for p in ["PAGAMENTO", "PAGO", "TARIFA", "TAXA", "COMPRA", "ENVIADO", "PIX_DEB", "DEB AUT", "DÉBITO AUT", "LIQUIDACAO BOLETO", "LIQUIDAÇÃO BOLETO", "SICREDI DEBITO"]):
                tipo_raw = "D"
            else:
                regex_entradas = r'\b(RECEBIMENTO|RECEBIDO|CREDITO|CRÉDITO|RESGATE|DEPOSITO|PIX_CRED|REND PAGO|RENDIMENTO|REND|SICREDI CREDITO' + custom_regex + r')\b'
                if re.search(regex_entradas, h_upper):
                    tipo_raw = "C"

            # 2. INTELIGÊNCIA DE CONTAS
            conta_contra = conta_caixa
            is_rendimento = bool(re.search(r'\b(RENDIMENTO|REND PAGO|REND)\b', h_upper))
            is_tarifa = bool(re.search(r'\b(TARIFA|TAXA|IOF|ENCARGO|JUROS|MANUTEN|CUSTO|MENSALIDADE|CESTA|TAR PLANO|PLANO ADAPT|TAR)\b', h_upper))
            is_aplic_resg = bool(re.search(r'\b(APLICACAO|APLICAÇÃO|APLIC|APL|RESGATE|RESG|RES|INVESTIMENTO|CDB|POUPANCA)\b', h_upper))
            regex_cartoes = r'\b(REDE|ALELO|AMEX|VISA|MASTERCARD|MASTER|CREDICARD|ELO|HIPERCARD|BANESCARD|GETNET|CR ANTECIPAÇÃO|RECEBIMENTO VENDAS|SICREDI CREDITO|SICREDI DEBITO' + custom_regex + r')\b'
            is_cartao = bool(re.search(regex_cartoes, h_upper))

            if is_tarifa and conta_tarifa: conta_contra = conta_tarifa
            elif is_rendimento and conta_rendimento: conta_contra = conta_rendimento
            elif is_aplic_resg and conta_aplicacao: conta_contra = conta_aplicacao
            elif is_cartao and conta_cartao: conta_contra = conta_cartao

            # 3. DEFINIÇÃO CONTÁBIL E HISTÓRICO DOMINIO
            if tipo_raw == "C":
                debito = conta_banco
                credito = conta_contra
                cod_historico = cod_hist_padrao_entrada
            else:
                debito = conta_contra
                credito = conta_banco
                cod_historico = cod_hist_padrao_saida
                
            historico_limpo = re.sub(r'[^A-Z0-9\.]', '', h_upper)

            linha = f"{d};{debito};{credito};{v_str};{cod_historico};{historico_limpo};;{cod_empresa};;\n"
            f.write(linha)

# ==========================================
# 🔐 AUTENTICAÇÃO E LICENÇA
# ==========================================
def carregar_token_local():
    if os.path.exists(ARQUIVO_TOKEN):
        try:
            with open(ARQUIVO_TOKEN, "r", encoding="utf-8") as f:
                return f.read().strip()
        except Exception:
            pass
    return ""

def salvar_token_local(token: str):
    desocultar_arquivo_windows(ARQUIVO_TOKEN)
    with open(ARQUIVO_TOKEN, "w", encoding="utf-8") as f:
        f.write(token.strip())
    ocultar_arquivo_windows(ARQUIVO_TOKEN)

def carregar_chave_publica():
    return serialization.load_pem_public_key(CHAVE_PUBLICA_PEM.strip().encode('utf-8'))

def checar_exigencia_login_online():
    headers = {'User-Agent': 'Mozilla/5.0', 'Cache-Control': 'no-cache, no-store, must-revalidate', 'Pragma': 'no-cache'}
    dados = None
    try:
        url_api = f"https://api.github.com/repos/DanBessa/licencas-conversor/contents/usuarios.json?_nocache={int(time.time()*1000)}"
        resp_api = requests.get(url_api, headers=headers, timeout=5)
        if resp_api.status_code == 200:
            conteudo_b64 = resp_api.json().get("content", "")
            if conteudo_b64:
                dados = json.loads(base64.b64decode(conteudo_b64).decode('utf-8'))
    except Exception:
        pass

    if not dados:
        try:
            url_raw = f"{URL_USUARIOS_ONLINE}?t={int(time.time()*1000)}"
            resp_raw = requests.get(url_raw, headers=headers, timeout=5)
            if resp_raw.status_code == 200:
                dados = resp_raw.json()
        except Exception:
            pass

    if isinstance(dados, dict):
        config_sistema = dados.get("_config_sistema", {})
        if CANAL_BUILD in config_sistema and isinstance(config_sistema[CANAL_BUILD], bool):
            return config_sistema[CANAL_BUILD]
        if CANAL_BUILD in config_sistema and isinstance(config_sistema[CANAL_BUILD], dict):
            return config_sistema[CANAL_BUILD].get("exigir_login", False)
        if f"exigir_login_{CANAL_BUILD}" in config_sistema:
            return bool(config_sistema[f"exigir_login_{CANAL_BUILD}"])
        if "exigir_login" in config_sistema:
            val = config_sistema["exigir_login"]
            if isinstance(val, dict):
                return val.get(CANAL_BUILD, False)
            return bool(val)

    return EXIGIR_LOGIN_NA_INICIALIZACAO

def validar_e_obter_dados_licenca(caminho_licenca=ARQUIVO_LICENCA):
    if not os.path.exists(caminho_licenca):
        return None, "Sem Licença", None, "Licença não encontrada."
    try:
        if os.path.getsize(caminho_licenca) == 0:
            desocultar_arquivo_windows(caminho_licenca)
            os.remove(caminho_licenca)
            return None, "Sem Licença", None, "Arquivo de licença vazio."

        with open(caminho_licenca, 'r', encoding='utf-8') as f:
            chave_ativacao = f.read().strip()

        if not chave_ativacao:
            return None, "Sem Licença", None, "Licença vazia."
            
        pacote_bytes = base64.b64decode(chave_ativacao)
        pacote = json.loads(pacote_bytes.decode('utf-8'))
        mensagem = base64.b64decode(pacote['dados'])
        assinatura = base64.b64decode(pacote['assinatura'])
        
        public_key = carregar_chave_publica()
        public_key.verify(
            assinatura, mensagem,
            padding.PSS(mgf=padding.MGF1(hashes.SHA256()), salt_length=padding.PSS.MAX_LENGTH),
            hashes.SHA256()
        )
        dados = json.loads(mensagem.decode('utf-8'))
        
        if dados.get("bloqueado", False):
            return False, "Bloqueado", dados.get('cliente'), "Acesso bloqueado pelo administrador."

        str_expira = dados['expira_em']
        dt_expiracao = None
        for formato in ("%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M:%S"):
            try:
                dt_expiracao = dt_cls.strptime(str_expira, formato)
                break
            except ValueError:
                pass
                
        if not dt_expiracao:
            for formato in ("%Y-%m-%d", "%d/%m/%Y"):
                try:
                    dt_expiracao = dt_cls.strptime(str_expira, formato).replace(hour=23, minute=59, second=59)
                    break
                except ValueError:
                    pass

        if not dt_expiracao:
            return False, "Erro Data", None, f"Formato inválido: '{str_expira}'"

        delta = dt_expiracao - dt_cls.now()
        total_segundos = int(delta.total_seconds())

        if total_segundos <= 0:
            return False, "Expirado", dados.get('cliente'), f"Expirada em {dt_expiracao.strftime('%d/%m/%Y %H:%M')}."

        if delta.days > 0:
            tempo_restante_str = f"{delta.days} dia(s)"
        elif total_segundos >= 3600:
            tempo_restante_str = f"{total_segundos // 3600} hora(s)"
        else:
            tempo_restante_str = f"{max(1, total_segundos // 60)} minuto(s)"

        return True, tempo_restante_str, dados.get('cliente'), f"Válida por mais {tempo_restante_str}."
        
    except InvalidSignature:
        return False, "Assinatura Inválida", None, "Chave RSA incompatível. Renove o usuário pelo Admin."
    except Exception as e:
        return False, "Inválida", None, f"Licença inválida: {e}"

def obter_base_usuarios_remota():
    headers = {'User-Agent': 'Mozilla/5.0', 'Cache-Control': 'no-cache, no-store, must-revalidate', 'Pragma': 'no-cache'}
    try:
        url_api = f"https://api.github.com/repos/DanBessa/licencas-conversor/contents/usuarios.json?_nocache={int(time.time()*1000)}"
        resp_api = requests.get(url_api, headers=headers, timeout=5)
        if resp_api.status_code == 200:
            conteudo_b64 = resp_api.json().get("content", "")
            if conteudo_b64:
                dados = json.loads(base64.b64decode(conteudo_b64).decode('utf-8'))
                return {str(k).strip().lower(): v for k, v in dados.items()}
    except Exception:
        pass

    try:
        url_raw = f"{URL_USUARIOS_ONLINE}?t={int(time.time()*1000)}"
        resp_raw = requests.get(url_raw, headers=headers, timeout=5)
        if resp_raw.status_code == 200:
            dados = resp_raw.json()
            return {str(k).strip().lower(): v for k, v in dados.items()}
    except Exception:
        pass

    return None

def autenticar_online(usuario, senha):
    try:
        usuarios_db = obter_base_usuarios_remota()
        if usuarios_db is None:
            return False, "Servidor de autenticação indisponível."
            
        user_key = usuario.strip().lower()
        if user_key not in usuarios_db:
            return False, f"Usuário '{user_key}' não encontrado."
            
        dados_user = usuarios_db[user_key]
        senha_hash_calculado = hashlib.sha256(senha.strip().encode('utf-8')).hexdigest()
        
        if senha_hash_calculado != dados_user.get("senha_hash"):
            return False, "Senha incorreta."
            
        payload_licenca = dados_user.get("licenca_payload")
        if not payload_licenca or not str(payload_licenca).strip():
            return False, "Conta sem licença ativa. Renove pelo Painel Admin."

        desocultar_arquivo_windows(ARQUIVO_LICENCA)
        with open(ARQUIVO_LICENCA, 'w', encoding='utf-8') as f:
            f.write(str(payload_licenca).strip())
        ocultar_arquivo_windows(ARQUIVO_LICENCA)
            
        valido, tempo_str, _, msg = validar_e_obter_dados_licenca(ARQUIVO_LICENCA)
        if valido:
            desocultar_arquivo_windows(ARQUIVO_SESSAO)
            with open(ARQUIVO_SESSAO, 'w', encoding='utf-8') as f:
                json.dump({"usuario": user_key, "senha_salva": senha.strip()}, f)
            ocultar_arquivo_windows(ARQUIVO_SESSAO)
            return True, f"Acesso liberado! Válido por {tempo_str}."
        else:
            if os.path.exists(ARQUIVO_LICENCA):
                desocultar_arquivo_windows(ARQUIVO_LICENCA)
                os.remove(ARQUIVO_LICENCA)
            return False, f"Licença incompatível:\n{msg}"
            
    except Exception as e:
        return False, f"Erro de conexão: {e}"

def sincronizar_licenca_segundo_plano(app_instance):
    def tarefa():
        try:
            if not os.path.exists(ARQUIVO_SESSAO): return
            with open(ARQUIVO_SESSAO, 'r', encoding='utf-8') as f: sessao = json.load(f)
            user_key = sessao.get("usuario", "").strip().lower()
            if not user_key: return

            usuarios_db = obter_base_usuarios_remota()
            if usuarios_db and user_key in usuarios_db:
                novo_payload = usuarios_db[user_key].get("licenca_payload")
                desocultar_arquivo_windows(ARQUIVO_LICENCA)
                with open(ARQUIVO_LICENCA, 'w', encoding='utf-8') as f: f.write(novo_payload)
                ocultar_arquivo_windows(ARQUIVO_LICENCA)
                app_instance.after(0, app_instance.atualizar_status_licenca_ui)
        except Exception: pass
    threading.Thread(target=tarefa, daemon=True).start()

# ==========================================
# 📚 MOTOR DE RUBRICAS & SYNC COM GITHUB
# ==========================================
def carregar_rubricas_locais():
    if not os.path.exists(ARQUIVO_RUBRICAS):
        salvar_rubricas(RUBRICAS_BASE)
        return dict(RUBRICAS_BASE)
    try:
        with open(ARQUIVO_RUBRICAS, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return dict(RUBRICAS_BASE)

def salvar_rubricas(rubricas_dict):
    desocultar_arquivo_windows(ARQUIVO_RUBRICAS)
    with open(ARQUIVO_RUBRICAS, 'w', encoding='utf-8') as f:
        json.dump(rubricas_dict, f, indent=4, ensure_ascii=False)
    ocultar_arquivo_windows(ARQUIVO_RUBRICAS)

def carregar_rubricas():
    rubricas_locais = carregar_rubricas_locais()
    headers = {'User-Agent': 'Mozilla/5.0', 'Cache-Control': 'no-cache, no-store'}
    rubricas_nuvem = {}
    for url in URLS_GITHUB_RUBRICAS:
        try:
            url_anti_cache = f"{url}?t={int(time.time()*1000)}"
            response = requests.get(url_anti_cache, headers=headers, timeout=4)
            if response.status_code == 200:
                dados = response.json()
                if isinstance(dados, dict) and dados:
                    rubricas_nuvem = dados
                    break
        except Exception:
            continue

    rubricas_unificadas = {**RUBRICAS_BASE, **rubricas_nuvem, **rubricas_locais}
    salvar_rubricas(rubricas_unificadas)
    return rubricas_unificadas

def api_github_enviar_rubricas(rubricas_dict, token):
    if not token:
        return False, "Token não configurado."

    url_api = f"https://api.github.com/repos/{GITHUB_REPO_RUBRICAS}/contents/{GITHUB_FILE_RUBRICAS}"
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
        "User-Agent": "FolhaBankIntegrador"
    }

    sha_atual = None
    try:
        resp_get = requests.get(f"{url_api}?ref={GITHUB_BRANCH_RUBRICAS}&t={int(time.time()*1000)}", headers=headers, timeout=6)
        if resp_get.status_code == 200:
            sha_atual = resp_get.json().get("sha")
    except Exception:
        pass

    conteudo_json = json.dumps(rubricas_dict, indent=4, ensure_ascii=False)
    conteudo_base64 = base64.b64encode(conteudo_json.encode('utf-8')).decode('utf-8')

    payload = {
        "message": f"Update rubricas_folha.json via App [{dt_cls.now().strftime('%d/%m/%Y %H:%M')}]",
        "content": conteudo_base64,
        "branch": GITHUB_BRANCH_RUBRICAS
    }
    if sha_atual:
        payload["sha"] = sha_atual

    try:
        resp_put = requests.put(url_api, headers=headers, json=payload, timeout=10)
        if resp_put.status_code in [200, 201]:
            return True, "Rubricas sincronizadas com o GitHub!"
        return False, f"Falha HTTP {resp_put.status_code}"
    except Exception as e:
        return False, f"Erro de conexão: {e}"

def sincronizar_rubricas_github_segundo_plano(rubricas_dict):
    token = carregar_token_local()
    if token:
        threading.Thread(target=lambda: api_github_enviar_rubricas(rubricas_dict, token), daemon=True).start()

def formatar_valor_folha(valor_str):
    v = valor_str.replace('*', '').strip()
    if ',' in v and '.' in v:
        v = v.replace('.', '')
    elif '.' in v and ',' not in v:
        v = v.replace('.', ',')
    if v.startswith(','):
        v = '0' + v
    if v.endswith(','):
        v = v + '00'
    return v

def obter_lancamento_rubrica(cod_procurado, de_para_rubricas):
    cod_limpo = str(cod_procurado).strip()
    if cod_limpo in de_para_rubricas:
        return str(de_para_rubricas[cod_limpo]).strip()
    if cod_limpo.isdigit():
        cod_sem_zero = str(int(cod_limpo))
        if cod_sem_zero in de_para_rubricas:
            return str(de_para_rubricas[cod_sem_zero]).strip()
        val_int = int(cod_limpo)
        for k, v in de_para_rubricas.items():
            k_str = str(k).strip()
            if k_str.isdigit() and int(k_str) == val_int:
                return str(v).strip()
    return None

def extrair_dados_folha(caminho_pdf, data_lancamento, de_para_rubricas):
    lancamentos = []
    rubricas_ignoradas = {"proventos": [], "descontos": []}
    
    try:
        texto_pdf = ""
        with pdfplumber.open(caminho_pdf) as pdf:
            for page in pdf.pages:
                extracted = page.extract_text()
                if extracted:
                    texto_pdf += extracted + "\n"
    except Exception as e:
        return None, str(e)

    inicio_prov = texto_pdf.find("PROVENTOS")
    inicio_desc = texto_pdf.find("DESCONTOS")
    inicio_inf = texto_pdf.find("INFORMATIVA")
    
    texto_proventos = ""
    texto_descontos = ""
    
    if inicio_prov != -1:
        if inicio_desc != -1:
            texto_proventos = texto_pdf[inicio_prov:inicio_desc]
            texto_descontos = texto_pdf[inicio_desc:inicio_inf] if inicio_inf != -1 else texto_pdf[inicio_desc:]
        else:
            texto_proventos = texto_pdf[inicio_prov:inicio_inf] if inicio_inf != -1 else texto_pdf[inicio_prov:]
    else:
        texto_proventos = texto_pdf
        
    def processar_bloco(texto, tipo):
        matches = []
        padrao_horizontal = r"^(\d+)\s+(.*?)\s+([\d\.,\*]+)$"
        padrao_horizontal_fallback = r"^(\d+)\s+(.*)\s([\d\.,\*]+)$"
        
        for linha in texto.split('\n'):
            linha = linha.strip()
            if not linha: continue
            
            # BLOQUEIO DE INFORMATIVAS
            if '*' in linha:
                continue

            match = re.match(padrao_horizontal, linha)
            if match:
                cod_rubrica = match.group(1).strip()
                descricao_bruta = match.group(2).strip()
                descricao = re.sub(r'\s+\d+\s+[\d:,\.]+$', '', descricao_bruta)
                valor_bruto = match.group(3).strip()
                matches.append((cod_rubrica, descricao, valor_bruto))
            else:
                match = re.match(padrao_horizontal_fallback, linha)
                if match:
                    matches.append((match.group(1).strip(), match.group(2).strip(), match.group(3).strip()))
        
        for cod_rubrica, descricao, valor_bruto in matches:
            valor_limpo = formatar_valor_folha(valor_bruto)
            desc_upper = descricao.upper()
            
            # ---> TRAVA ABSOLUTA: Pega EMPRÉSTIMO por extenso ou abreviado ("EMP.") <---
            if tipo == 'D' and any(palavra in desc_upper for palavra in ["EMPRESTIMO", "EMPRÉSTIMO", "EMP.", " EMP "]):
                lancamentos.append(["140", "", "", data_lancamento, valor_limpo, "", "", ""])
                continue
            
            cod_auto = obter_lancamento_rubrica(cod_rubrica, de_para_rubricas)
            
            if cod_auto:
                lancamentos.append([cod_auto, "", "", data_lancamento, valor_limpo, "", "", ""])
            else:
                item_ignorado = f"{cod_rubrica} - {descricao}"
                if tipo == 'P':
                    rubricas_ignoradas["proventos"].append(item_ignorado)
                elif tipo == 'D':
                    rubricas_ignoradas["descontos"].append(item_ignorado)

    if texto_proventos: processar_bloco(texto_proventos, 'P')
    if texto_descontos: processar_bloco(texto_descontos, 'D')
        
    matches_fgts = re.findall(r"Valor do FGTS:\s*([\d\.,]+)", texto_pdf)
    if matches_fgts:
        valor_bruto_fgts = matches_fgts[-1].strip()
        if valor_bruto_fgts not in ["0,00", "0.00", "0"]:
            valor_limpo_fgts = formatar_valor_folha(valor_bruto_fgts)
            lancamentos.append(["101", "", "", data_lancamento, valor_limpo_fgts, "", "", ""])
            
    return lancamentos, rubricas_ignoradas

def gerar_txt_folha(dados, caminho_saida):
    with open(caminho_saida, mode='w', newline='', encoding='utf-8') as arquivo:
        writer = csv.writer(arquivo, delimiter=',', quoting=csv.QUOTE_ALL)
        for linha in dados:
            writer.writerow(linha)

# ==========================================
# 🪟 MODAIS DO SISTEMA
# ==========================================
class JanelaSelecaoModeloModal(ctk.CTkToplevel):
    def __init__(self, master, model_config):
        super().__init__(master)
        self.modelo_escolhido = None
        self.title(model_config.get("titulo", "Seleção de Modelo"))
        self.geometry("640x380")
        self.resizable(False, False)
        self.configure(fg_color=THEME["bg_root"])
        
        self.transient(master)
        self.attributes("-topmost", True)

        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=25, pady=(20, 10))
        ctk.CTkLabel(header, text=model_config.get("titulo", "Seleção de Modelo"), font=FONTS["brand"], text_color=THEME["text_main"]).pack(anchor="w")
        ctk.CTkLabel(header, text="Passe o mouse ou clique em '👁️ Exemplo' para identificar o modelo:", font=FONTS["body"], text_color=THEME["text_muted"]).pack(anchor="w")

        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=25, pady=5)

        opcoes = model_config.get("opcoes", {})
        previews = model_config.get("previews", {})

        for key, label in opcoes.items():
            card = ctk.CTkFrame(
                body, fg_color=THEME["neutral_btn"], corner_radius=8,
                border_width=1, border_color=THEME["border"], height=48
            )
            card.pack(fill="x", pady=4)

            btn_modelo = ctk.CTkButton(
                card, text=f"  {label}", anchor="w",
                fg_color="transparent", hover_color=THEME["neutral_hover"], font=FONTS["body_bold"],
                text_color=THEME["text_main"], height=44, corner_radius=8, cursor="hand2",
                command=lambda k=key: self._selecionar(k)
            )
            btn_modelo.pack(side="left", fill="both", expand=True, padx=(4, 0), pady=2)

            nome_img = previews.get(key, f"{key}.png")
            caminho_img = resolver_caminho_preview(nome_img) or resolver_caminho_preview(f"{key}.png")

            if caminho_img:
                btn_preview = ctk.CTkButton(
                    card, text="👁️ Exemplo", width=88, height=30,
                    font=("Segoe UI", 10, "bold"), fg_color="#181B22", hover_color="#2D3442",
                    border_width=1, border_color="#334155", corner_radius=6, cursor="hand2"
                )
                btn_preview.pack(side="right", padx=(0, 8), pady=7)
                HoverAndClickPreview(btn_preview, caminho_img, titulo=label)

        self.protocol("WM_DELETE_WINDOW", self._cancelar)
        self.lift()
        self.focus_force()
        self.grab_set()

    def _selecionar(self, key):
        self.modelo_escolhido = key
        self.grab_release()
        self.destroy()

    def _cancelar(self):
        self.modelo_escolhido = None
        self.grab_release()
        self.destroy()
        
class JanelaCadastroRubricaModal(ctk.CTkToplevel):
    def __init__(self, master, codigo_preenchido="", callback_atualizacao=None):
        super().__init__(master)
        self.title("Cadastrar Rubrica")
        self.geometry("420x290")
        self.resizable(False, False)
        self.configure(fg_color=THEME["bg_root"])
        self.transient(master)
        self.grab_set()

        self.callback_atualizacao = callback_atualizacao

        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=25, pady=(20, 10))
        ctk.CTkLabel(header, text="Cadastrar Nova Rubrica", font=("Segoe UI", 16, "bold"), text_color=THEME["text_main"]).pack(anchor="w")
        ctk.CTkLabel(header, text="Vincule o código do PDF com o código do sistema contábil", font=("Segoe UI", 11), text_color=THEME["text_muted"]).pack(anchor="w")

        frame = ctk.CTkFrame(self, fg_color="transparent")
        frame.pack(fill="x", padx=25, pady=5)

        ctk.CTkLabel(frame, text="Código no PDF:", font=FONTS["body_bold"], text_color=THEME["text_main"]).pack(anchor="w")
        self.ent_cod = ctk.CTkEntry(frame, height=36, fg_color=THEME["input_bg"], border_color=THEME["border"], border_width=1)
        self.ent_cod.pack(fill="x", pady=(3, 10))
        if codigo_preenchido:
            self.ent_cod.insert(0, codigo_preenchido)

        ctk.CTkLabel(frame, text="Lançamento Automático:", font=FONTS["body_bold"], text_color=THEME["text_main"]).pack(anchor="w")
        self.ent_auto = ctk.CTkEntry(frame, height=36, fg_color=THEME["input_bg"], border_color=THEME["border"], border_width=1)
        self.ent_auto.pack(fill="x", pady=(3, 15))
        if codigo_preenchido:
            self.ent_auto.focus()

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=25, pady=(0, 15))

        ctk.CTkButton(
            btn_frame, text="Salvar Registro", fg_color=THEME["success"], hover_color=THEME["success_hover"],
            font=FONTS["body_bold"], height=38, corner_radius=8, command=self.salvar
        ).pack(side="left", fill="x", expand=True, padx=(0, 6))

        ctk.CTkButton(
            btn_frame, text="Cancelar", fg_color=THEME["neutral_btn"], hover_color=THEME["neutral_hover"],
            font=FONTS["body_bold"], height=38, corner_radius=8, command=self.destroy
        ).pack(side="right", fill="x", expand=True, padx=(6, 0))

    def salvar(self):
        cod = self.ent_cod.get().strip()
        cod_auto = self.ent_auto.get().strip()

        if not cod:
            messagebox.showwarning("Atenção", "Preencha o código do PDF.", parent=self)
            return

        try:
            rubricas_atuais = carregar_rubricas_locais()
            rubricas_atuais[cod] = cod_auto
            salvar_rubricas(rubricas_atuais)
            sincronizar_rubricas_github_segundo_plano(rubricas_atuais)

            messagebox.showinfo("Sucesso", f"Rubrica '{cod}' vinculada a '{cod_auto}'!", parent=self)
            if self.callback_atualizacao:
                self.callback_atualizacao()
            self.destroy()
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao salvar rubrica:\n{e}", parent=self)

class JanelaGerenciarRubricasModal(ctk.CTkToplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Gerenciador de Rubricas")
        self.geometry("560x580")
        self.resizable(False, False)
        self.configure(fg_color=THEME["bg_root"])
        self.transient(master)
        self.grab_set()

        self._criar_layout()

    def _criar_layout(self):
        top_bar = ctk.CTkFrame(self, fg_color="transparent")
        top_bar.pack(fill="x", padx=20, pady=(20, 10))

        self.entry_busca = ctk.CTkEntry(
            top_bar, placeholder_text="🔍 Pesquisar por código ou lançamento...", height=36,
            fg_color=THEME["input_bg"], border_color=THEME["border"], border_width=1, corner_radius=8
        )
        self.entry_busca.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.entry_busca.bind("<KeyRelease>", lambda e: self.recarregar_lista())

        ctk.CTkButton(
            top_bar, text="➕ Nova", width=80, height=36, fg_color=THEME["success"], hover_color=THEME["success_hover"],
            font=FONTS["body_bold"], corner_radius=8, command=lambda: JanelaCadastroRubricaModal(self, callback_atualizacao=self.recarregar_lista)
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            top_bar, text="🔄 Nuvem", width=85, height=36, fg_color=THEME["neutral_btn"], hover_color=THEME["neutral_hover"],
            font=FONTS["body_bold"], corner_radius=8, command=self._sincronizar_nuvem
        ).pack(side="left")

        self.lbl_contagem = ctk.CTkLabel(self, text="Carregando rubricas...", font=FONTS["body"], text_color=THEME["text_muted"])
        self.lbl_contagem.pack(anchor="w", padx=22, pady=(0, 8))

        self.scroll = ctk.CTkScrollableFrame(
            self, label_text="Tabela de Mapeamento (Código PDF ➜ Lançamento Automático)",
            fg_color=THEME["bg_card"], corner_radius=10, border_width=1, border_color=THEME["border"]
        )
        self.scroll.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        self.recarregar_lista()

    def _sincronizar_nuvem(self):
        carregar_rubricas()
        self.recarregar_lista()
        messagebox.showinfo("Nuvem", "Rubricas sincronizadas com o repositório remoto!", parent=self)

    def recarregar_lista(self):
        for w in self.scroll.winfo_children():
            w.destroy()

        rubricas = carregar_rubricas_locais()
        busca = self.entry_busca.get().strip().lower() if hasattr(self, 'entry_busca') else ""

        exibidos = 0
        for cod, cod_auto in sorted(rubricas.items(), key=lambda x: str(x[0])):
            cod_str = str(cod).lower()
            auto_str = str(cod_auto).lower()

            if busca and (busca not in cod_str and busca not in auto_str):
                continue

            exibidos += 1
            card = ctk.CTkFrame(self.scroll, fg_color=THEME["input_bg"], height=42, corner_radius=6, border_width=1, border_color=THEME["border"])
            card.pack(fill="x", pady=3, padx=2)

            ctk.CTkLabel(card, text=f"PDF: {cod}", font=FONTS["body_bold"], text_color=THEME["text_main"]).pack(side="left", padx=12)
            ctk.CTkLabel(card, text=f"➜ Lanç. Auto: {cod_auto}", font=FONTS["body_bold"], text_color=THEME["success"]).pack(side="left", padx=12)

            btn_del = ctk.CTkButton(
                card, text="🗑️", width=32, height=28, fg_color=THEME["danger"], hover_color=THEME["danger_hover"],
                corner_radius=6, command=lambda c=cod: self.remover_rubrica(c)
            )
            btn_del.pack(side="right", padx=8)

        total_txt = f"{exibidos} registros exibidos (Total: {len(rubricas)})" if busca else f"Total de {len(rubricas)} rubricas mapeadas"
        self.lbl_contagem.configure(text=total_txt)

        if exibidos == 0:
            ctk.CTkLabel(self.scroll, text="Nenhuma rubrica localizada.", text_color=THEME["text_dim"]).pack(pady=30)

    def remover_rubrica(self, cod):
        rubricas = carregar_rubricas_locais()
        if cod in rubricas:
            del rubricas[cod]
            salvar_rubricas(rubricas)
            sincronizar_rubricas_github_segundo_plano(rubricas)
            self.recarregar_lista()

class JanelaResultadoFolhaModal(ctk.CTkToplevel):
    def __init__(self, master, titulo, mensagem_info, ignoradas_dict, callback_gerar=None):
        super().__init__(master)
        self.callback_gerar = callback_gerar
        self.title(titulo)
        self.geometry("680x560")
        self.resizable(False, False)
        self.configure(fg_color=THEME["bg_root"])
        self.transient(master)
        self.grab_set()

        ctk.CTkLabel(self, text=titulo, font=FONTS["brand"], text_color=THEME["text_main"]).pack(pady=(20, 5))

        provs = ignoradas_dict.get("proventos", [])
        descs = ignoradas_dict.get("descontos", [])

        txt_resultado = ctk.CTkTextbox(
            self, font=FONTS["code"], wrap="word", height=75,
            fg_color=THEME["bg_card"], border_color=THEME["border"], border_width=1, corner_radius=8
        )
        txt_resultado.pack(fill="x", padx=25, pady=(5, 10))
        txt_resultado.insert("1.0", mensagem_info)
        txt_resultado.configure(state="disabled")

        self.checkbox_widgets = {}

        if provs or descs:
            ctk.CTkLabel(self, text="Selecione as rubricas e digite o lançamento automático para cadastrá-las em lote:", font=FONTS["body_bold"], text_color=THEME["warning"]).pack(anchor="w", padx=25, pady=(0, 5))

            self.scroll_rubricas = ctk.CTkScrollableFrame(
                self, fg_color=THEME["bg_card"], border_color=THEME["border"], border_width=1, corner_radius=8, height=160
            )
            self.scroll_rubricas.pack(fill="x", padx=25, pady=0)

            if provs:
                ctk.CTkLabel(self.scroll_rubricas, text="🟢 PROVENTOS:", font=FONTS["body_bold"], text_color=THEME["success"]).pack(anchor="w", pady=(5, 5), padx=5)
                for p in provs:
                    self._add_checkbox(p)

            if descs:
                ctk.CTkLabel(self.scroll_rubricas, text="🔴 DESCONTOS:", font=FONTS["body_bold"], text_color=THEME["danger"]).pack(anchor="w", pady=(15, 5), padx=5)
                for d in descs:
                    self._add_checkbox(d)

            lote_frame = ctk.CTkFrame(self, fg_color=THEME["input_bg"], corner_radius=8, border_width=1, border_color=THEME["border"])
            lote_frame.pack(fill="x", padx=25, pady=10)

            ctk.CTkLabel(lote_frame, text="Lançamento Automático:", font=FONTS["body_bold"], text_color=THEME["text_main"]).pack(side="left", padx=(15, 5), pady=12)
            self.ent_lote_auto = ctk.CTkEntry(lote_frame, width=120, height=36, fg_color=THEME["bg_card"], border_color=THEME["border"])
            self.ent_lote_auto.pack(side="left", padx=5, pady=12)

            ctk.CTkButton(
                lote_frame, text="✅ Vincular Selecionados", fg_color=THEME["success"], hover_color=THEME["success_hover"],
                font=FONTS["body_bold"], height=36, command=self.cadastrar_lote
            ).pack(side="right", padx=15, pady=12)

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=25, pady=(10, 20), side="bottom")

        ctk.CTkButton(
            btn_frame, text="⚙️ Gerenciar Rubricas", width=170, height=38, fg_color=THEME["primary"],
            hover_color=THEME["primary_hover"], font=FONTS["body_bold"], corner_radius=8,
            command=lambda: JanelaGerenciarRubricasModal(self)
        ).pack(side="left")

        if self.callback_gerar:
            ctk.CTkButton(
                btn_frame, text="📄 Gerar Arquivo", width=140, height=38, fg_color=THEME["success"],
                hover_color=THEME["success_hover"], font=FONTS["body_bold"], corner_radius=8,
                command=self.acao_gerar_arquivo
            ).pack(side="left", padx=10)

        ctk.CTkButton(
            btn_frame, text="Fechar", width=100, height=38, fg_color=THEME["neutral_btn"],
            hover_color=THEME["neutral_hover"], font=FONTS["body_bold"], corner_radius=8,
            command=self.destroy
        ).pack(side="right")

    def _add_checkbox(self, item_str):
        cod_rubrica = item_str.split(" - ")[0].strip()
        chk = ctk.CTkCheckBox(self.scroll_rubricas, text=item_str, font=FONTS["code"], text_color=THEME["text_main"])
        chk.pack(anchor="w", padx=15, pady=4)
        self.checkbox_widgets[cod_rubrica] = chk

    def acao_gerar_arquivo(self):
        pendentes = len(self.checkbox_widgets)
        if pendentes > 0:
            resp = messagebox.askyesno(
                "Atenção - Rubricas Pendentes", 
                f"Ainda existem {pendentes} rubricas na tela que não foram cadastradas.\n\n"
                "Se você prosseguir, elas ficarão de fora do arquivo TXT final.\n\n"
                "Deseja gerar o arquivo mesmo assim?",
                parent=self
            )
            if not resp:
                return
        
        self.destroy()
        if self.callback_gerar:
            self.callback_gerar()

    def cadastrar_lote(self):
        cod_auto = self.ent_lote_auto.get().strip()
        selecionados = [cod for cod, chk in self.checkbox_widgets.items() if chk.get() == 1]

        if not selecionados:
            messagebox.showwarning("Atenção", "Selecione pelo menos uma rubrica marcando as caixinhas.", parent=self)
            return
        if not cod_auto:
            messagebox.showwarning("Atenção", "Preencha o código numérico do lançamento contábil.", parent=self)
            return

        try:
            rubricas_atuais = carregar_rubricas_locais()
            for cod in selecionados: rubricas_atuais[cod] = cod_auto
            salvar_rubricas(rubricas_atuais)
            sincronizar_rubricas_github_segundo_plano(rubricas_atuais)

            messagebox.showinfo("Sucesso", f"{len(selecionados)} rubricas vinculadas!\nElas sumiram da lista abaixo. Pode continuar mapeando.", parent=self)
            
            for cod in selecionados:
                if cod in self.checkbox_widgets:
                    self.checkbox_widgets[cod].destroy()
                    del self.checkbox_widgets[cod]
            
            self.ent_lote_auto.delete(0, 'end')
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao salvar rubricas:\n{e}", parent=self)
            
# ==========================================
# 🪟 MODAL DE LOGIN
# ==========================================
class JanelaLoginDialog(ctk.CTkToplevel):
    def __init__(self, master=None, obrigatorio=False):
        super().__init__(master)
        self.obrigatorio = obrigatorio
        self.sucesso = False
        self.title("Autenticação de Acesso")
        self.geometry("440x480")
        self.resizable(False, False)
        self.configure(fg_color=THEME["bg_root"])
        
        self.transient(master)
        self.grab_set()
        
        if self.obrigatorio:
            self.protocol("WM_DELETE_WINDOW", self._fechar_aplicacao)

        card = ctk.CTkFrame(self, fg_color=THEME["bg_card"], corner_radius=12, border_width=1, border_color=THEME["border"])
        card.pack(fill="both", expand=True, padx=25, pady=25)

        ctk.CTkLabel(card, text="FolhaBank", font=FONTS["brand"], text_color=THEME["text_main"]).pack(pady=(20, 2))
        ctk.CTkLabel(card, text="Autenticação Corporativa", font=FONTS["sub_brand"], text_color=THEME["text_muted"]).pack(pady=(0, 18))

        ctk.CTkLabel(card, text="E-mail / Usuário", font=FONTS["body_bold"], text_color=THEME["text_main"]).pack(anchor="w", padx=30)
        self.entry_user = ctk.CTkEntry(
            card, placeholder_text="seu.email@empresa.com", height=38,
            fg_color=THEME["input_bg"], border_color=THEME["border"], border_width=1, corner_radius=8
        )
        self.entry_user.pack(fill="x", padx=30, pady=(3, 14))
        
        ctk.CTkLabel(card, text="Senha de Acesso", font=FONTS["body_bold"], text_color=THEME["text_main"]).pack(anchor="w", padx=30)
        self.entry_pass = ctk.CTkEntry(
            card, placeholder_text="••••••••", show="*", height=38,
            fg_color=THEME["input_bg"], border_color=THEME["border"], border_width=1, corner_radius=8
        )
        self.entry_pass.pack(fill="x", padx=30, pady=(3, 18))

        self._carregar_credenciais_salvas()

        self.btn_login = ctk.CTkButton(
            card, text="Entrar e Ativar", fg_color=THEME["primary"], hover_color=THEME["primary_hover"],
            font=FONTS["body_bold"], height=42, corner_radius=8, command=self._executar_login
        )
        self.btn_login.pack(fill="x", padx=30, pady=(0, 10))

        self.lbl_status = ctk.CTkLabel(card, text="", font=FONTS["status"], text_color=THEME["warning"])
        self.lbl_status.pack(pady=(0, 10))

        ctk.CTkLabel(
            card, text="Suporte: (31) 9 8277-3128 | danielbessaribeiro@gmail.com",
            font=("Segoe UI", 9), text_color=THEME["text_dim"]
        ).pack(side="bottom", pady=10)

    def _carregar_credenciais_salvas(self):
        if os.path.exists(ARQUIVO_SESSAO):
            try:
                with open(ARQUIVO_SESSAO, 'r', encoding='utf-8') as f:
                    sessao = json.load(f)
                u = sessao.get("usuario", "")
                p = sessao.get("senha_salva", "")
                if u: self.entry_user.insert(0, u)
                if p: self.entry_pass.insert(0, p)
            except Exception:
                pass

    def _fechar_aplicacao(self):
        self.destroy()
        if self.master:
            self.master.destroy()
        sys.exit(0)

    def _executar_login(self):
        user = self.entry_user.get().strip()
        senha = self.entry_pass.get().strip()

        if not user or not senha:
            self.lbl_status.configure(text="Preencha usuário e senha.", text_color=THEME["danger"])
            return

        self.btn_login.configure(state="disabled", text="Autenticando...")
        self.lbl_status.configure(text="Validando credenciais...", text_color=THEME["warning"])
        self.update_idletasks()

        sucesso, msg = autenticar_online(user, senha)
        if sucesso:
            messagebox.showinfo("Sucesso", msg, parent=self)
            self.sucesso = True
            self.destroy()
        else:
            self.btn_login.configure(state="normal", text="Entrar e Ativar")
            self.lbl_status.configure(text=msg, text_color=THEME["danger"])

# ==========================================
# 🧩 COMPONENTES MODERNOS & UPDATE
# ==========================================
class ModernButton(ctk.CTkButton):
    def __init__(self, master, **kwargs):
        anchor = kwargs.pop('anchor', 'center')
        fg_color = kwargs.pop('fg_color', THEME["neutral_btn"])
        hover_color = kwargs.pop('hover_color', THEME["neutral_hover"])
        font = kwargs.pop('font', FONTS["body_bold"])
        height = kwargs.pop('height', 44)
        corner_radius = kwargs.pop('corner_radius', 8)
        border_width = kwargs.pop('border_width', 1)
        border_color = kwargs.pop('border_color', THEME["border"])
        super().__init__(
            master=master, font=font, fg_color=fg_color, hover_color=hover_color,
            text_color=THEME["text_main"], height=height, corner_radius=corner_radius,
            anchor=anchor, border_width=border_width, border_color=border_color,
            border_spacing=10, compound="left", cursor="hand2", **kwargs
        )

class UpdateManager:
    def __init__(self, app_instance):
        self.app = app_instance
        self.current_exe_path = os.path.realpath(sys.executable if is_frozen() else sys.argv[0])
        self.current_exe_dir = get_app_dir()
        
    def cleanup_old_version(self):
        old_exe_path = self.current_exe_path + ".old"
        if os.path.exists(old_exe_path):
            try:
                for _ in range(5):
                    try:
                        os.remove(old_exe_path)
                        break
                    except PermissionError:
                        time.sleep(1)
            except Exception:
                pass

    def check_for_updates(self):
        def run_check():
            try:
                api_url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
                response = requests.get(api_url, timeout=8)
                response.raise_for_status()
                latest_release = response.json()
                latest_version_str = latest_release['tag_name'].lstrip('v').strip()
                if version.parse(latest_version_str) > version.parse(CURRENT_VERSION.strip()):
                    self.app.after(0, lambda: self._show_update_dialog(latest_version_str, latest_release))
                else:
                    self.app.after(0, self.app.show_up_to_date_status)
            except Exception as e:
                print(f"Erro ao checar atualizações: {e}")
        threading.Thread(target=run_check, daemon=True).start()

    def _show_update_dialog(self, new_version, release_info):
        message = (f"Uma nova versão ({new_version}) está disponível.\n"
                   f"Sua versão atual: {CURRENT_VERSION}\n\nDeseja atualizar agora?")
        if messagebox.askyesno("Atualização Disponível", message):
            threading.Thread(target=self._download_and_apply_update, args=(release_info,), daemon=True).start()

    def _download_and_apply_update(self, release_info):
        try:
            asset = next((a for a in release_info.get('assets', []) if a['name'].endswith('.exe')), None)
            if not asset: return
            temp_update_path = os.path.join(self.current_exe_dir, "new_" + os.path.basename(self.current_exe_path))
            self.app.after(0, lambda: self.app.update_status("Baixando atualização...", THEME["warning"]))
            response = requests.get(asset['browser_download_url'], stream=True, timeout=30)
            response.raise_for_status()
            with open(temp_update_path, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192): f.write(chunk)
            self.app.after(0, lambda: self.app.update_status("Atualização pronta. Reiniciando...", THEME["success"]))
            old_exe_path = self.current_exe_path + ".old"
            os.rename(self.current_exe_path, old_exe_path)
            os.rename(temp_update_path, self.current_exe_path)
            subprocess.Popen([self.current_exe_path])
            self.app.after(100, self.app.destroy)
        except Exception as e:
            self.app.after(0, lambda: messagebox.showerror("Erro na Atualização", f"Falha:\n{e}"))
            self.app.after(0, lambda: self.app.update_status("Falha ao atualizar.", THEME["danger"]))

def resolver_caminho_preview(chave_ou_nome):
    if not chave_ou_nome:
        return None

    chave_limpa = str(chave_ou_nome).replace(".png", "").strip().lower()
    
    pastas_busca = [
        os.path.join(get_bundle_dir(), "icons", "bancos"),
        os.path.join(get_app_dir(), "icons", "bancos"),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "icons", "bancos"),
        os.path.join(get_bundle_dir(), "icons"),
        os.path.join(get_app_dir(), "icons"),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "icons")
    ]

    nomes_tentativas = [
        f"{chave_limpa}.png",
        f"{chave_limpa}_preview.png",
        f"{chave_limpa}_mod1.png",
        f"{chave_limpa}_mod2.png",
        f"{chave_limpa}_mod3.png"
    ]

    for pasta in pastas_busca:
        if not os.path.exists(pasta):
            continue
            
        for nome in nomes_tentativas:
            caminho = os.path.join(pasta, nome)
            if os.path.exists(caminho):
                return caminho

        for arq in os.listdir(pasta):
            if arq.lower().startswith(chave_limpa) and arq.lower().endswith((".png", ".jpg", ".jpeg")):
                return os.path.join(pasta, arq)

    return None

class HoverAndClickPreview:
    def __init__(self, widget_botao, caminho_imagem, titulo="Modelo de Extrato"):
        self.widget = widget_botao
        self.caminho_imagem = caminho_imagem
        self.titulo = titulo
        self.popup = None
        self.img_ref = None

        self._vincular_hover(self.widget)
        self.widget.configure(command=self.abrir_modal_zoom)

    def _vincular_hover(self, w):
        w.bind("<Enter>", self.exibir_popup)
        w.bind("<Leave>", self.ocultar_popup)
        for sub in ["_canvas", "_text_label", "_image_label"]:
            if hasattr(w, sub) and getattr(w, sub):
                getattr(w, sub).bind("<Enter>", self.exibir_popup)
                getattr(w, sub).bind("<Leave>", self.ocultar_popup)

    def exibir_popup(self, event=None):
        if not self.caminho_imagem or not os.path.exists(self.caminho_imagem):
            return
        if self.popup is not None:
            return

        x = self.widget.winfo_rootx() + self.widget.winfo_width() + 12
        y = self.widget.winfo_rooty() - 20

        self.popup = tk.Toplevel(self.widget)
        self.popup.wm_overrideredirect(True)
        self.popup.wm_geometry(f"+{x}+{y}")
        self.popup.attributes("-topmost", True)
        self.popup.configure(bg="#272D3B")

        try:
            pil_img = Image.open(self.caminho_imagem)
            pil_img.thumbnail((360, 220), Image.Resampling.LANCZOS)
            self.img_ref = ImageTk.PhotoImage(pil_img)

            frame = tk.Frame(self.popup, bg="#181B22", padx=8, pady=8)
            frame.pack(padx=1, pady=1)

            lbl_tit = tk.Label(
                frame, text=f"Exemplo • {self.titulo}",
                font=("Segoe UI", 9), fg="#94A3B8", bg="#181B22"
            )
            lbl_tit.pack(anchor="w", pady=(0, 6))

            lbl_img = tk.Label(frame, image=self.img_ref, bg="#181B22")
            lbl_img.pack()
        except Exception:
            self.ocultar_popup()

    def ocultar_popup(self, event=None):
        if self.popup is not None:
            try:
                self.popup.destroy()
            except Exception:
                pass
            self.popup = None
            self.img_ref = None

    def abrir_modal_zoom(self):
        self.ocultar_popup()
        if not self.caminho_imagem or not os.path.exists(self.caminho_imagem):
            return

        janela_zoom = ctk.CTkToplevel(self.widget)
        janela_zoom.title(f"Visualização: {self.titulo}")
        janela_zoom.geometry("680x480")
        janela_zoom.configure(fg_color=THEME["bg_root"])
        janela_zoom.transient(self.widget)
        janela_zoom.attributes("-topmost", True)

        try:
            pil_img = Image.open(self.caminho_imagem)
            pil_img.thumbnail((620, 380), Image.Resampling.LANCZOS)
            tk_zoom = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=pil_img.size)

            header = ctk.CTkFrame(janela_zoom, fg_color="transparent")
            header.pack(fill="x", padx=20, pady=(15, 5))
            ctk.CTkLabel(header, text=self.titulo, font=FONTS["section"], text_color=THEME["text_main"]).pack(anchor="w")

            card_img = ctk.CTkFrame(janela_zoom, fg_color=THEME["bg_card"], corner_radius=8, border_width=1, border_color=THEME["border"])
            card_img.pack(fill="both", expand=True, padx=20, pady=(5, 15))

            lbl = ctk.CTkLabel(card_img, image=tk_zoom, text="")
            lbl.image = tk_zoom
            lbl.pack(expand=True, padx=10, pady=10)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao exibir imagem ampliada:\n{e}", parent=janela_zoom)

def extrair_pdf_universal(caminho_pdf: str) -> str:
    ignorar_termos = [
        "SALDO ANTERIOR", "SALDO FINAL", "SALDO ATUAL", "SALDO DISPONÍVEL",
        "SALDO DISPONIVEL", "SALDO DO DIA", "SALDO EM CONTA", "EXTRATO DE CONTA CORRENTE",
        "DATA LANÇAMENTOS", "DATA HISTÓRICO", "VALOR SALDO", "TOTAL APLICAÇÕES",
        "RESGATE AUTOMATICO", "APLICACAO AUTOMATICA", "S A L D O", "FOLHA"
    ]

    transacoes = []
    current_date = None
    ano_detectado = dt_cls.now().strftime("%Y")

    with pdfplumber.open(caminho_pdf) as pdf:
        texto_completo = ""
        for p in pdf.pages:
            texto_completo += (p.extract_text() or "") + "\n"

        m_ano = re.search(r'/\s*(202\d)\b', texto_completo) or re.search(r'\b(202\d)\b', texto_completo)
        if m_ano:
            ano_detectado = m_ano.group(1)

        for page in pdf.pages:
            texto = page.extract_text(layout=True) or ""
            for linha in texto.split("\n"):
                linha_limpa = linha.strip()
                if not linha_limpa:
                    continue

                if any(term in linha_limpa.upper() for term in ignorar_termos):
                    m_d_ign = re.search(r"\b(\d{2}/\d{2}(?:/\d{4})?)\b", linha_limpa)
                    if m_d_ign:
                        d_str = m_d_ign.group(1)
                        current_date = d_str if len(d_str) == 10 else f"{d_str}/{ano_detectado}"
                    continue

                m_data_full = re.search(r"\b(\d{2}/\d{2}/\d{4})\b", linha_limpa)
                m_data_short = re.match(r"^(\d{2}/\d{2})\b", linha_limpa)

                if m_data_full:
                    current_date = m_data_full.group(1)
                elif m_data_short:
                    current_date = f"{m_data_short.group(1)}/{ano_detectado}"

                val_pattern = r'([−–—\-]?\s*(?:R\$\s*)?[\d\.]*,\d{2}\s*(?:\([^)]+\))?\s*[−–—\-]?[CD]?)'
                valores = [m.strip() for m in re.findall(val_pattern, linha_limpa, re.IGNORECASE) if re.search(r'\d+,\d{2}', m)]

                if not valores or not current_date:
                    if transacoes and current_date and not re.search(r'\d+,\d{2}', linha_limpa):
                        texto_extra = linha_limpa.replace('|', ' ')
                        transacoes[-1]["Histórico"] = re.sub(r'\s+', ' ', transacoes[-1]["Histórico"] + " " + texto_extra).strip()
                    continue

                val_str = valores[0] if len(valores) >= 2 else valores[-1]

                hist_bruto = linha_limpa
                if m_data_full:
                    hist_bruto = hist_bruto.replace(m_data_full.group(1), "")
                elif m_data_short:
                    hist_bruto = hist_bruto.replace(m_data_short.group(1), "")

                for v in valores:
                    hist_bruto = hist_bruto.replace(v, "")

                hist_limpo = re.sub(r'\s+', ' ', hist_bruto.replace('|', ' ')).strip()

                f_val = abs(limpar_para_float(val_str))
                h_up = hist_limpo.upper()

                eh_credito = any(c in h_up for c in ["RECEBIDO", "RECEBIDA", "CREDITO", "CRÉDITO", "RESGATE"])
                eh_debito = any(d in h_up for d in ["TARIFA", "ENVIADO", "ENVIADA", "PAGAMENTO", "DEBITO", "DÉBITO", "COMPRA"])

                if eh_credito:
                    tipo = "C"
                    valor_final = f_val
                elif eh_debito or '(-)' in val_str or '(−)' in val_str or '(–)' in val_str or val_str.endswith('-'):
                    tipo = "D"
                    valor_final = -f_val
                else:
                    tipo = "C"
                    valor_final = f_val

                if valor_final != 0.0:
                    transacoes.append({
                        "Data": current_date,
                        "Histórico": hist_limpo if hist_limpo else "LANCAMENTO BANCARIO",
                        "Valor": valor_final,
                        "Tipo": tipo
                    })

    if not transacoes:
        raise ValueError(f"Nenhuma transação foi identificada no PDF '{os.path.basename(caminho_pdf)}'.")

    df = pd.DataFrame(transacoes)[["Data", "Histórico", "Valor", "Tipo"]]
    caminho_saida = os.path.splitext(caminho_pdf)[0] + ".xlsx"
    df.to_excel(caminho_saida, index=False)
    return caminho_saida

# ==========================================
# 🖥️ APLICAÇÃO PRINCIPAL
# ==========================================
class ConversorApp(ctk.CTk):
    def __init__(self):
        super().__init__(fg_color=THEME["bg_root"])
        
        self.login_obrigatorio = checar_exigencia_login_online()
        if self.login_obrigatorio:
            self.withdraw()

        self.update_manager = UpdateManager(self)
        self.update_manager.cleanup_old_version()
        
        if is_frozen() and not MODO_DESENVOLVIMENTO:
            self.update_manager.check_for_updates()
            
        self.title("FolhaBank Integrador")
        self.geometry("560x700")
        self.resizable(False, False)
        self.icons = self._load_icons()
        self._create_widgets()
        
        self.after(50, self._verificar_acesso_inicial)
        sincronizar_licenca_segundo_plano(self)

    def _verificar_acesso_inicial(self):
        if not self.login_obrigatorio:
            self.deiconify()
            self.atualizar_status_licenca_ui()
            return

        valido, tempo_str, cliente, msg = validar_e_obter_dados_licenca(ARQUIVO_LICENCA)
        if valido:
            self.deiconify()
            self.atualizar_status_licenca_ui()
        else:
            modal = JanelaLoginDialog(master=self, obrigatorio=True)
            self.wait_window(modal)
            if modal.sucesso:
                self.deiconify()
                self.atualizar_status_licenca_ui()
            else:
                self.destroy()
                sys.exit(0)

    @contextlib.contextmanager
    def patch_file_dialogs(self, return_path):
        if not return_path:
            yield
            return

        path_str = return_path[0] if isinstance(return_path, (list, tuple)) else str(return_path)
        path_list = list(return_path) if isinstance(return_path, (list, tuple)) else [return_path]

        def fake_askopenfilenames(*args, **kwargs): return path_list
        def fake_askopenfilename(*args, **kwargs): return path_str
        def fake_noop(*args, **kwargs): return "ok"

        import tkinter.messagebox as tk_mb
        import tkinter.filedialog as tk_fd
        
        orig_mb = {
            'showinfo': tk_mb.showinfo,
            'showwarning': tk_mb.showwarning,
            'showerror': tk_mb.showerror
        }
        orig_fd = {
            'askopenfilenames': tk_fd.askopenfilenames,
            'askopenfilename': tk_fd.askopenfilename
        }

        tk_mb.showinfo = fake_noop
        tk_mb.showwarning = fake_noop
        tk_fd.askopenfilenames = fake_askopenfilenames
        tk_fd.askopenfilename = fake_askopenfilename

        for mod_name, mod in list(sys.modules.items()):
            if mod and hasattr(mod, '__dict__'):
                d = mod.__dict__
                if 'showinfo' in d:
                    d['showinfo'] = fake_noop
                if 'showwarning' in d:
                    d['showwarning'] = fake_noop
                if 'messagebox' in d and hasattr(d['messagebox'], 'showinfo'):
                    try:
                        d['messagebox'].showinfo = fake_noop
                        d['messagebox'].showwarning = fake_noop
                    except Exception:
                        pass

        try:
            yield
        finally:
            tk_mb.showinfo = orig_mb['showinfo']
            tk_mb.showwarning = orig_mb['showwarning']
            tk_mb.showerror = orig_mb['showerror']
            tk_fd.askopenfilenames = orig_fd['askopenfilenames']
            tk_fd.askopenfilename = orig_fd['askopenfilename']

    def _load_icons(self):
        icons = {}
        pastas = [
            os.path.join(get_bundle_dir(), "icons", "bancos"),
            os.path.join(get_bundle_dir(), "icons")
        ]
        for key, config in CONVERTERS.items():
            if not config.get("icons"): continue
            for pasta in pastas:
                path = os.path.join(pasta, config['icons'])
                if os.path.exists(path):
                    try:
                        image = Image.open(path)
                        icons[key] = CTkImage(dark_image=image, light_image=image, size=(24, 24))
                        break
                    except Exception as e:
                        print(f"Erro no ícone '{key}': {e}")
        return icons

    def _create_widgets(self):
        container = ctk.CTkFrame(self, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=15, pady=15)
        
        header_frame = ctk.CTkFrame(container, fg_color="transparent")
        header_frame.pack(fill="x", pady=(0, 10))

        title_box = ctk.CTkFrame(header_frame, fg_color="transparent")
        title_box.pack(side="left")

        ctk.CTkLabel(title_box, text="FolhaBank Integrador", font=FONTS["brand"], text_color=THEME["text_main"]).pack(anchor="w")
        ctk.CTkLabel(title_box, text=f"Versão {CURRENT_VERSION}", font=FONTS["sub_brand"], text_color=THEME["text_muted"]).pack(anchor="w")

        self.pill_licenca = ctk.CTkFrame(header_frame, fg_color=THEME["bg_card"], corner_radius=20, border_width=1, border_color=THEME["border"])
        self.pill_licenca.pack(side="right", pady=4)
        
        self.lbl_licenca = ctk.CTkLabel(self.pill_licenca, text="", font=FONTS["status"])
        self.lbl_licenca.pack(padx=12, pady=4)
        self.atualizar_status_licenca_ui()

        if self.login_obrigatorio:
            ctk.CTkButton(
                container, text="👤 Trocar Usuário / Renovar", font=FONTS["status"], fg_color=THEME["neutral_btn"], 
                hover_color=THEME["neutral_hover"], height=24, corner_radius=6, command=lambda: self._abrir_modal_login(obrigatorio=False)
            ).pack(anchor="e", pady=(0, 6))

        tabs = ctk.CTkTabview(
            container, fg_color=THEME["bg_card"], segmented_button_fg_color=THEME["input_bg"],
            segmented_button_selected_color=THEME["primary"], segmented_button_selected_hover_color=THEME["primary_hover"],
            segmented_button_unselected_color=THEME["input_bg"], text_color=THEME["text_main"],
            corner_radius=12, border_width=1, border_color=THEME["border"]
        )
        tabs.pack(fill="both", expand=True, pady=(0, 8))
        
        self.tab_pdf = tabs.add("Extratos PDF")
        self._construir_aba_pdf(self.tab_pdf)

        self.tab_ofx = tabs.add("Extratos OFX")
        self._construir_aba_ofx(self.tab_ofx)

        self.tab_folha = tabs.add("Folha de Pagamento")
        self._construir_aba_folha(self.tab_folha)

        status_bar = ctk.CTkFrame(container, fg_color=THEME["bg_card"], height=32, corner_radius=8, border_width=1, border_color=THEME["border"])
        status_bar.pack(fill="x", side="bottom")

        self.dot_status = ctk.CTkLabel(status_bar, text="●", font=("Segoe UI", 12), text_color=THEME["success"])
        self.dot_status.pack(side="left", padx=(10, 4))

        self.status_label = ctk.CTkLabel(status_bar, text="Pronto para iniciar.", font=FONTS["status"], text_color=THEME["text_muted"])
        self.status_label.pack(side="left")

    def _construir_aba_pdf(self, parent):
        card = ctk.CTkFrame(parent, fg_color="transparent")
        card.pack(fill="both", expand=True, padx=20, pady=15)

        ctk.CTkLabel(card, text="Conversor de Extratos Bancários (PDF)", font=FONTS["section"], text_color=THEME["text_main"]).pack(anchor="w", pady=(0, 2))
        ctk.CTkLabel(card, text="Selecione os formatos de exportação desejados e converta seus extratos:", font=FONTS["body"], text_color=THEME["text_muted"]).pack(anchor="w", pady=(0, 15))

        card_formatos = ctk.CTkFrame(card, fg_color=THEME["input_bg"], corner_radius=10, border_width=1, border_color=THEME["border"])
        card_formatos.pack(fill="x", pady=(0, 20), padx=2)

        ctk.CTkLabel(card_formatos, text="Formatos de Saída:", font=FONTS["body_bold"], text_color=THEME["text_main"]).pack(anchor="w", padx=15, pady=(10, 5))

        formatos_grid = ctk.CTkFrame(card_formatos, fg_color="transparent")
        formatos_grid.pack(fill="x", padx=15, pady=(0, 12))

        self.chk_excel_var = ctk.BooleanVar(value=True)
        self.chk_alterdata_var = ctk.BooleanVar(value=True)

        self.chk_excel = ctk.CTkCheckBox(
            formatos_grid, text="📊 Planilha Excel (.xlsx)", variable=self.chk_excel_var,
            font=FONTS["body_bold"], text_color=THEME["text_main"], fg_color=THEME["primary"]
        )
        self.chk_excel.pack(anchor="w", pady=4)

        self.chk_alterdata = ctk.CTkCheckBox(
            formatos_grid, text="📑 TXT Sistema Contábil (Alterdata/Domínio)", variable=self.chk_alterdata_var,
            font=FONTS["body_bold"], text_color=THEME["text_main"], fg_color=THEME["primary"]
        )
        self.chk_alterdata.pack(anchor="w", pady=4)

        btn_iniciar_pdf = ModernButton(
            master=card, text="📄  Selecionar Extrato(s) PDF e Converter",
            command=self.iniciar_fluxo_conversao_pdf,
            fg_color=THEME["primary"], hover_color=THEME["primary_hover"],
            font=("Segoe UI", 13, "bold"), height=55, corner_radius=10
        )
        btn_iniciar_pdf.pack(fill="x", pady=(5, 5))

    def _construir_aba_ofx(self, parent):
        card = ctk.CTkFrame(parent, fg_color="transparent")
        card.pack(fill="both", expand=True, padx=20, pady=20)

        ctk.CTkLabel(card, text="Conversor de Arquivos OFX", font=FONTS["section"], text_color=THEME["text_main"]).pack(anchor="w", pady=(0, 2))
        ctk.CTkLabel(card, text="Converta arquivos de extrato financeiro padrão OFX em planilhas Excel e TXTs contábeis.", font=FONTS["body"], text_color=THEME["text_muted"]).pack(anchor="w", pady=(0, 20))

        btn_ofx = ModernButton(
            master=card, text="📊  Selecionar Arquivo(s) OFX e Converter",
            command=self.iniciar_fluxo_conversao_ofx,
            fg_color=THEME["primary"], hover_color=THEME["primary_hover"],
            font=("Segoe UI", 13, "bold"), height=58, corner_radius=10
        )
        btn_ofx.pack(fill="x", pady=10)

    def _construir_aba_folha(self, parent):
        card = ctk.CTkFrame(parent, fg_color="transparent")
        card.pack(fill="both", expand=True, padx=20, pady=15)

        ctk.CTkLabel(card, text="Integração de Folha de Pagamento", font=FONTS["section"], text_color=THEME["text_main"]).pack(anchor="w", pady=(0, 2))
        ctk.CTkLabel(card, text="CONVERSÃO EM FASE DE TESTES, SOMENTE PDF DO SISTEMAS DOMÍNIO", font=("Consolas", 10, "italic"), text_color=THEME["warning"]).pack(anchor="w", pady=(0, 15))

        box_data = ctk.CTkFrame(card, fg_color=THEME["input_bg"], corner_radius=8, border_width=1, border_color=THEME["border"])
        box_data.pack(fill="x", pady=(0, 15))

        ctk.CTkLabel(box_data, text="Data do Lançamento:", font=FONTS["body_bold"], text_color=THEME["text_main"]).pack(side="left", padx=12, pady=10)
        self.entry_data_folha = ctk.CTkEntry(
            box_data, width=120, height=32, fg_color=THEME["bg_card"], text_color=THEME["text_main"],
            border_color=THEME["border"], border_width=1, corner_radius=6, placeholder_text="DD/MM/AAAA"
        )
        self.entry_data_folha.insert(0, dt_cls.now().strftime("%d/%m/%Y"))
        self.entry_data_folha.pack(side="right", padx=12, pady=10)
        self.entry_data_folha.bind("<KeyRelease>", self._formatar_mascara_data)

        btn_converter_folha = ModernButton(
            master=card, text="📄  Selecionar PDF da Folha e Converter",
            fg_color=THEME["primary"], hover_color=THEME["primary_hover"],
            font=("Segoe UI", 12, "bold"), height=48, corner_radius=8,
            command=self.processar_conversao_folha
        )
        btn_converter_folha.pack(fill="x", pady=(0, 10))

        btn_rubricas = ModernButton(
            master=card, text="⚙️  Gerenciar Rubricas da Folha",
            fg_color=THEME["neutral_btn"], hover_color=THEME["neutral_hover"],
            font=FONTS["body_bold"], height=40, corner_radius=8,
            command=lambda: JanelaGerenciarRubricasModal(self)
        )
        btn_rubricas.pack(fill="x")

    def _formatar_mascara_data(self, event=None):
        if event and event.keysym in ("BackSpace", "Delete", "Left", "Right", "Up", "Down", "Tab", "Shift_L", "Shift_R", "Control_L", "Control_R"):
            return

        texto = self.entry_data_folha.get()
        numeros = re.sub(r"\D", "", texto)[:8]

        formatado = ""
        if len(numeros) > 4:
            formatado = f"{numeros[:2]}/{numeros[2:4]}/{numeros[4:]}"
        elif len(numeros) == 4:
            formatado = f"{numeros[:2]}/{numeros[2:]}/"
        elif len(numeros) > 2:
            formatado = f"{numeros[:2]}/{numeros[2:]}"
        elif len(numeros) == 2:
            formatado = f"{numeros[:2]}/"
        else:
            formatado = numeros

        if texto != formatado:
            self.entry_data_folha.delete(0, "end")
            self.entry_data_folha.insert(0, formatado)

    def processar_conversao_folha(self):
        data_lanc = self.entry_data_folha.get().strip()
        if not data_lanc:
            messagebox.showwarning("Atenção", "Preencha a data do lançamento.", parent=self)
            return

        caminho_pdf = filedialog.askopenfilename(title="Selecione o PDF da Folha", filetypes=[("Arquivos PDF", "*.pdf")])
        if not caminho_pdf:
            return

        self._set_buttons_state("disabled")
        self.update_status("Processando folha de pagamento... Aguarde.", THEME["warning"])
        
        def task():
            try:
                rubricas_db = carregar_rubricas()
                pasta = os.path.dirname(caminho_pdf)
                nome_base = os.path.splitext(os.path.basename(caminho_pdf))[0]
                caminho_txt = os.path.join(pasta, f"{nome_base}.txt")

                resultado = extrair_dados_folha(caminho_pdf, data_lanc, rubricas_db)
                if isinstance(resultado[1], str) and resultado[0] is None:
                    self.after(0, lambda: messagebox.showerror("Erro", f"Falha ao ler o PDF da Folha:\n{resultado[1]}"))
                    self.after(0, lambda: self.update_status("Erro de leitura.", THEME["danger"]))
                    return

                lancamentos, ignoradas_dict = resultado
                ignoradas_unicas = {
                    "proventos": list(dict.fromkeys(ignoradas_dict["proventos"])),
                    "descontos": list(dict.fromkeys(ignoradas_dict["descontos"]))
                }

                tem_ignoradas = bool(ignoradas_unicas["proventos"] or ignoradas_unicas["descontos"])

                if tem_ignoradas:
                    msg_pendente = (
                        "⚠️ AÇÃO NECESSÁRIA: Existem rubricas que não estão cadastradas.\n\n"
                        "Cadastre-as abaixo ou clique em 'Gerar Arquivo' para forçar a criação do arquivo ignorando os itens pendentes."
                    )
                    
                    def reprocessar_e_gerar():
                        db_atualizado = carregar_rubricas()
                        novos_lanc, _ = extrair_dados_folha(caminho_pdf, data_lanc, db_atualizado)
                        try:
                            gerar_txt_folha(novos_lanc, caminho_txt)
                            self.after(0, lambda: messagebox.showinfo("Sucesso", f"Arquivo TXT gerado com sucesso!\nSalvo em:\n{caminho_txt}"))
                            self.after(0, lambda: self.update_status("Folha gerada com sucesso!", THEME["success"]))
                        except Exception as e:
                            self.after(0, lambda: messagebox.showerror("Erro", f"Erro ao gerar TXT:\n{e}"))
                            self.after(0, lambda: self.update_status("Erro ao salvar TXT.", THEME["danger"]))
                            
                    self.after(0, lambda: JanelaResultadoFolhaModal(self, "Ação Necessária - Pendências", msg_pendente, ignoradas_unicas, reprocessar_e_gerar))
                    self.after(0, lambda: self.update_status("Aguardando ação do usuário.", THEME["warning"]))
                    return

                try:
                    gerar_txt_folha(lancamentos, caminho_txt)
                    msg_sucesso = (
                        f"✅ Arquivo TXT da Folha gerado com sucesso!\n\n"
                        f"Salvo em:\n{caminho_txt}\n\n"
                        f"Total de lançamentos gerados: {len(lancamentos)}\n"
                    )
                    self.after(0, lambda: JanelaResultadoFolhaModal(self, "Sucesso - Folha Gerada", msg_sucesso, ignoradas_unicas))
                    self.after(0, lambda: self.update_status("Folha gerada com sucesso!", THEME["success"]))
                except Exception as e:
                    self.after(0, lambda: messagebox.showerror("Erro", f"Erro ao salvar arquivo TXT:\n{e}"))
                    self.after(0, lambda: self.update_status("Erro ao salvar TXT.", THEME["danger"]))
            finally:
                self.after(0, lambda: self._set_buttons_state("normal"))

        threading.Thread(target=task, daemon=True).start()

    def _carregar_config_ofx(self):
        config_padrao = {
            "ultima_conta_banco": "", 
            "ultima_conta_caixa": "",
            "ultima_conta_aplicacao": "",
            "ultima_conta_rendimento": "",
            "ultima_conta_tarifa": "",
            "ultima_conta_cartao": "",
            "palavras_cartao": ""
        }
        if os.path.exists(ARQUIVO_CONFIG_OFX):
            try:
                with open(ARQUIVO_CONFIG_OFX, 'r', encoding='utf-8') as f:
                    dados_salvos = json.load(f)
                    config_padrao.update(dados_salvos)
            except Exception:
                pass
        return config_padrao
    
    def _salvar_config_ofx(self, banco, caixa, aplicacao="", rendimento="", tarifa="", cartao="", palavras_cartao=""):
        try:
            desocultar_arquivo_windows(ARQUIVO_CONFIG_OFX)
            with open(ARQUIVO_CONFIG_OFX, 'w', encoding='utf-8') as f:
                json.dump({
                    "ultima_conta_banco": banco, 
                    "ultima_conta_caixa": caixa,
                    "ultima_conta_aplicacao": aplicacao,
                    "ultima_conta_rendimento": rendimento,
                    "ultima_conta_tarifa": tarifa,
                    "ultima_conta_cartao": cartao,
                    "palavras_cartao": palavras_cartao
                }, f)
            ocultar_arquivo_windows(ARQUIVO_CONFIG_OFX)
        except Exception:
            pass

    def _obter_contas_gui_multiplos(self, lista_arquivos):
        janela = ctk.CTkToplevel(self)
        janela.title("Definir Contas (Múltiplos Arquivos)")
        janela.geometry("680x700") 
        janela.resizable(True, True)
        janela.configure(fg_color=THEME["bg_root"])
        janela.transient(self)
        janela.attributes("-topmost", True)

        config_salva = self._carregar_config_ofx()
        
        header = ctk.CTkFrame(janela, fg_color="transparent")
        header.pack(fill="x", padx=20, pady=(15, 10))
        ctk.CTkLabel(header, text="Defina as contas para cada arquivo TXT:", font=FONTS["brand"], text_color=THEME["text_main"]).pack(anchor="w")

        scroll_frame = ctk.CTkScrollableFrame(janela, fg_color=THEME["bg_card"], corner_radius=10, border_width=1, border_color=THEME["border"])
        scroll_frame.pack(fill="both", expand=True, padx=20, pady=5)

        # Escolha do Sistema Global
        ctk.CTkLabel(scroll_frame, text="Sistema Contábil Destino (Para todos):", font=FONTS["body_bold"], text_color=THEME["primary"]).pack(anchor="w", pady=(10, 2), padx=10)
        combo_sistema_global = ctk.CTkComboBox(
            scroll_frame, values=["Alterdata", "Domínio Sistemas"], 
            font=FONTS["body_bold"], fg_color=THEME["input_bg"], border_color=THEME["border"]
        )
        combo_sistema_global.pack(fill="x", padx=10, pady=(0, 15))

        entries = {}
        for i, nome_arquivo in enumerate(lista_arquivos):
            file_frame = ctk.CTkFrame(scroll_frame, fg_color=THEME["input_bg"], corner_radius=8, border_width=1, border_color=THEME["border"])
            file_frame.pack(fill="x", padx=5, pady=(8, 0))

            ctk.CTkLabel(file_frame, text=f'📄 Arquivo: "{nome_arquivo}"', font=FONTS["body_bold"], text_color=THEME["text_main"]).pack(pady=(8, 6), padx=12, anchor="w")

            frame_emp = ctk.CTkFrame(file_frame, fg_color="transparent")
            frame_emp.pack(fill="x", padx=12, pady=(0, 5))
            ctk.CTkLabel(frame_emp, text='Código Empresa (P/ Domínio):', text_color=THEME["warning"], font=FONTS["body_bold"], width=200, anchor="w").pack(side="left", padx=(0, 10))
            entry_emp = ctk.CTkEntry(frame_emp, height=32, fg_color=THEME["bg_card"], border_color=THEME["border"])
            entry_emp.pack(side="left", fill="x", expand=True)

            frame_banco = ctk.CTkFrame(file_frame, fg_color="transparent")
            frame_banco.pack(fill="x", padx=12, pady=(0, 5))
            ctk.CTkLabel(frame_banco, text='Conta de Bancos (Movimento):', text_color=THEME["text_main"], font=FONTS["body"], width=200, anchor="w").pack(side="left", padx=(0, 10))
            entry_banco = ctk.CTkEntry(frame_banco, height=32, fg_color=THEME["bg_card"], border_color=THEME["border"])
            entry_banco.insert(0, config_salva.get("ultima_conta_banco", ""))
            entry_banco.pack(side="left", fill="x", expand=True)

            frame_caixa = ctk.CTkFrame(file_frame, fg_color="transparent")
            frame_caixa.pack(fill="x", padx=12, pady=(0, 5))
            ctk.CTkLabel(frame_caixa, text='Conta de Caixa (Contrapartida):', text_color=THEME["text_main"], font=FONTS["body"], width=200, anchor="w").pack(side="left", padx=(0, 10))
            entry_caixa = ctk.CTkEntry(frame_caixa, height=32, fg_color=THEME["bg_card"], border_color=THEME["border"])
            entry_caixa.insert(0, config_salva.get("ultima_conta_caixa", ""))
            entry_caixa.pack(side="left", fill="x", expand=True)
            
            frame_aplicacao = ctk.CTkFrame(file_frame, fg_color="transparent")
            frame_aplicacao.pack(fill="x", padx=12, pady=(0, 5))
            ctk.CTkLabel(frame_aplicacao, text='Conta Aplicação/Resgate:', text_color=THEME["text_main"], font=FONTS["body"], width=200, anchor="w").pack(side="left", padx=(0, 10))
            entry_aplicacao = ctk.CTkEntry(frame_aplicacao, height=32, fg_color=THEME["bg_card"], border_color=THEME["border"])
            entry_aplicacao.insert(0, config_salva.get("ultima_conta_aplicacao", ""))
            entry_aplicacao.pack(side="left", fill="x", expand=True)
            
            frame_rendimento = ctk.CTkFrame(file_frame, fg_color="transparent")
            frame_rendimento.pack(fill="x", padx=12, pady=(0, 5))
            ctk.CTkLabel(frame_rendimento, text='Conta Rendimento:', text_color=THEME["text_main"], font=FONTS["body"], width=200, anchor="w").pack(side="left", padx=(0, 10))
            entry_rendimento = ctk.CTkEntry(frame_rendimento, height=32, fg_color=THEME["bg_card"], border_color=THEME["border"])
            entry_rendimento.insert(0, config_salva.get("ultima_conta_rendimento", ""))
            entry_rendimento.pack(side="left", fill="x", expand=True)
            
            frame_tarifa = ctk.CTkFrame(file_frame, fg_color="transparent")
            frame_tarifa.pack(fill="x", padx=12, pady=(0, 5))
            ctk.CTkLabel(frame_tarifa, text='Conta Tarifas:', text_color=THEME["text_main"], font=FONTS["body"], width=200, anchor="w").pack(side="left", padx=(0, 10))
            entry_tarifa = ctk.CTkEntry(frame_tarifa, height=32, fg_color=THEME["bg_card"], border_color=THEME["border"])
            entry_tarifa.insert(0, config_salva.get("ultima_conta_tarifa", ""))
            entry_tarifa.pack(side="left", fill="x", expand=True)

            frame_cartao = ctk.CTkFrame(file_frame, fg_color="transparent")
            frame_cartao.pack(fill="x", padx=12, pady=(0, 15))
            ctk.CTkLabel(frame_cartao, text='Conta Cartão:', text_color=THEME["text_main"], font=FONTS["body"], width=200, anchor="w").pack(side="left", padx=(0, 10))
            
            entry_cartao = ctk.CTkEntry(frame_cartao, height=32, fg_color=THEME["bg_card"], border_color=THEME["border"])
            entry_cartao.insert(0, config_salva.get("ultima_conta_cartao", ""))
            entry_cartao.pack(side="left", fill="x", expand=True)

            var_palavras = tk.StringVar(value=config_salva.get("palavras_cartao", ""))
            
            def criar_comando_editor(v_var, btn_ref):
                def cmd():
                    win_termos = ctk.CTkToplevel(janela)
                    win_termos.title("Editor de Termos - Cartão")
                    win_termos.geometry("450x260")
                    win_termos.configure(fg_color=THEME["bg_root"])
                    win_termos.transient(janela)
                    win_termos.attributes("-topmost", True)
                    win_termos.grab_set()

                    ctk.CTkLabel(win_termos, text="Adicione ou edite os termos (separados por vírgula):", font=FONTS["body_bold"], text_color=THEME["text_main"]).pack(pady=(20, 5), padx=20, anchor="w")
                    
                    txt_termos = ctk.CTkTextbox(win_termos, height=80, fg_color=THEME["input_bg"], border_color=THEME["border"], border_width=1)
                    txt_termos.pack(fill="x", padx=20, pady=5)
                    txt_termos.insert("1.0", v_var.get())
                    
                    def confirmar():
                        novos = txt_termos.get("1.0", "end-1c").strip()
                        v_var.set(novos)
                        if novos:
                            btn_ref.configure(text="✅ Termos Add", fg_color=THEME["success"])
                        else:
                            btn_ref.configure(text="➕ Termos Extras", fg_color=THEME["primary"])
                        win_termos.grab_release()
                        win_termos.destroy()

                    ctk.CTkButton(win_termos, text="Salvar Termos", command=confirmar, fg_color=THEME["success"], hover_color=THEME["success_hover"], font=FONTS["body_bold"], height=36, corner_radius=8).pack(pady=(15, 0), padx=20, fill="x")
                return cmd

            btn_extras = ctk.CTkButton(frame_cartao, text="✅ Termos Add" if var_palavras.get().strip() else "➕ Termos Extras", width=130, height=32, fg_color=THEME["success"] if var_palavras.get().strip() else THEME["primary"])
            btn_extras.configure(command=criar_comando_editor(var_palavras, btn_extras))
            btn_extras.pack(side="right", padx=(8, 0))

            entries[nome_arquivo] = {
                "empresa": entry_emp,
                "banco": entry_banco, 
                "caixa": entry_caixa, 
                "aplicacao": entry_aplicacao, 
                "rendimento": entry_rendimento,
                "tarifa": entry_tarifa,
                "cartao": entry_cartao,
                "palavras_cartao": var_palavras
            }
            if i == 0:
                entry_banco.focus()

        contas_dados = {}

        def on_submit():
            all_filled = True
            sis_escolhido = combo_sistema_global.get()
            
            for nome_arq, entry_map in entries.items():
                valor_empresa = entry_map["empresa"].get().strip()
                valor_banco = entry_map["banco"].get().strip()
                valor_caixa = entry_map["caixa"].get().strip()
                valor_aplicacao = entry_map["aplicacao"].get().strip()
                valor_rendimento = entry_map["rendimento"].get().strip()
                valor_tarifa = entry_map["tarifa"].get().strip()
                valor_cartao = entry_map["cartao"].get().strip()
                valor_palavras = entry_map["palavras_cartao"].get().strip()
                
                if not valor_banco or not valor_caixa:
                    all_filled = False
                    messagebox.showwarning("Atenção", "Preencha todas as contas Banco e Caixa antes de continuar.", parent=janela)
                    break
                    
                if sis_escolhido == "Domínio Sistemas" and not valor_empresa:
                    all_filled = False
                    messagebox.showwarning("Atenção", "Para exportar para o Domínio, você deve preencher o Código da Empresa.", parent=janela)
                    break
                
                contas_dados[nome_arq] = {
                    'sistema_exportacao': sis_escolhido,
                    'codigo_empresa': valor_empresa,
                    'banco': valor_banco, 
                    'caixa': valor_caixa, 
                    'aplicacao': valor_aplicacao, 
                    'rendimento': valor_rendimento,
                    'tarifa': valor_tarifa,
                    'cartao': valor_cartao,
                    'palavras_cartao': valor_palavras
                }

            if all_filled:
                if contas_dados:
                    primeira = next(iter(contas_dados.values()))
                    self._salvar_config_ofx(
                        primeira['banco'], primeira['caixa'], primeira['aplicacao'], 
                        primeira['rendimento'], primeira['tarifa'], primeira['cartao'], primeira['palavras_cartao']
                    )
                janela.grab_release()
                janela.destroy()

        ctk.CTkButton(
            janela, text="Confirmar e Gerar Arquivos", command=on_submit, height=42,
            font=FONTS["body_bold"], fg_color=THEME["success"], hover_color=THEME["success_hover"], corner_radius=8
        ).pack(pady=(10, 15), padx=20, fill="x")

        janela.lift()
        janela.focus_force()
        janela.grab_set()
        self.wait_window(janela)
        return contas_dados if len(contas_dados) == len(lista_arquivos) else None

    def _obter_contas_gui_UNICO(self, nome_arquivo):
        janela = ctk.CTkToplevel(self)
        janela.title("Parametrização Contábil")
        janela.geometry("540x720")
        janela.resizable(False, False)
        janela.configure(fg_color=THEME["bg_root"])
        janela.transient(self)
        janela.grab_set()

        config_salva = self._carregar_config_ofx()

        ctk.CTkLabel(janela, text="Parametrização Contábil", font=FONTS["brand"], text_color=THEME["text_main"]).pack(pady=(20, 5))
        ctk.CTkLabel(janela, text=f'Arquivo: {nome_arquivo}', font=FONTS["body_bold"], text_color=THEME["text_muted"]).pack(pady=(0, 15))

        scroll = ctk.CTkScrollableFrame(janela, fg_color=THEME["bg_card"], border_color=THEME["border"], border_width=1, corner_radius=8)
        scroll.pack(fill="both", expand=True, padx=25, pady=0)

        # 1. SISTEMA DESTINO
        ctk.CTkLabel(scroll, text="Sistema Contábil Destino:", font=FONTS["body_bold"], text_color=THEME["primary"]).pack(anchor="w", pady=(10, 2), padx=10)
        combo_sistema = ctk.CTkComboBox(
            scroll, values=["Alterdata", "Domínio Sistemas"], 
            font=FONTS["body_bold"], fg_color=THEME["input_bg"], border_color=THEME["border"]
        )
        combo_sistema.pack(fill="x", padx=10, pady=(0, 15))

        ctk.CTkLabel(scroll, text="Código da Empresa (Necessário p/ Domínio):", font=FONTS["body_bold"], text_color=THEME["warning"]).pack(anchor="w", pady=(5, 2), padx=10)
        ent_cod_empresa = ctk.CTkEntry(scroll, placeholder_text="Ex: 327", font=FONTS["code"], fg_color=THEME["input_bg"], border_color=THEME["border"])
        ent_cod_empresa.pack(fill="x", padx=10, pady=(0, 15))

        # 2. CONTAS OBRIGATORIAS
        ctk.CTkLabel(scroll, text="Conta de Bancos (Movimento): *", font=FONTS["body_bold"], text_color=THEME["text_dim"]).pack(anchor="w", pady=(5, 2), padx=10)
        entry_banco = ctk.CTkEntry(scroll, font=FONTS["code"], fg_color=THEME["input_bg"], border_color=THEME["border"])
        entry_banco.insert(0, config_salva.get("ultima_conta_banco", ""))
        entry_banco.pack(fill="x", padx=10, pady=(0, 5))

        ctk.CTkLabel(scroll, text="Conta de Caixa/Fornecedor (Contrapartida): *", font=FONTS["body_bold"], text_color=THEME["text_dim"]).pack(anchor="w", pady=(5, 2), padx=10)
        entry_caixa = ctk.CTkEntry(scroll, font=FONTS["code"], fg_color=THEME["input_bg"], border_color=THEME["border"])
        entry_caixa.insert(0, config_salva.get("ultima_conta_caixa", ""))
        entry_caixa.pack(fill="x", padx=10, pady=(0, 15))

        # 3. CONTAS OPCIONAIS
        ctk.CTkLabel(scroll, text="Contas Inteligentes (Opcionais):", font=FONTS["body_bold"], text_color=THEME["text_main"]).pack(anchor="w", pady=(15, 5), padx=10)
        
        ctk.CTkLabel(scroll, text="Conta p/ Aplicação/Resgate CDB:", font=FONTS["body_bold"], text_color=THEME["text_dim"]).pack(anchor="w", pady=(5, 2), padx=10)
        entry_aplicacao = ctk.CTkEntry(scroll, font=FONTS["code"], fg_color=THEME["input_bg"], border_color=THEME["border"])
        entry_aplicacao.insert(0, config_salva.get("ultima_conta_aplicacao", ""))
        entry_aplicacao.pack(fill="x", padx=10, pady=(0, 5))

        ctk.CTkLabel(scroll, text="Conta p/ Rendimento Poupança:", font=FONTS["body_bold"], text_color=THEME["text_dim"]).pack(anchor="w", pady=(5, 2), padx=10)
        entry_rendimento = ctk.CTkEntry(scroll, font=FONTS["code"], fg_color=THEME["input_bg"], border_color=THEME["border"])
        entry_rendimento.insert(0, config_salva.get("ultima_conta_rendimento", ""))
        entry_rendimento.pack(fill="x", padx=10, pady=(0, 5))

        ctk.CTkLabel(scroll, text="Conta p/ Tarifas/Cesta:", font=FONTS["body_bold"], text_color=THEME["text_dim"]).pack(anchor="w", pady=(5, 2), padx=10)
        entry_tarifa = ctk.CTkEntry(scroll, font=FONTS["code"], fg_color=THEME["input_bg"], border_color=THEME["border"])
        entry_tarifa.insert(0, config_salva.get("ultima_conta_tarifa", ""))
        entry_tarifa.pack(fill="x", padx=10, pady=(0, 5))

        ctk.CTkLabel(scroll, text="Conta p/ Recebimento de Cartões:", font=FONTS["body_bold"], text_color=THEME["text_main"]).pack(anchor="w", pady=(5, 2), padx=10)
        cartao_frame = ctk.CTkFrame(scroll, fg_color="transparent")
        cartao_frame.pack(fill="x", padx=10, pady=(0, 5))
        
        entry_cartao = ctk.CTkEntry(cartao_frame, font=FONTS["code"], fg_color=THEME["input_bg"], border_color=THEME["border"])
        entry_cartao.insert(0, config_salva.get("ultima_conta_cartao", ""))
        entry_cartao.pack(side="left", fill="x", expand=True, padx=(0, 5))

        palavras_cartao_var = tk.StringVar(value=config_salva.get("palavras_cartao", ""))

        def abrir_editor_termos():
            win_termos = ctk.CTkToplevel(janela)
            win_termos.title("Editor de Termos - Cartão")
            win_termos.geometry("450x260")
            win_termos.configure(fg_color=THEME["bg_root"])
            win_termos.transient(janela)
            win_termos.attributes("-topmost", True)
            win_termos.grab_set()

            ctk.CTkLabel(win_termos, text="Adicione ou edite os termos (separados por vírgula):", font=FONTS["body_bold"], text_color=THEME["text_main"]).pack(pady=(20, 5), padx=20, anchor="w")
            txt_termos = ctk.CTkTextbox(win_termos, height=80, fg_color=THEME["input_bg"], border_color=THEME["border"], border_width=1)
            txt_termos.pack(fill="x", padx=20, pady=5)
            txt_termos.insert("1.0", palavras_cartao_var.get())

            def confirmar():
                novos = txt_termos.get("1.0", "end-1c").strip()
                palavras_cartao_var.set(novos)
                if novos:
                    btn_extras.configure(text="✅ Termos Add", fg_color=THEME["success"])
                else:
                    btn_extras.configure(text="➕ Termos Extras", fg_color=THEME["neutral_btn"])
                win_termos.grab_release()
                win_termos.destroy()

            ctk.CTkButton(win_termos, text="Salvar", command=confirmar, fg_color=THEME["primary"], font=FONTS["body_bold"]).pack(pady=15)

        btn_extras = ctk.CTkButton(cartao_frame, text="✅ Termos Add" if palavras_cartao_var.get().strip() else "➕ Termos Extras", width=80, fg_color=THEME["success"] if palavras_cartao_var.get().strip() else THEME["neutral_btn"], hover_color=THEME["neutral_hover"], font=FONTS["body_bold"], command=abrir_editor_termos)
        btn_extras.pack(side="right")

        contas_dados = {}

        def on_submit():
            b = entry_banco.get().strip()
            c = entry_caixa.get().strip()
            sistema = combo_sistema.get()
            empresa = ent_cod_empresa.get().strip()

            if not b or not c:
                messagebox.showwarning("Atenção", "Preencha a conta de Banco e a Contrapartida (Caixa).", parent=janela)
                return
                
            if sistema == "Domínio Sistemas" and not empresa:
                messagebox.showwarning("Atenção", "Para exportar para o Domínio, você deve preencher o Código da Empresa.", parent=janela)
                return

            self._salvar_config_ofx(b, c, entry_aplicacao.get().strip(), entry_rendimento.get().strip(), entry_tarifa.get().strip(), entry_cartao.get().strip(), palavras_cartao_var.get().strip())

            contas_dados['sistema_exportacao'] = sistema
            contas_dados['codigo_empresa'] = empresa
            contas_dados['banco'] = b
            contas_dados['caixa'] = c
            contas_dados['aplicacao'] = entry_aplicacao.get().strip()
            contas_dados['rendimento'] = entry_rendimento.get().strip()
            contas_dados['tarifa'] = entry_tarifa.get().strip()
            contas_dados['cartao'] = entry_cartao.get().strip()
            contas_dados['palavras_cartao'] = palavras_cartao_var.get().strip()
            
            janela.grab_release()
            janela.destroy()

        ctk.CTkButton(
            janela, text="Confirmar e Gerar Arquivo", fg_color=THEME["success"], hover_color=THEME["success_hover"], 
            font=FONTS["body_bold"], height=42, corner_radius=8, command=on_submit
        ).pack(fill="x", padx=25, pady=20)

        janela.lift()
        janela.focus_force()
        janela.grab_set()
        self.wait_window(janela)
        return contas_dados if 'banco' in contas_dados else None

    def atualizar_status_licenca_ui(self):
        if not self.login_obrigatorio:
            self.pill_licenca.pack_forget()
            return

        self.pill_licenca.pack(side="right", pady=4)
        valido, tempo_str, cliente, msg = validar_e_obter_dados_licenca(ARQUIVO_LICENCA)

        if valido:
            self.lbl_licenca.configure(
                text=f"🟢 {cliente} ({tempo_str})",
                text_color=THEME["success"]
            )
        elif valido is False:
            self.lbl_licenca.configure(text=f"⚠️ {msg}", text_color=THEME["danger"])
        else:
            self.lbl_licenca.configure(text="⚠️ Não Autenticado", text_color=THEME["warning"])

    def _abrir_modal_login(self, obrigatorio=False):
        modal = JanelaLoginDialog(master=self, obrigatorio=obrigatorio)
        self.wait_window(modal)
        self.atualizar_status_licenca_ui()

    def show_up_to_date_status(self):
        self.update_status("Você já está na versão mais recente.", THEME["success"])
        self.after(5000, lambda: self.update_status("Pronto para iniciar."))

    def _executar_funcao_dinamica(self, func, path):
        try:
            sig = inspect.signature(func)
            if len(sig.parameters) > 0:
                return func(path)
            return func()
        except Exception:
            try:
                return func(path)
            except TypeError:
                return func()

    def _escolher_modelo(self, model_config):
        dialog = JanelaSelecaoModeloModal(self, model_config)
        self.wait_window(dialog)
        return dialog.modelo_escolhido

    def iniciar_fluxo_conversao_pdf(self):
        caminhos_raw = filedialog.askopenfilenames(title="Selecione os arquivos PDF", filetypes=[("Arquivos PDF", "*.pdf")])
        if not caminhos_raw: 
            return
            
        if isinstance(caminhos_raw, str):
            try:
                caminhos_pdf = list(self.tk.splitlist(caminhos_raw))
            except Exception:
                caminhos_pdf = [caminhos_raw]
        else:
            caminhos_pdf = list(caminhos_raw)
            
        caminhos_pdf = [os.path.normpath(str(p).strip().strip('{}')) for p in caminhos_pdf if str(p).strip()]
        if not caminhos_pdf:
            return

        banco_key = self._escolher_banco_dialog()
        if not banco_key:
            return

        modelo_escolhido = None
        cfg_banco = CONVERTERS.get(banco_key, {})
        if cfg_banco.get("type") == "model_choice":
            modelo_escolhido = self._escolher_modelo(cfg_banco["model_config"])
            if not modelo_escolhido:
                return

        self.processar_conversao(banco_key, caminho_pdf=caminhos_pdf, modelo_escolhido=modelo_escolhido)

    def iniciar_fluxo_conversao_ofx(self):
        caminhos_raw = filedialog.askopenfilenames(
            title="Selecione os arquivos OFX",
            filetypes=[("Arquivos OFX/QFX", "*.ofx;*.qfx;*.OFX;*.QFX"), ("Todos os Arquivos", "*.*")]
        )
        if not caminhos_raw:
            return

        if isinstance(caminhos_raw, str):
            try:
                caminhos_ofx = list(self.tk.splitlist(caminhos_raw))
            except Exception:
                caminhos_ofx = [caminhos_raw]
        else:
            caminhos_ofx = list(caminhos_raw)

        caminhos_ofx = [os.path.normpath(str(p).strip().strip('{}')) for p in caminhos_ofx if str(p).strip()]
        if not caminhos_ofx:
            return

        self.processar_conversao("ofx", caminho_pdf=caminhos_ofx)

    def _escolher_banco_dialog(self):
        banco_sel = ctk.StringVar(value="")
        janela = ctk.CTkToplevel(self)
        janela.title("Selecionar Instituição Bancária")
        janela.geometry("680x520")
        janela.configure(fg_color=THEME["bg_root"])
        janela.resizable(False, False)

        janela.transient(self)
        janela.attributes("-topmost", True)

        header = ctk.CTkFrame(janela, fg_color="transparent")
        header.pack(fill="x", padx=25, pady=(18, 8))
        ctk.CTkLabel(header, text="Selecione o Banco Emissor:", font=FONTS["brand"], text_color=THEME["text_main"]).pack(anchor="w")
        ctk.CTkLabel(header, text="Clique no banco desejado para prosseguir:", font=FONTS["body"], text_color=THEME["text_muted"]).pack(anchor="w")

        scroll_frame = ctk.CTkScrollableFrame(
            janela, fg_color=THEME["bg_card"], corner_radius=10, border_width=1, border_color=THEME["border"]
        )
        scroll_frame.pack(fill="both", expand=True, padx=25, pady=(0, 20))

        for key, config in sorted(CONVERTERS.items(), key=lambda i: i[1]['nome']):
            if config.get('aba') == 'pdf':
                card = ctk.CTkFrame(
                    scroll_frame, fg_color=THEME["neutral_btn"], corner_radius=8,
                    border_width=1, border_color=THEME["border"], height=48
                )
                card.pack(fill="x", padx=5, pady=4)

                is_multiplo = config.get("type") == "model_choice"

                btn_banco = ctk.CTkButton(
                    card, text=f"  {config['nome']}", image=self.icons.get(key), anchor="w",
                    fg_color="transparent", hover_color=THEME["neutral_hover"], font=FONTS["body_bold"],
                    text_color=THEME["text_main"], height=44, corner_radius=8, cursor="hand2",
                    command=lambda k=key: (banco_sel.set(k), janela.destroy())
                )
                btn_banco.pack(side="left", fill="both", expand=True, padx=(4, 0), pady=2)

                if not is_multiplo:
                    caminho_img = resolver_caminho_preview(key) or resolver_caminho_preview(config.get("preview"))
                    if caminho_img:
                        btn_preview = ctk.CTkButton(
                            card, text="👁️ Exemplo", width=88, height=30,
                            font=("Segoe UI", 10, "bold"), fg_color="#181B22", hover_color="#2D3442",
                            border_width=1, border_color="#334155", corner_radius=6, cursor="hand2"
                        )
                        btn_preview.pack(side="right", padx=(0, 8), pady=7)
                        HoverAndClickPreview(btn_preview, caminho_img, titulo=config['nome'])

        janela.lift()
        janela.focus_force()
        janela.grab_set()
        self.wait_window(janela)
        return banco_sel.get()

    def _set_buttons_state(self, new_state: str):
        for tab in [self.tab_pdf, self.tab_ofx]:
            for w in tab.winfo_children():
                if isinstance(w, ModernButton) or isinstance(w, ctk.CTkCheckBox):
                    w.configure(state=new_state)

    def update_status(self, message, color=None):
        self.status_label.configure(text=message)
        self.dot_status.configure(text_color=color or THEME["success"])
        self.update_idletasks()

    def processar_conversao(self, key, caminho_pdf=None, modelo_escolhido=None):
        paths = [caminho_pdf] if isinstance(caminho_pdf, str) else list(caminho_pdf or [])
        if not paths:
            return

        gerar_excel = self.chk_excel_var.get()
        gerar_alt = self.chk_alterdata_var.get()

        contas_map = {}
        if gerar_alt:
            nomes_arquivos = [os.path.basename(p) for p in paths]
            if len(nomes_arquivos) > 1:
                contas_map = self._obter_contas_gui_multiplos(nomes_arquivos)
            else:
                contas_unica = self._obter_contas_gui_UNICO(nomes_arquivos[0])
                if contas_unica:
                    contas_map = {nomes_arquivos[0]: contas_unica}
                else:
                    contas_map = None

            if contas_map is None:
                return

        self._executar_processamento_em_thread(key, paths, gerar_excel, gerar_alt, contas_map, modelo_escolhido=modelo_escolhido)

    def _executar_processamento_em_thread(self, key, paths, gerar_excel, gerar_alt, contas_map, modelo_escolhido=None):
        self._set_buttons_state("disabled")
        self.update_status("Processando arquivos... Aguarde.", THEME["warning"])

        def task():
            try:
                resultados = self.run_converter(key, paths, modelo_escolhido=modelo_escolhido)
                resultados_validos = [r for r in resultados if r and os.path.exists(str(r))]

                if resultados_validos:
                    arquivos_finais = []

                    for caminho_retornado in resultados_validos:
                        try:
                            df_bruto = carregar_dados_extraidos(caminho_retornado)
                            df_extraido = normalizar_dataframe_extrato(df_bruto)

                            pasta_dest = os.path.dirname(caminho_retornado)
                            base_nome = os.path.splitext(os.path.basename(caminho_retornado))[0]

                            for suf in ["_stone", "_bradesco", "_temp", "_Consolidado"]:
                                if base_nome.endswith(suf):
                                    base_nome = base_nome[:-len(suf)]

                            caminho_excel = os.path.join(pasta_dest, f"{base_nome}.xlsx")
                            
                            # Exportação Planilha
                            if gerar_excel:
                                df_extraido.to_excel(caminho_excel, index=False)
                                if os.path.exists(caminho_excel):
                                    arquivos_finais.append(caminho_excel)

                            # Exportação Sistema Contábil
                            if gerar_alt:
                                nome_original_pdf = f"{base_nome}.pdf"
                                config_especifica = {}
                                if contas_map:
                                    config_especifica = (
                                        contas_map.get(nome_original_pdf) or 
                                        contas_map.get(os.path.basename(caminho_retornado)) or 
                                        next(iter(contas_map.values()), {})
                                    )

                                sistema_escolhido = config_especifica.get("sistema_exportacao", "Alterdata")
                                
                                if sistema_escolhido == "Alterdata":
                                    txt_alt = os.path.join(pasta_dest, f"{base_nome}_Alterdata.txt")
                                    gerar_txt_alterdata(df_extraido, txt_alt, config_contas=config_especifica)
                                    if os.path.exists(txt_alt):
                                        arquivos_finais.append(txt_alt)
                                        
                                elif sistema_escolhido == "Domínio Sistemas":
                                    txt_dom = os.path.join(pasta_dest, f"{base_nome}_Dominio.txt")
                                    gerar_txt_dominio(df_extraido, txt_dom, config_contas=config_especifica)
                                    if os.path.exists(txt_dom):
                                        arquivos_finais.append(txt_dom)

                            # Limpeza temp
                            if caminho_retornado.lower().endswith(".csv") and os.path.exists(caminho_retornado):
                                try: os.remove(caminho_retornado)
                                except Exception: pass

                            if not gerar_excel and os.path.exists(caminho_excel) and gerar_alt:
                                try: os.remove(caminho_excel)
                                except Exception: pass

                        except Exception as e_txt:
                            print(f"Erro ao exportar arquivos: {e_txt}")

                    total_gerados = len(arquivos_finais)
                    if total_gerados > 0:
                        msg = f"Conversão concluída com sucesso!\n\nTotal de arquivos gerados: {total_gerados}"
                        self.after(0, lambda: messagebox.showinfo("Sucesso", msg))
                        self.after(0, lambda: self.update_status(f"{total_gerados} arquivo(s) gerado(s).", THEME["success"]))
                    else:
                        self.after(0, lambda: messagebox.showwarning("Aviso", "Nenhum arquivo final pôde ser salvo."))
                        self.after(0, lambda: self.update_status("Nenhum arquivo gerado.", THEME["warning"]))
                else:
                    self.after(0, lambda: messagebox.showwarning("Aviso", "Nenhuma transação foi identificada pelo conversor."))
                    self.after(0, lambda: self.update_status("Nenhum arquivo gerado.", THEME["warning"]))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Erro de Processamento", f"Ocorreu uma falha:\n{e}"))
                self.after(0, lambda: self.update_status("Falha no processamento.", THEME["danger"]))
            finally:
                self.after(0, lambda: self._set_buttons_state("normal"))
                self.after(5000, lambda: self.update_status("Pronto para iniciar."))

        threading.Thread(target=task, daemon=True).start()

    def run_converter(self, key, caminho_pdf=None, modelo_escolhido=None):
        config = CONVERTERS[key]
        paths = [caminho_pdf] if isinstance(caminho_pdf, str) else list(caminho_pdf or [])
        res = []

        if key == "ofx" or config.get("type") == "ofx":
            for path in paths:
                try:
                    res.append(extrair_extrato_ofx(path))
                except Exception as e:
                    print(f"Erro ao converter OFX '{path}': {e}")
            return res

        for path in paths:
            caminho_saida = None
            try:
                with self.patch_file_dialogs(path):
                    if config.get("type") == "simple_run":
                        mod = importlib.import_module(config['module'])
                        func = getattr(mod, config['function'], None) or getattr(mod, 'main', None) or getattr(mod, 'iniciar_processamento', None)
                        caminho_saida = self._executar_funcao_dinamica(func, path)
                        
                    elif config.get("type") == "multi_file":
                        caminho_saida = self._run_multi_file_converter(key, config, paths)
                        if isinstance(caminho_saida, list):
                            res.extend(caminho_saida)
                            break
                        
                    elif config.get("type") == "model_choice":
                        caminho_saida = self._run_model_choice_converter_core(key, config, path, modelo_escolhido)
            except Exception as e_mod:
                print(f"Aviso no conversor específico ({key}): {e_mod}")

            if not caminho_saida or not os.path.exists(str(caminho_saida)) or os.path.getsize(str(caminho_saida)) == 0:
                try:
                    caminho_saida = extrair_pdf_universal(path)
                except Exception as e_fallback:
                    print(f"Falha no extrator universal para '{path}': {e_fallback}")

            if caminho_saida:
                res.append(caminho_saida)
                    
        return res

    def _run_model_choice_converter_core(self, key, config, path, modelo_escolhido):
        mod_name = None

        if key == "itau":
            if modelo_escolhido == "modelo1": mod_name = "conversor_itaumod1"
            elif modelo_escolhido == "modelo2": mod_name = "conversor_itaumod2"
            else: mod_name = "conversor_itaumod3"
        elif key == "bb":
            mod_name = "conversor_bbmod1" if modelo_escolhido == "modelo1" else "conversor_bbmod2"
        elif key == "cef":
            mod_name = "conversor_cefmod1" if modelo_escolhido == "modelo1" else "conversor_cefmod2"
        elif key == "santander":
            mod_name = "conversor_santandermod1" if modelo_escolhido == "modelo1" else "conversor_santandermod2"
        elif key == "sicoob":
            if modelo_escolhido == "modelo1": mod_name = "conversor_sicoobmod1"
            elif modelo_escolhido == "modelo2": mod_name = "conversor_sicoobmod2"
            else: mod_name = "conversor_sicoobmod3"
        elif key == "safra":
            mod_name = "conversor_saframod1" if modelo_escolhido == "modelo1" else "conversor_saframod2"

        if mod_name:
            try:
                mod = importlib.import_module(mod_name)
                func = getattr(mod, "iniciar_processamento", None) or getattr(mod, "main", None)
                if not func:
                    raise AttributeError(f"O módulo '{mod_name}' não possui a função 'iniciar_processamento' nem 'main'.")

                with self.patch_file_dialogs(path):
                    return self._executar_funcao_dinamica(func, path)
            except Exception as e_mod_exec:
                print(f"❌ Erro na execução de '{mod_name}': {e_mod_exec}")
                raise e_mod_exec
        return None

    def _run_multi_file_converter(self, key, config, paths):
        mod = importlib.import_module(config['module'])
        dfs = [mod.extrair_texto_pdf(p) for p in paths]
        comb = pd.concat([d for d in dfs if d is not None], ignore_index=True)
        path_s = os.path.splitext(paths[0])[0] + "_Consolidado.xlsx"
        comb.to_excel(path_s, index=False)
        return [path_s]

# ==========================================
# 🚀 PONTO DE ENTRADA
# ==========================================
if __name__ == "__main__":
    app = ConversorApp()
    app.mainloop()