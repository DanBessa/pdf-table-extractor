# conversor_paycash.py (CORRIGIDO)

import pdfplumber
import pandas as pd
import re
import os
from tkinter import filedialog # Mantém apenas o filedialog
# REMOVIDO: import tkinter as tk, messagebox

def clean_valor(valor_str):
    """Limpa e converte uma string de moeda para um número float."""
    if not valor_str or not isinstance(valor_str, str):
        return 0.0
    valor_limpo = valor_str.replace('R$', '').strip().replace('.', '').replace(',', '.')
    try:
        return float(valor_limpo)
    except (ValueError, TypeError):
        return 0.0

def formatar_para_csv(valor_numerico):
    """Formata um número para o padrão CSV brasileiro (ex: 1234,56)."""
    valor_str = f"{valor_numerico:.2f}"
    valor_str_brl = valor_str.replace('.', ',')
    return valor_str_brl

def extrair_dados_do_pdf(pdf_path):
    """Extrai dados de um único PDF PayCash."""
    transacoes = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            texto_completo = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])

        for linha in texto_completo.split('\n'):
            if re.match(r'^\d{2}/\d{2}/\d{4}', linha):
                partes = linha.split()
                if len(partes) < 2: continue
                
                data = partes[0]
                historico = "Não Identificado"
                valor = 0.0

                valores_monetarios = re.findall(r'R\$\s?[\d.,]+', linha)
                
                if len(valores_monetarios) >= 2:
                    valor_transacao_str = valores_monetarios[-2]
                    
                    if "Emissão de TED" in linha or "Pix Enviado" in linha:
                        historico = "Pix Enviado" if "Pix Enviado" in linha else "Emissão de TED"
                        valor = -clean_valor(valor_transacao_str)
                    elif "Deposito Cofre" in linha:
                        historico = "Deposito Cofre"
                        valor = clean_valor(valor_transacao_str)
                    
                    if valor != 0:
                        transacoes.append({
                            "Data": data,
                            "Histórico": historico,
                            "Valor": valor # Salva como float
                        })

        if not transacoes:
            return pd.DataFrame()

        return pd.DataFrame(transacoes)
    
    except Exception as e:
        raise Exception(f"Falha ao processar o arquivo '{os.path.basename(pdf_path)}'.\n\nDetalhes: {e}")

def iniciar_processamento(pdf_path):
    """
    Função principal chamada pelo menu.
    """
    # Esta chamada será interceptada pelo patch
    # pdf_path = filedialog.askopenfilenames(
    #     title="Selecione o extrato Pay Cash",
    #     filetypes=[("Arquivos PDF", "*.pdf")]
    # )

    if not pdf_path:
        raise UserWarning("Nenhum arquivo selecionado.")

    try:
        df_transacoes = extrair_dados_do_pdf(pdf_path)

        if df_transacoes is None or df_transacoes.empty:
            raise UserWarning(f"Nenhuma transação encontrada no arquivo: {os.path.basename(pdf_path)}")

        df_transacoes['Valor'] = df_transacoes['Valor'].apply(formatar_para_csv)
        
        caminho_csv = os.path.splitext(pdf_path)[0] + ".csv"
        df_transacoes.to_csv(caminho_csv, index=False, sep=";", encoding="utf-8-sig")
        
        return caminho_csv

    except Exception as e:
        raise Exception(f"Falha ao processar o arquivo '{os.path.basename(pdf_path)}'.\n\nDetalhes: {e}")

# REMOVIDO: Lógica de salvamento de Excel