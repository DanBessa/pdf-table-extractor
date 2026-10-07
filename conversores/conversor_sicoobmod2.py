# conversor_sicoobmod2.py (CORRIGIDO)

import pdfplumber
import pandas as pd
from tkinter import filedialog # Mantém apenas o filedialog
import os
import re
# REMOVIDO: import tkinter as tk, messagebox

def extrair_ano_do_pdf(pdf_pages):
    """Extrai o ano da linha 'PERÍODO' na primeira página para construir a data completa."""
    try:
        primeira_pagina_texto = pdf_pages[0].extract_text(x_tolerance=2)
        match = re.search(r"PERÍODO: \d{2}\/\d{2}\/(\d{4})", primeira_pagina_texto)
        if match:
            return match.group(1)
    except Exception:
        pass
    # Retorna o ano atual se não encontrar
    return str(pd.Timestamp.now().year)

def extrair_dados_do_pdf(caminho_pdf):
    """
    Extrai dados de um extrato Sicoob (Modelo 2) e RETORNA um DataFrame.
    """
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            ano = extrair_ano_do_pdf(pdf.pages)
            texto_completo = "\n".join([page.extract_text(x_tolerance=2) or "" for page in pdf.pages])
    except Exception as e:
        # MUDANÇA: Substitui messagebox por raise
        raise Exception(f"Não foi possível ler o arquivo PDF:\n{e}")

    texto_completo = re.sub(r".*HISTÓRICO DE MOVIMENTAÇÃO\n", "", texto_completo, flags=re.DOTALL)
    texto_completo = re.sub(r"SALDO ANTERIOR.*?\n", "", texto_completo, flags=re.DOTALL)
    texto_completo = re.sub(r"\nRESUMO.*", "", texto_completo, flags=re.DOTALL)
    
    blocos = re.split(r'\n(?=\d{2}/\d{2})', texto_completo.strip())
    transacoes = []

    for bloco in blocos:
        texto_bloco = re.sub(r'\s{2,}', ' ', bloco.replace('\n', ' ').strip())
        if "SALDO DO DIA" in texto_bloco:
            continue
        
        match_valor_tipo = re.search(r'(\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2}|\d+\.\d{2})\s*([CD])', texto_bloco)
        data_match = re.match(r'(\d{2}/\d{2})', texto_bloco)
        
        if data_match and match_valor_tipo:
            data = f"{data_match.group(1)}/{ano}" 
            valor_str = match_valor_tipo.group(1)
            tipo = match_valor_tipo.group(2)
            
            descricao = texto_bloco
            descricao = re.sub(r'^\d{2}/\d{2}\s*', '', descricao).strip()
            descricao = descricao.replace(match_valor_tipo.group(0), '', 1).strip()
            descricao = re.sub(r'\s{2,}', ' ', descricao).strip()

            valor_numerico = float(valor_str.replace('.', '').replace(',', '.'))
            if tipo == 'D':
                valor_numerico *= -1

            if descricao:
                transacoes.append([data, descricao, valor_numerico])

    if not transacoes:
        return pd.DataFrame()

    df = pd.DataFrame(transacoes, columns=["Data", "Lancamento", "Valor"])
    return df

def iniciar_processamento():
    """Função chamada pelo programa principal para iniciar a conversão."""
    
    # MUDANÇA: de askopenfilenames para askopenfilename
    arquivo_path = filedialog.askopenfilename(
        title="Selecione o extrato (Sicoob - Modelo 2)",
        filetypes=[("Arquivos PDF", "*.pdf")]
    )
    if not arquivo_path:
        raise UserWarning("Nenhum arquivo foi selecionado.")

    nome_arquivo_original = os.path.basename(arquivo_path)

    try:
        df_transacoes = extrair_dados_do_pdf(arquivo_path)
        
        if df_transacoes is None or df_transacoes.empty:
             raise UserWarning(f"Nenhuma transação encontrada em '{nome_arquivo_original}'.")

        nome_base, _ = os.path.splitext(arquivo_path)
        caminho_csv = nome_base + '.csv'
        
        # Formata a coluna 'Valor' como string com vírgula para o leitor de Decimal
        df_transacoes['Valor'] = df_transacoes['Valor'].apply(lambda x: str(f"{x:.2f}").replace('.', ','))
        
        # Renomeia a coluna para o padrão do gerador de TXT
        df_transacoes.rename(columns={"Lancamento": "Histórico"}, inplace=True)
        
        df_transacoes.to_csv(caminho_csv, index=False, sep=';', encoding='utf-8-sig')
        
        return caminho_csv

    except Exception as e:
        # Relança a exceção para o 'processar_conversao' tratar
        raise e

