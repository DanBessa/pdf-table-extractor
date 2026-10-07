# conversor_pagbank.py (CORRIGIDO)

import pdfplumber
import pandas as pd
import re
# REMOVIDO: from tkinter import messagebox

def extrair_texto_pdf(pdf_path):
    """
    Extrai transações de um único arquivo PDF e RETORNA um DataFrame.
    (Chamado pelo 'run_multi_file_converter' do app principal)
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])

        pattern_corrected = re.compile(r"(\d{2}/\d{2}/\d{4})\s+(.+?)\s+(-?R?\$\s?[\d\.]+,\d{2})")
        matches = pattern_corrected.findall(text)

        if not matches:
            print(f"Nenhuma transação encontrada no arquivo: {pdf_path}")
            return pd.DataFrame()

        df = pd.DataFrame(matches, columns=["Data", "Descrição", "Valor"])
        
        # Formata o valor para o padrão CSV brasileiro (com vírgula)
        # 1. Remove "R$" e pontos de milhar
        df['Valor'] = df['Valor'].str.replace(r'R\$\s?', '', regex=True).str.replace('.', '', regex=False)
        # 2. Adiciona hífen no início se for " - "
        df['Valor'] = df['Valor'].str.replace(r'\s+-\s+', '-', regex=True).str.strip()
        # 3. Garante que o formato final seja "1234,56" ou "-1234,56"
        
        df.rename(columns={"Descrição": "Histórico"}, inplace=True)
        
        return df

    except Exception as e:
        # Relança a exceção para o app principal
        print(f"Erro ao processar PagBank: {e}")
        raise Exception(f"Não foi possível processar o arquivo:\n{pdf_path}\n\nErro: {e}")
