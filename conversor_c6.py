# conversor_c6.py (CORRIGIDO)

import os
import re
import pdfplumber
import pandas as pd
from tkinter import filedialog, messagebox
import traceback

def limpar_valor(valor_str):
    """
    Limpa a string de valor e a converte para float, tratando o formato brasileiro
    e sinais negativos separados.
    """
    if not isinstance(valor_str, str):
        return 0.0

    is_negative = '-' in valor_str
    valor_limpo = re.sub(r'[^\d,]', '', valor_str)
    valor_limpo = valor_limpo.replace(',', '.')
    
    try:
        valor_float = float(valor_limpo)
        if is_negative:
            return -abs(valor_float)
        return valor_float
    except (ValueError, TypeError):
        return 0.0

def extrair_dados_do_pdf(pdf_path, senha):
    """
    Extrai dados de transação de um extrato C6 Bank.
    """
    transacoes = []
    
    with pdfplumber.open(pdf_path, password=senha) as pdf:
        ano = None
        texto_completo_para_ano = "".join([p.extract_text() or "" for p in pdf.pages])
        ano_match = re.search(r'Período \d{1,2} de \w+ de (\d{4})', texto_completo_para_ano) or \
                    re.search(r'exportado no dia \d{1,2} de \w+ de (\d{4})', texto_completo_para_ano)
        if ano_match:
            ano = ano_match.group(1)
        else:
            raise ValueError("Não foi possível encontrar o ano no extrato.")

        data_transacao_atual = None

        for page in pdf.pages:
            texto_pagina = page.extract_text(x_tolerance=2)
            if not texto_pagina:
                continue

            linhas = texto_pagina.split('\n')
            
            for linha in linhas:
                linha_limpa = linha.strip()

                if not linha_limpa or "Saldo do dia" in linha_limpa or "Data Lançamento" in linha_limpa:
                    continue
                
                # Tenta encontrar e atualizar a data do contexto de forma mais segura
                data_match = re.match(r'(\d{2}/\d{2})', linha_limpa)
                if data_match:
                    try:
                        # Valida se a data extraída é válida (ex: não contém mês "00")
                        dia, mes = data_match.group(1).split('/')
                        if 1 <= int(mes) <= 12 and 1 <= int(dia) <= 31:
                            data_transacao_atual = f"{data_match.group(1)}/{ano}"
                    except (ValueError, IndexError):
                        continue # Pula se o formato da data estiver quebrado
                
                # Procura por uma transação (descrição + valor no final)
                transacao_match = re.search(r'^(.*?)\s+(-?R\$\s?[\d\.,]+)$', linha_limpa)
                
                if data_transacao_atual and transacao_match:
                    descricao, valor_str = transacao_match.groups()
                    descricao_limpa = descricao.strip()
                    
                    # Remove a data contábil da descrição, se houver
                    descricao_limpa = re.sub(r'^\d{2}/\d{2}\s*', '', descricao_limpa).strip()
                    
                    valor_float = limpar_valor(valor_str)
                    
                    if descricao_limpa and valor_float != 0.0:
                        transacoes.append({
                            "Data": data_transacao_atual,
                            "Lançamento": descricao_limpa,
                            "Valor": valor_float
                        })

    if not transacoes:
        return pd.DataFrame()

    return pd.DataFrame(transacoes).drop_duplicates().reset_index(drop=True)

def iniciar_processamento(pdf_path=None):
    """
    Função de entrada chamada pelo menu principal.
    """
    if not pdf_path:
        pdf_path = filedialog.askopenfilename(
            title="Selecione o extrato do C6 Bank",
            filetypes=[("PDF files", "*.pdf")]
        )
        if not pdf_path:
            return False

    try:
        df = extrair_dados_do_pdf(pdf_path, senha='062237')

        if df.empty:
            messagebox.showwarning("Aviso", "Nenhuma transação válida foi encontrada no arquivo.")
            return False

        base_name = os.path.splitext(pdf_path)[0]
        output_csv_path = base_name + ".csv"
        
        # A validação de data na extração previne o erro nesta linha
        df['Data'] = pd.to_datetime(df['Data'], format='%d/%m/%Y').dt.strftime('%d/%m/%Y')
        df.to_csv(output_csv_path, index=False, sep=';', encoding='utf-8-sig', decimal=',')
        
        return True

    except Exception as e:
        traceback.print_exc()
        messagebox.showerror("Erro Crítico", f"Ocorreu um erro inesperado ao processar o extrato do C6 Bank.\n\n{e}")
        return False