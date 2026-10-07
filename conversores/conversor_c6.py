# conversor_c6.py (CORRIGIDO)

import os
import re
import pdfplumber
import pandas as pd
import traceback
# REMOVIDO: tkinter (não é mais necessário)

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
            ano_fallback = re.search(r'\d{2}/\d{2}/(\d{4})', texto_completo_para_ano)
            if ano_fallback:
                ano = ano_fallback.group(1)
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
                
                data_match = re.match(r'(\d{2}/\d{2})', linha_limpa)
                if data_match:
                    try:
                        dia, mes = data_match.group(1).split('/')
                        if 1 <= int(mes) <= 12 and 1 <= int(dia) <= 31:
                            data_transacao_atual = f"{data_match.group(1)}/{ano}"
                    except (ValueError, IndexError):
                        continue 
                
                transacao_match = re.search(r'^(.*?)\s+(-?R\$\s?[\d\.,]+)$', linha_limpa)
                
                if data_transacao_atual and transacao_match:
                    descricao, valor_str = transacao_match.groups()
                    descricao_limpa = descricao.strip()
                    
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

# --- CORREÇÃO PRINCIPAL ---
def iniciar_processamento(pdf_path): # <--- 1. Aceita 'pdf_path' como argumento
    """
    Função de entrada chamada pelo menu principal.
    Recebe o caminho do PDF diretamente.
    """
    # 2. Remove o filedialog (não é mais necessário)
    if not pdf_path:
        raise UserWarning("Nenhum caminho de PDF foi fornecido ao conversor C6.")

    try:
        # 3. Usa o 'pdf_path' recebido
        df = extrair_dados_do_pdf(pdf_path, senha='062237') 

        if df.empty:
            raise UserWarning("Nenhuma transação válida foi encontrada no arquivo.")

        base_name = os.path.splitext(pdf_path)[0]
        output_csv_path = base_name + ".csv"
        
        df['Data'] = pd.to_datetime(df['Data'], format='%d/%m/%Y').dt.strftime('%d/%m/%Y')
        # Formata o valor para o padrão CSV brasileiro (com vírgula)
        df['Valor'] = df['Valor'].apply(lambda x: str(f"{x:.2f}").replace('.', ','))
        
        df.rename(columns={"Lançamento": "Histórico"}, inplace=True) # Renomeia para o padrão
        df = df[["Data", "Histórico", "Valor"]] # Garante a ordem
        
        df.to_csv(output_csv_path, index=False, sep=';', encoding='utf-8-sig')
        
        # 4. Retorna o caminho do CSV
        return output_csv_path

    except Exception as e:
        traceback.print_exc()
        # Relança a exceção para o app principal mostrar o erro
        raise Exception(f"Erro ao processar extrato C6 Bank:\n{e}")