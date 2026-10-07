# conversor_sicoobmod1.py (Corrigido)

import pdfplumber
import pandas as pd
from tkinter import filedialog # Mantém apenas o filedialog
import os
import re
# REMOVIDO: messagebox

def extrair_dados_do_pdf(caminho_pdf):
    """
    Extrai dados de transações (Modelo 1) e RETORNA um DataFrame.
    """
    date_pattern = re.compile(r"^(\d{2}\/\d{2}\/\d{4})")
    value_pattern = re.compile(r"([\d\.,]+)([CD])$")
    
    transacoes = []
    data_atual = None

    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            for page in pdf.pages:
                texto_pagina = page.extract_text(x_tolerance=2)
                if not texto_pagina:
                    continue

                linhas = texto_pagina.split('\n')
                
                for linha in linhas:
                    if "SALDO ANTERIOR" in linha or "SALDO DO DIA" in linha or "EXTRATO CONTA CORRENTE" in linha:
                        continue

                    match_data = date_pattern.search(linha)
                    if match_data:
                        data_atual = match_data.group(1)

                    match_valor = value_pattern.search(linha.strip())
                    if match_valor and data_atual:
                        valor_original = f"{match_valor.group(1)}{match_valor.group(2)}"
                        lancamento = linha[:match_valor.start()].strip()
                        
                        if match_data:
                            lancamento = lancamento[match_data.end():].strip()
                        
                        lancamento = re.sub(r"^\S+\s", "", lancamento, count=1)

                        if lancamento:
                            transacoes.append([data_atual, lancamento.strip(), valor_original])

        if not transacoes:
            return pd.DataFrame()

        df = pd.DataFrame(transacoes, columns=["Data", "Lancamento", "Valor_Original"])

        def formatar_valor(valor_str):
            """
            Formata o valor para o padrão CSV brasileiro.
            Ex: "1.234,56D" -> "-1234,56"
            """
            is_debit = valor_str.endswith('D')
            valor_numerico_str = valor_str[:-1]
            valor_sem_ponto = valor_numerico_str.replace('.', '')
            
            if is_debit:
                return '-' + valor_sem_ponto
            else:
                return valor_sem_ponto

        df['Valor'] = df['Valor_Original'].apply(formatar_valor)
        df_final = df[["Data", "Lancamento", "Valor"]]
        
        return df_final

    except Exception as e:
        # Relança o erro para o app principal
        raise Exception(f"Ocorreu um erro ao processar o PDF:\n{e}")

def iniciar_processamento():
    """Função chamada pelo programa principal para iniciar a conversão."""
    
    # Esta chamada será interceptada pelo patch
    path = filedialog.askopenfilename(
        title="Selecione o extrato (Sicoob - Modelo 1)",
        filetypes=[("PDF", "*.pdf")]
    )
    if not path:
        raise UserWarning("Nenhum arquivo selecionado.") 

    df_transacoes = extrair_dados_do_pdf(path)
    
    if df_transacoes is None or df_transacoes.empty:
        raise UserWarning("Nenhuma transação válida encontrada no arquivo selecionado.")

    csv_path = os.path.splitext(path)[0] + '.csv'
    
    # Renomeia a coluna para o padrão do gerador de TXT
    df_transacoes.rename(columns={"Lancamento": "Histórico"}, inplace=True)
    
    df_transacoes.to_csv(csv_path, index=False, sep=';', encoding='utf-8-sig')

    return csv_path # Retorna o caminho para a mensagem de sucesso

# REMOVIDO: Lógica de múltiplos arquivos e salvamento de Excel