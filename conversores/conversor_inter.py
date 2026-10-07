import pandas as pd
import pdfplumber
import re
import os
from tkinter import filedialog # Mantém apenas o filedialog
# REMOVIDO: messagebox

def processar_pdf_inter(pdf_path):
    """
    Extrai dados de um único PDF do Inter e retorna um DataFrame.
    """
    datas, historicos, valores = [], [], []
    
    meses = {
        "Janeiro": "01", "Fevereiro": "02", "Março": "03", "Abril": "04",
        "Maio": "05", "Junho": "06", "Julho": "07", "Agosto": "08",
        "Setembro": "09", "Outubro": "10", "Novembro": "11", "Dezembro": "12"
    }
    
    date_pattern = re.compile(r"(\d{1,2}) de (\w+) de (\d{4})")
    valor_pattern = re.compile(r"(-?)R\$\s*(\d{1,3}(?:\.\d{3})*,\d{2})")
    ultima_data = "01/01/2000"

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                lines = text.split("\n")
                for line in lines:
                    date_match = date_pattern.search(line)
                    if date_match:
                        dia, mes, ano = date_match.groups()
                        mes_numero = meses.get(mes, "00")
                        ultima_data = f"{int(dia):02d}/{mes_numero}/{ano}"

                    match = valor_pattern.search(line)
                    if match:
                        sinal = match.group(1)
                        valor = match.group(2)
                        historico = line[:match.start()].strip()
                        
                        # Limpa o histórico
                        historico = historico.replace('"', '').replace("'", "")
                        # Remove a data do início do histórico se ela existir
                        historico = re.sub(r'^\d{1,2} de \w+ de \d{4}', '', historico).strip()
                        
                        # Evita linhas de "Saldo"
                        if "saldo" in historico.lower():
                            continue

                        # Formata o valor
                        valor_final_str = f"{sinal}{valor}"
                        valor_final_float = float(re.sub(r"\.(?=\d{3},)", "", valor_final_str).replace(',', '.'))
                        
                        if historico and valor_final_float != 0.0:
                            datas.append(ultima_data)
                            historicos.append(historico)
                            valores.append(valor_final_float) # Salva como float

    return pd.DataFrame({"Data": datas, "Histórico": historicos, "Valor": valores})

def iniciar_processamento(pdf_path):
    """
    Função principal que lida com um único arquivo
    e retorna o caminho do CSV.
    """
    # Esta chamada será interceptada pelo patch
    # pdf_path = filedialog.askopenfilename(
    #     title="Selecione o extrato Inter",
    #     filetypes=[("Arquivos PDF", "*.pdf")]
    # )
    if not pdf_path:
        raise UserWarning("Nenhum arquivo selecionado.")

    try:
        df = processar_pdf_inter(pdf_path)
        if df.empty:
            raise UserWarning(f"Nenhuma transação encontrada em {os.path.basename(pdf_path)}.")

        # Salva o CSV com o mesmo nome do PDF
        nome_base, _ = os.path.splitext(pdf_path)
        caminho_csv = nome_base + ".csv"
        
        # Formata o valor para o padrão CSV brasileiro (com vírgula)
        df['Valor'] = df['Valor'].apply(lambda x: str(f"{x:.2f}").replace('.', ','))

        df.to_csv(caminho_csv, index=False, sep=';', encoding='utf-8-sig')
        
        # Retorna o caminho do CSV para o app principal
        return caminho_csv
        
    except Exception as e:
        # Relança a exceção para o app principal
        raise Exception(f"Erro ao processar '{os.path.basename(pdf_path)}':\n{e}")