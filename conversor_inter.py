import pandas as pd
import pdfplumber
import re
import os
from tkinter import messagebox

# A função agora se chama iniciar_processamento para corresponder ao CONVERTERS
def iniciar_processamento(pdf_path):
    """
    Extrai as transações do Banco Inter a partir de um caminho de PDF
    fornecido e salva em CSV.
    """
    try:
        if not pdf_path or not os.path.exists(pdf_path):
            # Levanta um erro se o caminho do arquivo for inválido
            raise FileNotFoundError("Caminho do arquivo para o conversor Inter é inválido.")

        csv_path = os.path.splitext(pdf_path)[0] + ".csv"
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
                            ultima_data = f"{dia}/{mes_numero}/{ano}"

                        match = valor_pattern.search(line)
                        if match:
                            sinal = match.group(1)
                            valor = match.group(2)
                            historico = line[:match.start()].strip()
                            valor = f"-{valor}" if sinal == "-" else valor
                            valor = re.sub(r"\.(?=\d{3},)", "", valor)
                            historico = historico.replace('"', '').replace("'", "")
                            datas.append(ultima_data)
                            historicos.append(historico)
                            valores.append(valor)

        df = pd.DataFrame({"Data": datas, "Histórico": historicos, "Valor": valores})
        df.to_csv(csv_path, index=False, sep=';', encoding="utf-8-sig")
        
        # O menu principal já mostra uma mensagem de sucesso, então esta é opcional
        # messagebox.showinfo("Sucesso (Inter)", f"Arquivo salvo com sucesso:\n{csv_path}")

    except Exception as e:
        # Propaga o erro para o menu principal para que ele possa exibir a janela de erro detalhada
        raise e