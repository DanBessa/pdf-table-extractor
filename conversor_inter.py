import pandas as pd
import pdfplumber
import re
import os
import tkinter as tk
from tkinter import filedialog, messagebox

def selecionar_pdfs(pdf_path=None):
    """Se nenhum caminho for passado, abre a caixa de seleção"""
    if not pdf_path:
        root = tk.Tk()
        root.withdraw()
        pdf_path = filedialog.askopenfilename(title="Banco Inter", filetypes=[("Arquivos PDF", "*.pdf")])

        if not pdf_path:
            messagebox.showwarning("Aviso", "Nenhum arquivo PDF foi selecionado.")
            return None

    return pdf_path

def extrair_dados_inter(pdf_path):
    """Extrai as transações do Banco Inter a partir do PDF e salva em CSV."""
    if not pdf_path or not os.path.exists(pdf_path):
        print("Erro: Caminho do arquivo inválido!")
        return None

    csv_path = os.path.splitext(pdf_path)[0] + ".csv"
    datas, historicos, valores = [], [], []

    meses = {
        "Janeiro": "01", "Fevereiro": "02", "Março": "03", "Abril": "04",
        "Maio": "05", "Junho": "06", "Julho": "07", "Agosto": "08",
        "Setembro": "09", "Outubro": "10", "Novembro": "11", "Dezembro": "12"
    }

    date_pattern = re.compile(r"(\d{1,2}) de (\w+) de (\d{4})")
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
                    
                    elif any(kw in line for kw in ["Pix recebido", "Pix enviado", "Pagamento efetuado"]):
                        parts = line.split("R$")
                        if len(parts) >= 2:
                            historico = parts[0].strip()
                            valor = parts[1].strip().split()[0].replace(".", "").replace(",", ".")
                            datas.append(ultima_data)
                            historicos.append(historico)
                            valores.append(valor)

    df = pd.DataFrame({"Data": datas, "Histórico": historicos, "Valor": valores})
    df.to_csv(csv_path, index=False, sep=';', encoding='utf-8')

    print(f"Arquivo salvo com sucesso: {csv_path}")
    return csv_path

if __name__ == "__main__":
    pdf_path = selecionar_pdfs()
    if pdf_path:
        extrair_dados_inter(pdf_path)
