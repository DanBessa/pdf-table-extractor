# Seu código original de 'conversor_caixa.py' parece ser projetado para linha de comando.
# Esta versão é simplificada para ser chamada pelo menu.
import pdfplumber
import pandas as pd
import os
from tkinter import filedialog

def main():
    pdf_path = filedialog.askopenfilename(title="Selecione o PDF da Caixa", filetypes=[("PDF files", "*.pdf")])
    if not pdf_path:
        raise UserWarning("Nenhum arquivo selecionado.")
    
    all_text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                all_text += text + "\n"
    
    # Esta é uma extração simplificada. A lógica complexa do seu script original
    # pode ser colada aqui se esta não for suficiente.
    import re
    transactions = []
    date_pattern = re.compile(r'(\d{2}/\d{2}/\d{4})')
    for line in all_text.split('\n'):
        if date_pattern.search(line):
            parts = line.split()
            if len(parts) > 2:
                date = parts[0]
                description = " ".join(parts[1:-1])
                value = parts[-1]
                transactions.append([date, description, value])
    
    if not transactions:
        raise Exception("Nenhuma transação encontrada no PDF da Caixa.")

    df = pd.DataFrame(transactions, columns=['Data', 'Histórico', 'Valor/Saldo'])
    csv_path = os.path.splitext(pdf_path)[0] + ".csv"
    df.to_csv(csv_path, index=False, sep=';', encoding='utf-8-sig')