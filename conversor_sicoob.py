import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import unicodedata
import re

def remover_caracteres(texto):
    """Remove caracteres especiais e normaliza o texto."""
    texto = unicodedata.normalize("NFKD", texto)  # Remove acentos
    texto = re.sub(r"[^\w\s,/.-]", "", texto)  # Mantém letras, números, espaços e alguns caracteres úteis
    return texto.strip()

def selecionar_pdf():
    """Abre uma janela para selecionar o arquivo PDF."""
    caminho_pdf = filedialog.askopenfilename(
        title="Selecione o arquivo PDF",
        filetypes=[("Arquivos PDF", "*.pdf")])
    
    if not caminho_pdf:
        messagebox.showwarning("Aviso", "Nenhum arquivo selecionado!")
        return
    
    extrair_dados(caminho_pdf)

def extrair_dados(pdf_path):
    """Extrai as colunas Data, Lançamento e Valor do PDF e salva em CSV."""
    data_list = []
    lancamento_list = []
    valor_list = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                lines = text.split("\n")
                for line in lines:
                    parts = line.split()
                    if len(parts) >= 3:
                        data = parts[0]  # Primeira parte é a data
                        valor = parts[-1]  # Última parte é o valor
                        lancamento = " ".join(parts[1:-1])  # O restante é o lançamento

                        # Verifica se a data está no formato esperado (DD/MM)
                        if "/" in data and len(data) == 5:
                            data_list.append(remover_caracteres(data))
                            lancamento_list.append(remover_caracteres(lancamento))
                            valor_list.append(remover_caracteres(valor))

    if not data_list:
        messagebox.showerror("Erro", "Nenhum dado foi extraído do PDF.")
        return

    # Criar DataFrame
    df = pd.DataFrame({"Data": data_list, "Lançamento": lancamento_list, "Valor": valor_list})

    # Criar o caminho do CSV no mesmo local do PDF
    caminho_csv = os.path.splitext(pdf_path)[0] + ".csv"

    # Salvar o CSV
    df.to_csv(caminho_csv, index=False, sep=";", encoding="utf-8")
    
    messagebox.showinfo("Sucesso", f"Arquivo CSV salvo em:\n{caminho_csv}")

# Criar janela Tkinter
root = tk.Tk()
root.withdraw()  # Oculta a janela principal

if __name__ == "__main__":
    selecionar_pdf()
