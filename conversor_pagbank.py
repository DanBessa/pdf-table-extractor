import pdfplumber
import pandas as pd
import re
import os
import unicodedata
import tkinter as tk
from tkinter import filedialog, messagebox

def remover_caracteres(texto):
    texto = unicodedata.normalize("NFKD", texto)
    texto = re.sub(r"[^\w\s,/.-]", "", texto)
    return texto.strip()

def selecionar_pdfs():
    root = tk.Tk()
    root.withdraw()
    pdf_path = filedialog.askopenfilenames(title="PagBank", filetypes=[("Arquivos PDF", "*.pdf")])

    if not pdf_path:
        messagebox.showwarning("Aviso", "Nenhum arquivo PDF foi selecionado.")
        return


def extrair_texto_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])

    pattern_corrected = re.compile(r"(\d{2}/\d{2}/\d{4})\s+(.+?)\s+(-?R?\$\s?[\d\.]+,\d{2})")
    matches = pattern_corrected.findall(text)

    df = pd.DataFrame(matches, columns=["Data", "Descrição", "Valor"])
    caminho_csv = os.path.splitext(pdf_path)[0] + ".csv"
    df.to_csv(caminho_csv, index=False, sep=";", encoding="utf-8-sig")

    print(f"Arquivo salvo em: {caminho_csv}")
    messagebox.showinfo("Sucesso", f"Arquivo CSV salvo em:\n{caminho_csv}")

if __name__ == "__main__":
    selecionar_pdfs()
