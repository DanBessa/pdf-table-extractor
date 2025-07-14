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
    """
    Abre uma janela para o usuário selecionar um ou mais arquivos PDF.
    Retorna uma tupla com os caminhos dos arquivos selecionados.
    """
    root = tk.Tk()
    root.withdraw()
    # askopenfilenames retorna uma tupla de strings com os caminhos dos arquivos
    pdf_paths = filedialog.askopenfilenames(
        title="Selecione os extratos PagBank", 
        filetypes=[("Arquivos PDF", "*.pdf")]
    )

    # CORREÇÃO: Adicionado o retorno dos caminhos dos arquivos selecionados.
    return pdf_paths

def extrair_texto_pdf(pdf_path):
    """
    Extrai transações de um único arquivo PDF e o salva como CSV.
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])

        pattern_corrected = re.compile(r"(\d{2}/\d{2}/\d{4})\s+(.+?)\s+(-?R?\$\s?[\d\.]+,\d{2})")
        matches = pattern_corrected.findall(text)

        if not matches:
            print(f"Nenhuma transação encontrada no arquivo: {pdf_path}")
            # Silenciosamente ignora ou pode mostrar um aviso por arquivo
            return

        df = pd.DataFrame(matches, columns=["Data", "Descrição", "Valor"])
        caminho_csv = os.path.splitext(pdf_path)[0] + ".csv"
        df.to_csv(caminho_csv, index=False, sep=";", encoding="utf-8-sig")

        print(f"Arquivo salvo em: {caminho_csv}")
        # A mensagem de sucesso agora será controlada pelo script principal (conversor.py)
        # para não aparecer para cada arquivo.
        
    except Exception as e:
        messagebox.showerror("Erro ao Processar Arquivo", f"Não foi possível processar o arquivo:\n{pdf_path}\n\nErro: {e}")


# --- Bloco de Execução Principal (Corrigido) ---
# Este bloco serve para testar o script de forma independente.
if __name__ == "__main__":
    # CORREÇÃO: A lógica foi reestruturada para ser clara e funcional.
    
    # 1. Chama a função de seleção UMA VEZ e armazena o resultado.
    caminhos_pdf = selecionar_pdfs()

    # 2. Verifica se o usuário selecionou algum arquivo.
    if not caminhos_pdf:
        messagebox.showwarning("Aviso", "Nenhum arquivo PDF foi selecionado.")
    else:
        # 3. Itera sobre cada caminho de arquivo na tupla retornada.
        for caminho_do_arquivo in caminhos_pdf:
            extrair_texto_pdf(caminho_do_arquivo)
        
        # 4. Mostra uma mensagem de sucesso única no final do processo.
        messagebox.showinfo("Sucesso", f"{len(caminhos_pdf)} arquivo(s) processado(s) com sucesso!")