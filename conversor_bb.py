import pdfplumber
import csv
import unicodedata
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import os
import re  

def remover_caracteres_especiais(texto):
    """ Remove acentos e caracteres especiais. """
    return ''.join(
        c for c in unicodedata.normalize('NFKD', texto) 
        if not unicodedata.combining(c)
    )

def formatar_valor_financeiro(valor, proximo_texto):
    """ Se o valor terminar com 'D' ou o próximo texto for 'D', transforma corretamente antes de remover. """
    valor = valor.strip().replace(" ", "")  
    valor = remover_caracteres_especiais(valor)  

    if valor.endswith("D") or valor.endswith("d") or proximo_texto in ["D", "d"]:  
        valor = "-" + valor[:-1].strip() if valor.endswith(("D", "d")) else "-" + valor

    valor = re.sub(r"[DCdc]$", "", valor)  # Remove apenas se estiver no final
    return valor

def extract_data_from_pdf(pdf_path, codigo_banco, codigo_aplicacao):
    extracted_data = []
    current_date = ""
    capture_next = False  
    historico_y = None  

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            words = page.extract_words()
            current_history = ""  

            i = 0
            while i < len(words):
                word = words[i]
                x0 = float(word["x0"])
                y0 = float(word["top"])  
                text = remover_caracteres_especiais(word["text"].strip())

                # Captura a Data
                if 50 <= x0 <= 90 and "/" in text and len(text) == 10:
                    current_date = text
                    capture_next = True  
                    historico_y = None  
                    i += 1
                    continue

                # Captura a PRIMEIRA linha do histórico e ignora a segunda
                if 200 <= x0 <= 360 and capture_next:
                    if historico_y is None:  
                        historico_y = y0
                        current_history = text
                    elif abs(y0 - historico_y) < 2:  
                        current_history += " " + text
                    else:  
                        capture_next = False  

                # Captura o Valor financeiro e aplica a formatação correta
                elif 470 <= x0 <= 510:
                    proximo_texto = words[i + 1]["text"].strip() if i + 1 < len(words) else ""  
                    valor_formatado = formatar_valor_financeiro(text, proximo_texto)  

                    # 🔍 Normaliza o histórico para evitar problemas de detecção
                    current_history_normalized = remover_caracteres_especiais(current_history).lower()

                    # 📌 Se o histórico contém "Rende Fácil"
                    if "rende facil" in current_history_normalized:  
                        if valor_formatado.startswith("-"):  # Se for negativo
                            debito, credito = codigo_aplicacao, codigo_banco
                        else:  # Se for positivo
                            debito, credito = codigo_banco, codigo_aplicacao
                    else:
                        debito, credito = (codigo_banco, "") if valor_formatado.startswith("-") else ("", codigo_banco)

                    # Se débito ou crédito ficarem vazios, insere '6'
                    debito = debito if debito else "6"
                    credito = credito if credito else "6"

                    if current_history:  
                        extracted_data.append([debito, credito, current_date, valor_formatado, current_history])  

                    current_history = ""  
                    historico_y = None  
                
                i += 1  

    # Criar o nome do arquivo CSV na mesma pasta do PDF
    output_csv = os.path.splitext(pdf_path)[0] + ".csv"

    with open(output_csv, "w", newline="", encoding="utf-8") as csvfile:
        writer = csv.writer(csvfile, delimiter=";")  
        writer.writerow(["Débito", "Crédito", "Data", "Valor", "Histórico"])  # ✅ Ajuste na ordem das colunas
        for linha in extracted_data:
            writer.writerow(linha)

    print(f"✅ CSV gerado: {output_csv}")  # Exibe no terminal para acompanhar a conversão

def selecionar_pdfs():
    root = tk.Tk()
    root.withdraw()  
    
    # Pergunta o código do banco ao usuário
    codigo_banco = simpledialog.askstring("Código do Banco", "Digite o código do banco:")
    if not codigo_banco:
        messagebox.showwarning("Aviso", "Nenhum código foi inserido. Cancelando operação.")
        return

    # Pergunta o código da conta de aplicação
    codigo_aplicacao = simpledialog.askstring("Código da Conta de Aplicação", "Digite o código da conta de aplicação:")
    if not codigo_aplicacao:
        messagebox.showwarning("Aviso", "Nenhum código de aplicação foi inserido. Cancelando operação.")
        return

    # Seleciona vários arquivos PDF
    pdf_paths = filedialog.askopenfilenames(title="Selecione os arquivos PDF", filetypes=[("Arquivos PDF", "*.pdf")])
    
    if not pdf_paths:
        messagebox.showwarning("Aviso", "Nenhum arquivo PDF foi selecionado.")
        return

    # Processa cada PDF selecionado
    for pdf_path in pdf_paths:
        extract_data_from_pdf(pdf_path, codigo_banco, codigo_aplicacao)

    messagebox.showinfo("Sucesso", "✅ Todos os PDFs foram processados com sucesso!")

if __name__ == "__main__":
    selecionar_pdfs()
