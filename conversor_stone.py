import pandas as pd
import unicodedata
import pdfplumber
import os

# Função para remover caracteres especiais e normalizar o texto
def remover_caracteres_invalidos(texto):
    if isinstance(texto, str):
        return ''.join(c for c in unicodedata.normalize('NFKD', texto) if c.isascii()).replace("|", " ")
    return ""

# Função para extrair dados do PDF
def extrair_dados_pdf(caminho_pdf):
    dados_extraidos = []
    
    with pdfplumber.open(caminho_pdf) as pdf:
        for page in pdf.pages:
            linhas = page.extract_text().split('\n') if page.extract_text() else []
            for i in range(len(linhas)):
                partes = linhas[i].split()
                
                if len(partes) > 3 and '/' in partes[0]:
                    data = partes[0]  # Primeira posição é a data
                    valor = partes[-2] if len(partes) >= 2 else ""  # O penúltimo elemento geralmente é o valor
                    
                    historico = " ".join(partes[1:-2]) if len(partes) > 2 else ""
                    
                    if i + 1 < len(linhas):
                        proxima_linha = linhas[i + 1].split()
                        if len(proxima_linha) > 1 and not any(c.isdigit() for c in proxima_linha[-2:]):
                            historico += " " + " ".join(proxima_linha)
                    
                    if data and historico and valor:
                        dados_extraidos.append((data, historico, valor))
    
    return dados_extraidos

# Fluxo principal do programa
def selecionar_pdfs(pdf_path):
    if not pdf_path:
        print("Nenhum arquivo PDF fornecido.")
        return
    
    dados = extrair_dados_pdf(pdf_path)

    if dados:
        df = pd.DataFrame(dados, columns=["Data", "Lancamento", "Valor"])
        
        # Remover caracteres especiais
        df = df.map(remover_caracteres_invalidos)
        
        # Adicionar hífen nos valores de débito
        df["Valor"] = df.apply(
            lambda row: "-" + row["Valor"] if row["Lancamento"].split()[0].lower() in ["debito", "tarifa", "transferencia", "saque"] and not row["Valor"].startswith("-") else row["Valor"], 
            axis=1
        )
        
        # Gerar o caminho correto para salvar o CSV
        caminho_csv = os.path.splitext(pdf_path)[0] + ".csv"
        df.to_csv(caminho_csv, index=False, sep=";")
        
        print(f"Arquivo salvo em: {caminho_csv}")
    else:
        print("Nenhum dado extraído do PDF.")
