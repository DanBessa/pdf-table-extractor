# conversor_safra.py (VERSÃO FINAL - CORREÇÃO DO AVISO)

import pdfplumber
import pandas as pd
import re
import os
import tkinter as tk
from tkinter import filedialog
from collections import defaultdict

def clean_valor(valor_str):
    """
    Limpa e converte uma string de moeda para um número float, lidando
    com os formatos "1.234,56" e "5.00" do extrato Safra.
    """
    if not valor_str or not isinstance(valor_str, str):
        return 0.0
    
    cleaned_str = valor_str.strip()
    
    if ',' in cleaned_str:
        cleaned_str = cleaned_str.replace('.', '').replace(',', '.')
    
    cleaned_str = re.sub(r'[^\d.-]', '', cleaned_str)

    try:
        return float(cleaned_str)
    except (ValueError, TypeError):
        return 0.0

def formatar_para_brl(valor_numerico):
    """Formata um número para o padrão de moeda brasileiro (ex: 1.234,56)."""
    valor_str = f"{valor_numerico:,.2f}"
    valor_str = valor_str.replace(",", "TEMP").replace(".", ",").replace("TEMP", ".")
    return valor_str

def iniciar_processamento():
    """
    Função principal que será chamada pelo menu. Lida com a seleção, 
    processamento e salvamento dos extratos do Banco Safra.
    """
    root = tk.Tk()
    root.withdraw()

    pdf_paths = filedialog.askopenfilenames(
        title="Selecione os extratos do Banco Safra",
        filetypes=[("Arquivos PDF", "*.pdf")]
    )

    if not pdf_paths:
        return None

    for pdf_path in pdf_paths:
        try:
            todas_transacoes = []
            with pdfplumber.open(pdf_path) as pdf:
                data_atual = None
                for page in pdf.pages:
                    # Lógica de extração avançada baseada em coordenadas de palavras
                    palavras = page.extract_words(x_tolerance=2, y_tolerance=2, keep_blank_chars=False)

                    linhas = defaultdict(list)
                    for p in palavras:
                        linhas[round(p['top'], 0)].append(p)

                    for top in sorted(linhas.keys()):
                        palavras_na_linha = sorted(linhas[top], key=lambda p: p['x0'])
                        linha_texto = " ".join([p['text'] for p in palavras_na_linha])

                        match_data = re.fullmatch(r"(\d{2}/\d{2}/\d{4})", linha_texto.strip())
                        if match_data:
                            data_atual = match_data.group(1)
                            continue
                        
                        if "Descrição" in linha_texto and "Valor (R$)" in linha_texto:
                            continue

                        descricao_palavras = []
                        valor_palavras = []
                        for palavra in palavras_na_linha:
                            if palavra['x0'] < 390:
                                descricao_palavras.append(palavra['text'])
                            else:
                                valor_palavras.append(palavra['text'])
                        
                        if data_atual and valor_palavras:
                            descricao_final = " ".join(descricao_palavras)
                            valor_str = "".join(valor_palavras)
                            valor_numerico = clean_valor(valor_str)

                            if valor_numerico != 0 and descricao_final:
                                todas_transacoes.append({
                                    "Data": data_atual,
                                    "Descrição": descricao_final,
                                    "Valor": valor_numerico
                                })

            if not todas_transacoes:
                print(f"Nenhuma transação encontrada no arquivo: {os.path.basename(pdf_path)}")
                continue

            df = pd.DataFrame(todas_transacoes)
            df['Valor'] = df['Valor'].apply(formatar_para_brl)
            
            caminho_csv = os.path.splitext(pdf_path)[0] + ".csv"
            df.to_csv(caminho_csv, index=False, sep=";", encoding="utf-8-sig")
            print(f"Arquivo salvo: {caminho_csv}")

        except Exception as e:
            raise Exception(f"Falha ao processar o arquivo '{os.path.basename(pdf_path)}'.\n\nDetalhes: {e}")
            
    return True