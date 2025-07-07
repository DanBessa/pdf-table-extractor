import pdfplumber
import re
import pandas as pd
from typing import Optional
import os # Importado para manipular nomes e caminhos de arquivos

# Adiciona as importações necessárias para a interface gráfica
import tkinter as tk
from tkinter import filedialog

def limpar_e_converter_valor(valor_str: Optional[str]) -> float:
    """
    Converte a string de valor do extrato para um número float.
    Ex: "40.000,00 D" -> -40000.0
    Ex: "200,00 C" -> 200.0
    """
    if not valor_str:
        return 0.0

    valor_limpo = valor_str.replace('.', '').replace(',', '.').strip()

    valor_final = 0.0
    match = re.search(r'([\d\.]+)\s*([CD])', valor_limpo)
    if match:
        valor_numerico, tipo = match.groups()
        valor_final = float(valor_numerico)
        if tipo == 'D':
            valor_final *= -1
            
    return valor_final

def extrair_transacoes_de_pdf(caminho_pdf: str) -> Optional[pd.DataFrame]:
    """
    Extrai as transações de um extrato bancário em PDF com um layout específico.
    """
    transacoes = []
    padrao_linha_transacao = re.compile(r'^\d{2}/\d{2}/\d{2,4}')

    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            linhas_texto = []
            for pagina in pdf.pages:
                texto_pagina = pagina.extract_text(x_tolerance=2, y_tolerance=3)
                if texto_pagina:
                    linhas_texto.extend(texto_pagina.split('\n'))
            
            transacao_atual = None
            for linha in linhas_texto:
                if padrao_linha_transacao.search(linha):
                    if transacao_atual:
                        descricao_completa = ' '.join(transacao_atual['Lançamento']).strip()
                        transacao_atual['Lançamento'] = re.sub(r'\s+', ' ', descricao_completa)
                        transacoes.append(transacao_atual)

                    partes = linha.split()
                    data = partes[0]
                    
                    valor_str = None
                    valor_match = re.search(r'([\d\.,]+\s[CD])$', linha)
                    if valor_match:
                        valor_str = valor_match.group(1)

                    descricao_inicial = linha.replace(data, '', 1).strip()
                    if valor_str:
                        descricao_inicial = descricao_inicial.replace(valor_str, '').strip()

                    transacao_atual = {
                        "Data": data,
                        "Lançamento": [descricao_inicial],
                        "Valor": limpar_e_converter_valor(valor_str)
                    }
                elif transacao_atual:
                    if not re.search(r'(Lançamentos|Histórico|Saldo Anterior|SALDO|G336)', linha):
                        transacao_atual['Lançamento'].append(linha.strip())
            
            if transacao_atual:
                descricao_completa = ' '.join(transacao_atual['Lançamento']).strip()
                transacao_atual['Lançamento'] = re.sub(r'\s+', ' ', descricao_completa)
                transacoes.append(transacao_atual)
        
        if not transacoes:
            return None

        df = pd.DataFrame(transacoes)
        df = df[~df['Lançamento'].str.contains("Saldo Anterior", na=False)]
        df = df[df['Valor'] != 0.0]
        
        return df

    except Exception as e:
        print(f"Ocorreu um erro ao processar o arquivo {os.path.basename(caminho_pdf)}: {e}")
        return None

# --- Bloco principal para execução do script ---
if __name__ == '__main__':
    # --- Parte 1: Seleção dos arquivos PDF de entrada ---
    root = tk.Tk()
    root.withdraw()

    caminhos_dos_arquivos = filedialog.askopenfilenames(
        title="Selecione os extratos em PDF para processar",
        filetypes=[("Arquivos PDF", "*.pdf"), ("Todos os arquivos", "*.*")]
    )

    # --- Parte 2: Processar cada arquivo e salvar individualmente ---
    if not caminhos_dos_arquivos:
        print("Nenhum arquivo foi selecionado. Encerrando o programa.")
    else:
        sucessos = 0
        print(f"\n{len(caminhos_dos_arquivos)} arquivo(s) selecionado(s). Iniciando processamento...\n")
        
        for arquivo_path in caminhos_dos_arquivos:
            nome_arquivo_original = os.path.basename(arquivo_path)
            print(f"--- Processando: {nome_arquivo_original} ---")
            
            df_transacoes = extrair_transacoes_de_pdf(arquivo_path)
            
            if df_transacoes is not None and not df_transacoes.empty:
                print(f"{len(df_transacoes)} transações encontradas.")
                
                # Gera o novo nome do arquivo, trocando a extensão .pdf por .csv
                nome_base, _ = os.path.splitext(arquivo_path)
                caminho_csv = nome_base + '.csv'
                
                try:
                    # Salva o DataFrame no novo caminho do CSV
                    df_transacoes.to_csv(caminho_csv, index=False, sep=';', encoding='utf-8-sig')
                    print(f"SUCESSO! Resultado salvo em: {caminho_csv}\n")
                    sucessos += 1
                except Exception as e:
                    print(f"ERRO! Não foi possível salvar o arquivo CSV para '{nome_arquivo_original}'. Motivo: {e}\n")
                    
            else:
                print("Nenhuma transação válida foi encontrada neste arquivo.\n")

        print("="*60)
        print(f"Processamento concluído. {sucessos} de {len(caminhos_dos_arquivos)} arquivo(s) foram convertidos com sucesso.")
        print("="*60)
