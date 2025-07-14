# conversor_sicoobmod1.py

import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import re

def extrair_dados_do_pdf(caminho_pdf):
    """
    Extrai dados de transações (Modelo 1) e RETORNA um DataFrame.
    """
    # Esta regex de data é muito específica e pode falhar.
    # O ideal é usar o método do Modelo 2, mas mantendo este para compatibilidade.
    date_pattern = re.compile(r"^(\d{2}\/\d{2}\/\d{4})")
    value_pattern = re.compile(r"([\d\.,]+)([CD])$")
    
    transacoes = []
    data_atual = None

    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            for page in pdf.pages:
                texto_pagina = page.extract_text(x_tolerance=2)
                if not texto_pagina:
                    continue

                linhas = texto_pagina.split('\n')
                
                for linha in linhas:
                    # Lógica de extração original
                    if "SALDO ANTERIOR" in linha or "SALDO DO DIA" in linha or "EXTRATO CONTA CORRENTE" in linha:
                        continue

                    match_data = date_pattern.search(linha)
                    if match_data:
                        data_atual = match_data.group(1)

                    match_valor = value_pattern.search(linha.strip())
                    if match_valor and data_atual:
                        valor = f"{match_valor.group(1)}{match_valor.group(2)}"
                        lancamento = linha[:match_valor.start()].strip()
                        
                        if match_data:
                            lancamento = lancamento[match_data.end():].strip()
                        
                        lancamento = re.sub(r"^\S+\s", "", lancamento, count=1)

                        if lancamento:
                            transacoes.append([data_atual, lancamento.strip(), valor])

        if not transacoes:
            return pd.DataFrame() # Retorna DataFrame vazio se não achar nada

        df = pd.DataFrame(transacoes, columns=["Data", "Lancamento", "Valor_Original"])

        def formatar_valor(valor_str):
            valor_limpo = valor_str.replace('.', ',')
            if valor_limpo.endswith('D'):
                return '-' + valor_limpo[:-1]
            elif valor_limpo.endswith('C'):
                return valor_limpo[:-1]
            return valor_limpo

        df['Valor'] = df['Valor_Original'].apply(formatar_valor)
        df_final = df[["Data", "Lancamento", "Valor"]]
        
        # ALTERAÇÃO: A função agora retorna o DataFrame.
        return df_final

    except Exception as e:
        messagebox.showerror("Erro no Modelo 1", f"Ocorreu um erro ao processar o PDF:\n{e}")
        return None # Retorna None para indicar que houve uma exceção

def iniciar_processamento():
    """Função chamada pelo programa principal para iniciar a conversão."""
    caminhos_dos_arquivos = filedialog.askopenfilenames(
        title="Selecione os extratos (Sicoob - Modelo 1)",
        filetypes=[("Arquivos PDF", "*.pdf")]
    )
    if not caminhos_dos_arquivos:
        raise UserWarning("Nenhum arquivo foi selecionado.")

    sucessos, erros = 0, []
    for arquivo_path in caminhos_dos_arquivos:
        nome_arquivo_original = os.path.basename(arquivo_path)
        try:
            # Chama a função de extração que agora retorna o DataFrame
            df_transacoes = extrair_dados_do_pdf(arquivo_path)
            
            # A lógica abaixo agora funciona, pois df_transacoes não será None em caso de sucesso
            if df_transacoes is not None and not df_transacoes.empty:
                nome_base, _ = os.path.splitext(arquivo_path)
                caminho_csv = nome_base + '.csv'
                df_transacoes.to_csv(caminho_csv, index=False, sep=';', encoding='utf-8-sig', decimal=',')
                print(f"SUCESSO! '{nome_arquivo_original}' salvo em: {caminho_csv}")
                sucessos += 1
            elif df_transacoes is not None: # Caso o DataFrame esteja vazio
                 print(f"AVISO! Nenhuma transação encontrada em '{nome_arquivo_original}'.")
                 erros.append(f"{nome_arquivo_original} (vazio)")
            else: # Caso df_transacoes seja None (erro na extração)
                erros.append(nome_arquivo_original)

        except Exception as e:
            print(f"ERRO! Falha em '{nome_arquivo_original}': {e}")
            erros.append(nome_arquivo_original)
    
    if erros:
        raise Exception(f"Falha ao processar {len(erros)} arquivo(s): {', '.join(erros)}")
    if sucessos == 0:
        raise Exception("Nenhuma transação foi extraída dos arquivos selecionados.")

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    # Para testes diretos, pode-se chamar iniciar_processamento()
    try:
        iniciar_processamento()
        messagebox.showinfo("Concluído", "Processamento finalizado.")
    except Exception as e:
        messagebox.showerror("Erro Final", str(e))