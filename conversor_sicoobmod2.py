# conversor_sicoobmod2.py

import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import re

def extrair_ano_do_pdf(pdf_pages):
    """Extrai o ano da linha 'PERÍODO' na primeira página para construir a data completa."""
    try:
        primeira_pagina_texto = pdf_pages[0].extract_text(x_tolerance=2)
        match = re.search(r"PERÍODO: \d{2}\/\d{2}\/(\d{4})", primeira_pagina_texto)
        if match:
            return match.group(1)
    except Exception:
        pass
    messagebox.showwarning("Ano não encontrado", "Não foi possível determinar o ano do extrato. Usando o ano atual como padrão.")
    return str(pd.Timestamp.now().year)

def extrair_dados_do_pdf(caminho_pdf):
    """
    Extrai dados de um extrato Sicoob (Modelo 2) e RETORNA um DataFrame.
    """
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            ano = extrair_ano_do_pdf(pdf.pages)
            texto_completo = "\n".join([page.extract_text(x_tolerance=2) or "" for page in pdf.pages])
    except Exception as e:
        messagebox.showerror("Erro de Leitura", f"Não foi possível ler o arquivo PDF:\n{e}")
        return None

    texto_completo = re.sub(r".*HISTÓRICO DE MOVIMENTAÇÃO\n", "", texto_completo, flags=re.DOTALL)
    texto_completo = re.sub(r"SALDO ANTERIOR.*?\n", "", texto_completo, flags=re.DOTALL)
    texto_completo = re.sub(r"\nRESUMO.*", "", texto_completo, flags=re.DOTALL)
    
    blocos = re.split(r'\n(?=\d{2}/\d{2})', texto_completo.strip())
    transacoes = []

    for bloco in blocos:
        texto_bloco = re.sub(r'\s{2,}', ' ', bloco.replace('\n', ' ').strip())
        if "SALDO DO DIA" in texto_bloco:
            continue
        
        match_valor_tipo = re.search(r'(\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2}|\d+\.\d{2})\s*([CD])', texto_bloco)
        data_match = re.match(r'(\d{2}/\d{2})', texto_bloco)
        
        if data_match and match_valor_tipo:
            data = f"{data_match.group(1)}/{ano}"
            valor_str = match_valor_tipo.group(1)
            tipo = match_valor_tipo.group(2)
            
            descricao = texto_bloco
            descricao = re.sub(r'^\d{2}/\d{2}\s*', '', descricao).strip()
            descricao = descricao.replace(match_valor_tipo.group(0), '', 1).strip()
            descricao = re.sub(r'\s{2,}', ' ', descricao).strip()

            valor_numerico = float(valor_str.replace('.', '').replace(',', '.'))
            if tipo == 'D':
                valor_numerico *= -1

            if descricao:
                transacoes.append([data, descricao, valor_numerico])

    if not transacoes:
        return pd.DataFrame()

    df = pd.DataFrame(transacoes, columns=["Data", "Lancamento", "Valor"])
    df['Data'] = pd.to_datetime(df['Data'], format='%d/%m/%Y').dt.strftime('%d/%m/%Y')
    
    # ALTERAÇÃO: A função agora retorna o DataFrame.
    return df

def iniciar_processamento():
    """Função chamada pelo programa principal para iniciar a conversão."""
    caminhos_dos_arquivos = filedialog.askopenfilenames(
        title="Selecione os extratos (Sicoob - Modelo 2)",
        filetypes=[("Arquivos PDF", "*.pdf")]
    )
    if not caminhos_dos_arquivos:
        raise UserWarning("Nenhum arquivo foi selecionado.")

    sucessos, erros = 0, []
    for arquivo_path in caminhos_dos_arquivos:
        nome_arquivo_original = os.path.basename(arquivo_path)
        try:
            df_transacoes = extrair_dados_do_pdf(arquivo_path)
            
            if df_transacoes is not None and not df_transacoes.empty:
                nome_base, _ = os.path.splitext(arquivo_path)
                caminho_csv = nome_base + '.csv'
                df_transacoes.to_csv(caminho_csv, index=False, sep=';', encoding='utf-8-sig', decimal=',')
                print(f"SUCESSO! '{nome_arquivo_original}' salvo em: {caminho_csv}")
                sucessos += 1
            elif df_transacoes is not None:
                 print(f"AVISO! Nenhuma transação encontrada em '{nome_arquivo_original}'.")
                 erros.append(f"{nome_arquivo_original} (vazio)")
            else:
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
    try:
        iniciar_processamento()
        messagebox.showinfo("Concluído", "Processamento finalizado.")
    except Exception as e:
        messagebox.showerror("Erro Final", str(e))