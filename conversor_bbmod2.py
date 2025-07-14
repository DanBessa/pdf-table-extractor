import pdfplumber
import re
import pandas as pd
from typing import Optional, List
import os
import tkinter as tk
from tkinter import filedialog

# Se a lógica for diferente, altere estas funções.
# Caso contrário, pode deixar como está.
def _limpar_e_converter_valor(valor_str: Optional[str]) -> float:
    if not valor_str:
        return 0.0
    match = re.search(r'([\d\.,]+)\s*([CD])', valor_str)
    if match:
        valor_numerico, tipo = match.groups()
        valor_limpo = valor_numerico.replace('.', '').replace(',', '.').strip()
        valor_final = float(valor_limpo)
        if tipo == 'D':
            valor_final *= -1
        return valor_final
    return 0.0

def _extrair_transacoes_de_pdf(caminho_pdf: str) -> Optional[pd.DataFrame]:
    transacoes: List[dict] = []
    padrao_linha_transacao = re.compile(r'^\d{2}/\d{2}/\d{2,4}')
    padrao_valor_geral = re.compile(r'([\d\.,]+\s[CD])')

    with pdfplumber.open(caminho_pdf) as pdf:
        linhas_texto: List[str] = []
        for pagina in pdf.pages:
            texto_pagina = pagina.extract_text(x_tolerance=2, y_tolerance=3)
            if texto_pagina:
                linhas_texto.extend(texto_pagina.split('\n'))
        
        transacao_atual = None
        for linha in linhas_texto:
            if padrao_linha_transacao.search(linha):
                if transacao_atual and transacao_atual.get('Valor') is not None:
                    descricao_completa = ' '.join(transacao_atual['Lançamento']).strip()
                    transacao_atual['Lançamento'] = re.sub(r'\s+', ' ', descricao_completa)
                    transacoes.append(transacao_atual)
                
                data = linha.split()[0]
                todos_valores_encontrados = padrao_valor_geral.findall(linha)
                valor_str = todos_valores_encontrados[0] if todos_valores_encontrados else None
                
                descricao_inicial = linha.replace(data, '', 1).strip()
                if valor_str:
                    for v_str in todos_valores_encontrados:
                        descricao_inicial = descricao_inicial.replace(v_str, '').strip()

                transacao_atual = {
                    "Data": data,
                    "Lançamento": [descricao_inicial],
                    "Valor": _limpar_e_converter_valor(valor_str)
                }
            elif transacao_atual:
                if not re.search(r'(Lançamentos|Histórico|Saldo Anterior|SALDO|G336)', linha):
                    transacao_atual['Lançamento'].append(linha.strip())
        
        if transacao_atual and transacao_atual.get('Valor') is not None:
            descricao_completa = ' '.join(transacao_atual['Lançamento']).strip()
            transacao_atual['Lançamento'] = re.sub(r'\s+', ' ', descricao_completa)
            transacoes.append(transacao_atual)
    
    if not transacoes:
        return None

    df = pd.DataFrame(transacoes)
    df = df[~df['Lançamento'].str.contains("Saldo Anterior", na=False)]
    df = df[df['Valor'] != 0.0]
    return df

def iniciar_processamento():
    caminhos_dos_arquivos = filedialog.askopenfilenames(
        title="Selecione os extratos (BB - Modelo 2)",
        filetypes=[("Arquivos PDF", "*.pdf")]
    )
    if not caminhos_dos_arquivos:
        raise UserWarning("Nenhum arquivo foi selecionado.")

    sucessos, erros = 0, []
    for arquivo_path in caminhos_dos_arquivos:
        nome_arquivo_original = os.path.basename(arquivo_path)
        try:
            df_transacoes = _extrair_transacoes_de_pdf(arquivo_path)
            if df_transacoes is not None and not df_transacoes.empty:
                nome_base, _ = os.path.splitext(arquivo_path)
                caminho_csv = nome_base + '.csv'
                df_transacoes.to_csv(caminho_csv, index=False, sep=';', encoding='utf-8-sig', decimal=',')
                print(f"SUCESSO! '{nome_arquivo_original}' salvo em: {caminho_csv}")
                sucessos += 1
            else:
                erros.append(nome_arquivo_original)
        except Exception as e:
            print(f"ERRO! Falha em '{nome_arquivo_original}': {e}")
            erros.append(nome_arquivo_original)
    
    if erros:
        raise Exception(f"Falha ao processar {len(erros)} arquivo(s): {', '.join(erros)}")
    if sucessos == 0:
        raise Exception("Nenhuma transação foi extraída dos arquivos selecionados.")