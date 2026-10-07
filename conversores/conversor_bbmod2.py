import pdfplumber
import re
import pandas as pd
from typing import Optional, List
import os
from tkinter import filedialog, messagebox

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
    padrao_valor_geral = re.compile(r'([\d\.,]+\s[CD])(?!\\w)')

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
                
                partes = linha.split()
                data = partes[0]
                todos_valores_encontrados = padrao_valor_geral.findall(linha)
                valor_str = todos_valores_encontrados[0] if todos_valores_encontrados else None

                descricao_inicial = ' '.join(partes[4:]) if len(partes) > 4 else ''
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

def iniciar_processamento(caminho_pdf=None):
    """
    Função adaptada para ser invocada pelo menu principal (ConversorApp).
    Garante suporte ao interceptador patch_file_dialogs e ao fluxo unificado.
    """
    # Fallback caso a rota principal falhe em passar o parâmetro
    if not caminho_pdf:
        caminho_pdf = filedialog.askopenfilename(
            title="Selecione o extrato (BB - Modelo 1)",
            filetypes=[("Arquivos PDF", "*.pdf")]
        )
    if not caminho_pdf:
        raise UserWarning("Nenhum arquivo selecionado.")

    try:
        df_transacoes = _extrair_transacoes_de_pdf(caminho_pdf)
        
        if df_transacoes is None or df_transacoes.empty:
            raise UserWarning(f"Nenhuma transação válida foi encontrada no arquivo informado.")

        # Padroniza os nomes de colunas esperados pela leitura do Alterdata no menu principal
        df_transacoes = df_transacoes.rename(columns={"Lançamento": "Descrição"})

        # Define salvamento automático em CSV para integração imediata com o gerador do Alterdata
        nome_base, _ = os.path.splitext(caminho_pdf)
        caminho_salvar = nome_base + ".csv"
        
        # Converte a coluna numérica para string substituindo o ponto decimal por vírgula
        df_transacoes['Valor'] = df_transacoes['Valor'].apply(lambda x: str(f"{x:.2f}").replace('.', ','))
        
        df_transacoes.to_csv(caminho_salvar, index=False, sep=';', encoding='utf-8-sig')
        
        # Retorna a string do caminho para o menu gerenciar as caixas de sucesso e o arquivo TXT
        return caminho_salvar

    except Exception as e:
        raise Exception(f"Erro no processador interno do Banco do Brasil: {e}")