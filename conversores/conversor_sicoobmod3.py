import pandas as pd
import pdfplumber
import re
import os
from tkinter import filedialog
from typing import List
import datetime
import traceback

def extrair_dados_sicoob(pdf_path: str) -> pd.DataFrame:
    print(f"[DEBUG] Iniciando extração do arquivo: {pdf_path}")
    transacoes = []
    ano_atual = datetime.date.today().year
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            print(f"[DEBUG] PDF aberto com sucesso. Total de páginas: {len(pdf.pages)}")
            texto_completo = ""
            for i, pagina in enumerate(pdf.pages):
                texto_pagina = pagina.extract_text()
                if texto_pagina:
                    texto_completo += texto_pagina + "\n"
            
            ano_match = re.search(r'Período:\s*\d{2}/\d{2}/(\d{4})', texto_completo)
            if ano_match:
                ano_atual = int(ano_match.group(1))

            linhas_texto = texto_completo.split('\n')
            
            for linha in linhas_texto:
                if not linha.strip():
                    continue
                
                padrao = re.compile(
                    r"^(?P<data>\d{2}/\d{2}(?:/\d{4})?)\s+"  
                    r"(?P<meio>.*)\s+"                      
                    r"(?P<valor>R\$\s*[\d\.]*,\d{2}\s?[CD])$" 
                )
                
                match = padrao.match(linha.strip())

                if match:
                    dados = match.groupdict()
                    
                    meio_str = dados['meio'].strip()
                    meio_upper = meio_str.upper()
                    
                    # ---> BLOQUEIO IMPLACÁVEL DE SALDOS <---
                    if meio_upper.startswith("SALDO") or "SALDO DO DIA" in meio_upper or "S A L D O" in meio_upper or "SALDO ATUAL" in meio_upper:
                        continue

                    data_str = dados['data']
                    data = f"{data_str}/{ano_atual}" if len(data_str) < 10 else data_str
                    
                    valor_str = dados['valor']
                    tipo_transacao = valor_str[-1]
                    valor_limpo_str = re.sub(r'[^\d,]', '', valor_str)
                    
                    try:
                        valor_numerico = float(valor_limpo_str.replace(',', '.'))
                        if tipo_transacao == 'D':
                            valor_numerico = -abs(valor_numerico)
                    except ValueError:
                        continue 

                    documento = ""
                    historico = meio_str
                    
                    match_doc = re.match(r"^([\d\w]+)\s+(.*)", meio_str)
                    if match_doc:
                        doc_candidate = match_doc.group(1)
                        if doc_candidate.isdigit() or doc_candidate.lower() in ['pix', 'cashback']:
                            documento = doc_candidate
                            historico = match_doc.group(2)

                    transacoes.append({
                        'Data': data,
                        'Histórico': historico.strip(),
                        'Documento': documento,
                        'Valor': valor_numerico
                    })

        df = pd.DataFrame(transacoes)
        print(f"[DEBUG] DataFrame criado com {len(df)} registros")
        return df

    except Exception as e:
        print(f"[DEBUG] ERRO durante a extração: {e}")
        print(f"[DEBUG] Traceback completo: {traceback.format_exc()}")
        raise Exception(f"Ocorreu um erro ao processar o arquivo PDF '{os.path.basename(pdf_path)}'.\n\nDetalhes: {e}")

def iniciar_processamento(caminho_pdf=None):
    print("[DEBUG] Iniciando processamento do Sicoob Modelo 3")
    
    if not caminho_pdf:
        caminho_pdf = filedialog.askopenfilename(
            title="Selecione o extrato do Sicoob (Modelo 3)",
            filetypes=[("Arquivos PDF", "*.pdf")]
        )

    if not caminho_pdf:
        raise UserWarning("Nenhum arquivo selecionado.")
        
    if isinstance(caminho_pdf, (list, tuple)):
        caminho_pdf = caminho_pdf[0]

    df_transacoes = extrair_dados_sicoob(caminho_pdf)
    if df_transacoes.empty:
        raise UserWarning("Nenhuma transação válida foi encontrada no arquivo selecionado.")

    base_name = os.path.splitext(caminho_pdf)[0]
    save_path = base_name + ".xlsx"

    try:
        df_transacoes.to_excel(save_path, index=False)
        print(f"[DEBUG] Arquivo salvo com sucesso: {save_path}")
        return save_path 
    except Exception as e:
        print(f"[DEBUG] Erro ao salvar: {e}")
        raise Exception(f"Não foi possível salvar o arquivo.\n\nErro: {e}")

main = iniciar_processamento

if __name__ == "__main__":
    iniciar_processamento()