# conversor_saframod2.py (CORRIGIDO)

import pandas as pd
import pdfplumber
import re
import os
from tkinter import filedialog # Mantém apenas o filedialog
import datetime 
# REMOVIDO: messagebox

def extrair_dados_extrato_safra(caminho_pdf: str) -> pd.DataFrame:
    """
    Extrai as transações de um extrato do Banco Safra, adicionando o ano atual
    às datas extraídas.
    """
    lancamentos_brutos = []
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            for pagina in pdf.pages:
                tabelas = pagina.extract_tables(table_settings={
                    "vertical_strategy": "text",
                    "horizontal_strategy": "text",
                    "snap_tolerance": 5,
                })
                
                for tabela in tabelas:
                    if not tabela: continue

                    header_row_index = -1
                    for i, linha in enumerate(tabela):
                        linha_str = " ".join(filter(None, [str(celula) for celula in linha]))
                        if "Lançamento" in linha_str and "Valor" in linha_str and "Data" in linha_str:
                            header_row_index = i
                            break
                    
                    if header_row_index != -1:
                        lancamentos_brutos.extend(tabela[header_row_index + 1:])

    except Exception as e:
        raise Exception(f"Não foi possível processar o arquivo PDF.\n\nDetalhes: {e}")

    if not lancamentos_brutos:
        return pd.DataFrame()

    ano_atual = datetime.date.today().year

    dados_processados = []
    for linha in lancamentos_brutos:
        linha_limpa = [str(celula).strip() for celula in linha if celula and str(celula).strip()]
        
        if len(linha_limpa) < 1:
            continue

        primeiro_elemento = linha_limpa[0]
        match = re.match(r'(\d{2}/\d{2})(.*)', primeiro_elemento, re.DOTALL)
        
        if match:
            data_sem_ano = match.group(1).strip()
            data = f"{data_sem_ano}/{ano_atual}"
            
            resto_do_primeiro_elemento = match.group(2).strip()
            
            if len(linha_limpa) < 2:
                continue
            
            valor_str_sujo = linha_limpa[-1]
            # Limpa o valor para o padrão do app principal
            valor_formatado = valor_str_sujo.replace('.', '').replace(',', ',')
            if valor_formatado.endswith('-'):
                valor_formatado = '-' + valor_formatado[:-1]
            
            elementos_descricao = [resto_do_primeiro_elemento] + linha_limpa[1:-1]
            descricao = " ".join(filter(None, elementos_descricao))

            dados_processados.append([data, descricao, valor_formatado])

    if not dados_processados:
        return pd.DataFrame()

    df = pd.DataFrame(dados_processados, columns=["Data", "Histórico", "Valor"])
    
    # Remove valores nulos ou inválidos
    df.dropna(subset=["Valor", "Data"], inplace=True)
    df = df[df["Valor"] != "0,00"]

    return df

def iniciar_processamento():
    """
    Função principal chamada pelo menu da aplicação.
    """
    # Esta chamada será interceptada pelo patch
    caminho_pdf = filedialog.askopenfilename(
        title="Selecione o extrato PDF do Banco Safra (Modelo 2)",
        filetypes=[("Arquivos PDF", "*.pdf")]
    )
    if not caminho_pdf:
        raise UserWarning("Nenhum arquivo selecionado.")

    df_resultado = extrair_dados_extrato_safra(caminho_pdf)

    if df_resultado.empty:
        raise UserWarning("Nenhuma transação válida foi encontrada no arquivo PDF.")

    # Salva como CSV com o mesmo nome do PDF
    caminho_salvar = os.path.splitext(caminho_pdf)[0] + ".csv"

    try:
        df_resultado.to_csv(caminho_salvar, index=False, sep=";", encoding="utf-8-sig")
        return caminho_salvar # Retorna o caminho do CSV
    except Exception as e:
        raise Exception(f"Não foi possível salvar o arquivo CSV.\n\nErro: {e}")