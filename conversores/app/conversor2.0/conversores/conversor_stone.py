# conversor_stone.py (VERSÃO FINAL E CORRIGIDA)

import os
import pdfplumber
import pandas as pd
from tkinter import filedialog, messagebox
import traceback

def extrair_dados_do_pdf(pdf_path):
    """
    Extrai todas as tabelas de um extrato da Stone e as combina em um único DataFrame,
    com uma lógica de detecção de cabeçalho mais robusta.
    """
    dados_completos = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            # A função extract_tables é ideal para este PDF, pois as tabelas são bem definidas
            tabelas = page.extract_tables()
            if not tabelas:
                continue

            for tabela in tabelas:
                # --- LÓGICA DE DETECÇÃO DE CABEÇALHO APRIMORADA ---
                for row in tabela:
                    # Garante que a linha não seja vazia e tenha conteúdo na primeira célula
                    if not row or not row[0]:
                        continue
                    
                    # Se a primeira célula contiver "Data", consideramos um cabeçalho e pulamos
                    if "Data" in row[0]:
                        continue
                    
                    # Garante que a linha tenha o número esperado de colunas para ser uma transação
                    if len(row) == 5:
                        dados_completos.append(row)

    if not dados_completos:
        return pd.DataFrame()

    # Cria o DataFrame final com o cabeçalho correto
    df = pd.DataFrame(dados_completos, columns=['Data', 'Descrição', 'Valor', 'Taxa', 'Líquido'])
    return df

def iniciar_processamento():
    """
    Função de entrada que orquestra o fluxo de conversão.
    Retorna True (sucesso) ou False (cancelamento/falha).
    """
    pdf_path = filedialog.askopenfilename(
        title="Selecione o extrato da Stone",
        filetypes=[("PDF files", "*.pdf")]
    )
    if not pdf_path:
        return False # Usuário cancelou

    try:
        df = extrair_dados_do_pdf(pdf_path)

        if df.empty:
            messagebox.showwarning("Aviso", "Nenhuma tabela de transação válida foi encontrada no arquivo.")
            return False

        output_csv_path = os.path.splitext(pdf_path)[0] + ".csv"
        
        # Salva o arquivo CSV com ; como separador
        df.to_csv(output_csv_path, index=False, sep=';', encoding='utf-8-sig')
        
        return True

    except Exception as e:
        traceback.print_exc()
        messagebox.showerror("Erro Crítico", f"Ocorreu um erro inesperado ao processar o extrato.\n\n{e}")
        return False