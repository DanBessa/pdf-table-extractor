# conversor_protege.py

import pdfplumber
import pandas as pd
import re
import os
import tkinter as tk
from tkinter import filedialog

def clean_valor(valor_str):
    """Limpa e converte uma string de moeda para um número float que o Python entende."""
    if not valor_str or not isinstance(valor_str, str):
        return 0.0
    valor_limpo = valor_str.replace('R$', '').strip().replace('.', '').replace(',', '.')
    try:
        return float(valor_limpo)
    except (ValueError, TypeError):
        return 0.0

def formatar_para_brl(valor_numerico):
    """Formata um número para o padrão de moeda brasileiro (ex: 1.234,56)."""
    valor_str = f"{valor_numerico:,.2f}"
    valor_str = valor_str.replace(",", "TEMP").replace(".", ",").replace("TEMP", ".")
    return valor_str

def iniciar_processamento():
    """
    Função principal chamada pelo menu. Lida com a seleção de arquivos,
    processamento e salvamento dos extratos PROTEGE Cash.
    """
    # Esconde a janela raiz do tkinter que não será usada
    root = tk.Tk()
    root.withdraw()

    # Abre a janela para o usuário selecionar um ou mais arquivos
    pdf_paths = filedialog.askopenfilenames(
        title="Selecione os extratos PROTEGE Cash",
        filetypes=[("Arquivos PDF", "*.pdf")]
    )

    # Se o usuário cancelar a seleção, a função retorna None (ou False)
    if not pdf_paths:
        return None

    # Processa cada arquivo que foi selecionado
    for pdf_path in pdf_paths:
        try:
            transacoes = []
            with pdfplumber.open(pdf_path) as pdf:
                texto_completo = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])

            for linha in texto_completo.split('\n'):
                # Verifica se a linha começa com uma data (DD/MM/AAAA)
                if re.match(r'^\d{2}/\d{2}/\d{4}', linha):
                    data = linha.split()[0]
                    historico = "Não Identificado"
                    valor = 0.0

                    valores_monetarios = re.findall(r'R\$\s?[\d.,]+', linha)
                    
                    if len(valores_monetarios) >= 2:
                        valor_transacao_str = valores_monetarios[-2]
                        
                        if "Emissão de TED" in linha or "Pix Enviado" in linha:
                            historico = "Pix Enviado" if "Pix Enviado" in linha else "Emissão de TED"
                            valor = -clean_valor(valor_transacao_str)
                        elif "Deposito Cofre" in linha:
                            historico = "Deposito Cofre"
                            valor = clean_valor(valor_transacao_str)
                        
                        if valor != 0:
                            transacoes.append({
                                "Data": data,
                                "Histórico": historico,
                                "Valor": valor
                            })

            if not transacoes:
                print(f"Nenhuma transação encontrada no arquivo: {os.path.basename(pdf_path)}")
                continue

            df = pd.DataFrame(transacoes)
            df['Valor'] = df['Valor'].apply(formatar_para_brl)
            
            caminho_csv = os.path.splitext(pdf_path)[0] + ".csv"
            df.to_csv(caminho_csv, index=False, sep=";", encoding="utf-8-sig")
            print(f"Arquivo salvo: {caminho_csv}")

        except Exception as e:
            # Se ocorrer um erro, ele será capturado e exibido pelo menuprincipal.py
            raise Exception(f"Falha ao processar o arquivo '{os.path.basename(pdf_path)}'.\n\nDetalhes: {e}")
            
    # Retorna True para indicar ao menu principal que o processo foi concluído
    return True