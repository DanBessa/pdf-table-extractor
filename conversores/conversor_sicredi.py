# conversor_sicredi.py (V2 - Mais Robusto)

import pdfplumber
import pandas as pd
import re
from tkinter import filedialog 
import os

def formatar_para_csv(valor_numerico):
    """
    Formata um número float para o padrão CSV brasileiro (ex: -1234,56 ou 1234,56).
    O f-string f"{valor_numerico:.2f}" garante que o hífen ('-') seja incluído
    para números floats negativos.
    """
    # Se o número é muito grande (milhares), a formatação padrão em Python já lida com o sinal.
    valor_str = f"{valor_numerico:.2f}" 
    return valor_str.replace('.', ',')

def extrair_dados(pdf_path):
    """Extrai dados de um Sicredi PDF usando extração de tabela e retorna um DataFrame."""
    transacoes = []
    date_pattern = re.compile(r'^\d{2}/\d{2}/\d{4}$') 
    
    # Lista de palavras-chave comuns de DÉBITO no Sicredi.
    # Adicionada "CESTA DE RELACIONAMENTO" e removido "APLICACAO" (geralmente crédito)
    DEBIT_KEYWORDS = ["TARIFA", "IOF", "DEVOLUCAO", "LIQ TED", "PAGTO", "PIX ENVIADO", 
                      "RESGATE", "DEB. AUT.", "MED PIX F DEB", "CESTA DE RELACIONAMENTO"]

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tabelas = page.extract_tables()
                for tabela in tabelas:
                    for linha in tabela:
                        if not linha or len(linha) < 4:
                            continue

                        data = str(linha[0]).strip() if linha[0] else ""
                        
                        if not date_pattern.match(data):
                            continue

                        descricao = str(linha[1]).strip().replace('\n', ' ') if linha[1] else ""
                        doc = str(linha[2]).strip().replace('\n', ' ') if linha[2] else ""
                        valor_str = str(linha[3]).strip() if linha[3] else "0"
                        
                        # --- Lógica de Parse do Valor APRIMORADA ---
                        
                        # 1. Limpa o valor para tentar conversão bruta, mantendo o hífen
                        valor_limpo = valor_str.replace('.', '').replace(',', '.')
                        
                        valor = 0.0 # Inicializa
                        
                        try:
                            # Tenta converter diretamente. Se o PDF já incluiu o '-', este é o valor final.
                            valor = float(valor_limpo)
                            
                        except ValueError:
                            # Se falhar (ex: string vazia ou formato estranho), ignora a linha.
                            print(f"Aviso: Linha ignorada em '{os.path.basename(pdf_path)}' por valor inválido: {linha}")
                            continue
                            
                        # 2. Re-avalia o sinal apenas se o valor for positivo (e potencialmente um débito)
                        if valor >= 0:
                            if any(kw in descricao.upper() for kw in DEBIT_KEYWORDS) or \
                               "DEBITO TED/IB" in descricao.upper(): # Adiciona TED/IB como detecção específica
                                
                                # Se detectado como débito, garante o sinal negativo
                                valor = -abs(valor)
                            else:
                                # Garante que créditos sejam positivos
                                valor = abs(valor) 

                        # --- Fim da Lógica de Parse do Valor ---

                        transacoes.append({
                            "Data": data,
                            "Histórico": f"{descricao} {doc}".strip(),
                            "Valor": valor 
                        })
    
        if not transacoes:
            return pd.DataFrame()

        return pd.DataFrame(transacoes)

    except Exception as e:
        print(f"Erro ao processar o arquivo {os.path.basename(pdf_path)}: {e}")
        raise Exception(f"Erro ao processar o arquivo {os.path.basename(pdf_path)}:\n{e}")

def iniciar_processamento(pdf_path):
    """
    Função principal que processa o arquivo PDF e salva como CSV.
    """
    if not pdf_path:
        raise Exception("Nenhum arquivo selecionado.") 

    df_transacoes = extrair_dados(pdf_path)
    
    if df_transacoes.empty:
        raise Exception("Nenhuma transação foi encontrada no arquivo selecionado.")

    save_path = os.path.splitext(pdf_path)[0] + ".csv"

    try:
        # Aplica a formatação que preserva o hífen e usa vírgula decimal
        df_transacoes['Valor'] = df_transacoes['Valor'].apply(formatar_para_csv)
        
        df_final_csv = df_transacoes[['Data', 'Histórico', 'Valor']]
        
        # Salva no formato CSV brasileiro (com separador ;)
        df_final_csv.to_csv(save_path, index=False, sep=';', encoding='utf-8-sig')
        
        return save_path 
    except Exception as e:
        raise Exception(f"Não foi possível salvar o arquivo CSV.\n\nErro: {e}")

# Exemplo de uso (chame esta função no seu script principal, se necessário)
if __name__ == '__main__':
    # Simula a seleção de arquivo para testes:
    # pdf_path = filedialog.askopenfilename(
    #      title="Selecione o extrato Sicredi",
    #      filetypes=[("Arquivos PDF", "*.pdf")]
    # )
    
    # Substitua pelo seu caminho de teste se não estiver usando GUI/Tkinter
    # pdf_path = "caminho/para/seu/extrato_F1_sicredi.pdf" 

    # if 'pdf_path' in locals() and pdf_path:
    #     try:
    #         caminho_csv = iniciar_processamento(pdf_path)
    #         print(f"✅ Sucesso! Arquivo CSV salvo em: {caminho_csv}")
    #     except Exception as e:
    #         print(f"❌ Erro no processamento:\n{e}")
    # else:
    #     print("Nenhum arquivo selecionado ou caminho não definido.")
    pass