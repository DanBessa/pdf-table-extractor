import tkinter as tk
from tkinter import filedialog,messagebox
import pandas as pd
import pdfplumber
import traceback
import re
import os
from collections import defaultdict

def selecionar_arquivo_pdf():
    """Abre uma janela para o usuário selecionar um arquivo PDF."""
    root = tk.Tk()
    root.withdraw()
    filepath = filedialog.askopenfilename(
        title="Selecione o extrato PDF do Banestes",
        filetypes=[("Arquivos PDF", "*.pdf")]
    )
    if not filepath:
        print("Nenhum arquivo selecionado. Encerrando.")
        exit()
    return filepath

def extrair_dados_do_pdf(caminho_pdf):
    """
    Processa o extrato PDF analisando a posição de cada palavra na página.
    """
    print(f"Iniciando extração avançada por coordenadas: {caminho_pdf}")
    
    # --- Parâmetros de Layout ---
    COLUNA_DATA_FIM_X = 75
    COLUNA_VALOR_INICIO_X = 480
    # ---------------------------

    transacoes = []
    
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            dia_atual = ""
            for page in pdf.pages:
                words = page.extract_words(x_tolerance=2, y_tolerance=2, keep_blank_chars=True)
                linhas_agrupadas = defaultdict(list)
                for word in words:
                    chave_y = round(word['top'], 0) 
                    linhas_agrupadas[chave_y].append(word)

                for y in sorted(linhas_agrupadas.keys()):
                    palavras_na_linha = sorted(linhas_agrupadas[y], key=lambda w: w['x0'])
                    
                    col_data_str, col_desc_str, col_valor_str = "", "", ""

                    for palavra in palavras_na_linha:
                        if palavra['x0'] < COLUNA_DATA_FIM_X:
                            col_data_str += palavra['text']
                        elif palavra['x0'] > COLUNA_VALOR_INICIO_X:
                            col_valor_str += palavra['text']
                        else:
                            col_desc_str += palavra['text'] + " "

                    col_data_str = col_data_str.strip()
                    col_desc_str = col_desc_str.strip()
                    col_valor_str = col_valor_str.strip()

                    if re.match(r'^\d{2}$', col_data_str):
                        dia_atual = col_data_str

                    if col_desc_str and col_valor_str and re.search(r'[\d]', col_valor_str):
                        if "lançamento" in col_desc_str.lower():
                            continue

                        # Limpa o valor e converte para float (ex: 4750.00)
                        valor_numerico = float(re.sub(r'[^\d,-]', '', col_valor_str).replace('.', '').replace(',', '.'))
                        
                        palavras_debito = ['Pix Enviado', 'Pagamento', 'Tarifa', 'Cesta']
                        if any(keyword in col_desc_str for keyword in palavras_debito) and valor_numerico > 0:
                            valor_numerico *= -1
                        
                        transacoes.append({
                            "Data": f"{dia_atual}/JUN/25",
                            "Lançamento": col_desc_str,
                            "Valor (R$)": valor_numerico
                        })
    except Exception as e:
        print(f"Ocorreu um erro inesperado durante o processamento: {e}")
        return None

    if not transacoes:
        print("Nenhuma transação foi extraída com a análise de coordenadas.")
        return None

    df = pd.DataFrame(transacoes)
    return df
def iniciar_processamento():
    """
    Função de entrada que orquestra o fluxo de conversão.
    """
    pdf_path = filedialog.askopenfilename(
        title="Selecione o extrato do Banestes",
        filetypes=[("PDF files", "*.pdf")]
    )
    if not pdf_path:
        return False

    try:
        df = extrair_dados_do_pdf(pdf_path)
        if df.empty:
            messagebox.showwarning("Aviso", "Nenhuma transação válida foi encontrada no arquivo.")
            return False

        output_csv_path = os.path.splitext(pdf_path)[0] + ".csv"
        df.to_csv(output_csv_path, index=False, sep=';', encoding='utf-8-sig', decimal=',')
        return True
    except Exception as e:
        traceback.print_exc()
        messagebox.showerror("Erro Crítico", f"Ocorreu um erro inesperado ao processar o extrato.\n\n{e}")
        return False

def main():
    caminho_pdf = selecionar_arquivo_pdf()
    if not caminho_pdf: return

    df_resultado = extrair_dados_do_pdf(caminho_pdf)
    
    if df_resultado is not None and not df_resultado.empty:
        print("\n### Extração por Coordenadas Concluída com Sucesso! ###")
        # Define o formato de exibição do float com 2 casas decimais no print
        pd.options.display.float_format = '{:.2f}'.format
        print(df_resultado.to_string())
        
        # Define o caminho do CSV para a mesma pasta do PDF original.
        diretorio_pdf, nome_arquivo = os.path.split(caminho_pdf)
        nome_base = os.path.splitext(nome_arquivo)[0]
        caminho_csv = os.path.join(diretorio_pdf, f"{nome_base}_extraido.csv")
        
        try:
            # Salva o arquivo CSV no caminho definido.
            # O formato do float será padrão (com ponto decimal).
            df_resultado.to_csv(caminho_csv, index=False, encoding='utf-8-sig', sep=';', float_format='%.2f')
            print(f"\nOs dados foram salvos com sucesso em:\n{caminho_csv}")
        except Exception as e:
            print(f"\nErro ao salvar o arquivo CSV: {e}")

if __name__ == "__main__":
    main()