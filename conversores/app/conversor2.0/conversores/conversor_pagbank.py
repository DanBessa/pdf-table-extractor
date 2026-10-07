import pdfplumber
import pandas as pd
import re
import os
import tkinter as tk
from tkinter import filedialog, messagebox

def clean_valor(valor_str):
    """
    Limpa e converte uma string de moeda (ex: 'R$ 1.234,56') para um número float.
    """
    if not valor_str or not isinstance(valor_str, str):
        return 0.0
    # Remove 'R$', espaços, e pontos de milhar. Troca vírgula por ponto decimal.
    valor_limpo = valor_str.replace('R$', '').strip().replace('.', '').replace(',', '.')
    try:
        return float(valor_limpo)
    except (ValueError, TypeError):
        return 0.0

def selecionar_pdfs():
    """
    Abre uma janela para o usuário selecionar um ou mais arquivos PDF do extrato.
    Retorna uma tupla com os caminhos dos arquivos selecionados.
    """
    root = tk.Tk()
    root.withdraw() 
    pdf_paths = filedialog.askopenfilenames(
        title="Selecione os extratos PROTEGE Cash",
        filetypes=[("Arquivos PDF", "*.pdf")]
    )
    return pdf_paths

def extrair_dados_pdf(pdf_path):
    """
    Extrai transações de um único arquivo PDF da PROTEGE Cash e o salva como CSV.
    """
    transacoes = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            # Junta o texto de todas as páginas em uma única string
            texto_completo = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])

        # Processa a string de texto linha por linha
        for linha in texto_completo.split('\n'):
            # Procura por linhas que começam com uma data, indicando uma transação
            if re.match(r'^\d{2}/\d{2}/\d{4}', linha):
                data = linha.split()[0]
                historico = "Não Identificado"
                valor = 0.0

                # Encontra todos os valores monetários (ex: R$ 100,00) na linha
                valores_monetarios = re.findall(r'R\$\s?[\d.,]+', linha)
                
                # A lógica é que o valor da transação é o penúltimo valor monetário na linha
                # (o último é sempre o saldo).
                if len(valores_monetarios) >= 2:
                    valor_transacao_str = valores_monetarios[-2]
                    
                    # Determina se o valor é débito (negativo) ou crédito (positivo)
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
            print(f"Nenhuma transação encontrada no arquivo: {pdf_path}")
            return

        df = pd.DataFrame(transacoes)
        # Formata a coluna 'Valor' para ter duas casas decimais
        df['Valor'] = df['Valor'].map('{:,.2f}'.format)
        
        caminho_csv = os.path.splitext(pdf_path)[0] + ".csv"
        df.to_csv(caminho_csv, index=False, sep=";", encoding="utf-8-sig")

        print(f"Arquivo salvo em: {caminho_csv}")

    except Exception as e:
        messagebox.showerror("Erro ao Processar Arquivo", f"Não foi possível processar o arquivo:\n{pdf_path}\n\nErro: {e}")

# --- Bloco de Execução Principal ---
if __name__ == "__main__":
    # 1. Chama a função para o usuário selecionar os arquivos PDF.
    caminhos_pdf = selecionar_pdfs()

    # 2. Verifica se algum arquivo foi selecionado.
    if not caminhos_pdf:
        messagebox.showwarning("Aviso", "Nenhum arquivo PDF foi selecionado.")
    else:
        # 3. Processa cada arquivo selecionado.
        for caminho_do_arquivo in caminhos_pdf:
            extrair_dados_pdf(caminho_do_arquivo)
        
        # 4. Mostra uma mensagem de sucesso única no final.
        messagebox.showinfo("Sucesso", f"{len(caminhos_pdf)} arquivo(s) processado(s) com sucesso!")