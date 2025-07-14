import pdfplumber
import pandas as pd

def extrair_tabelas_pdf(pdf_path):
    # Lista para armazenar DataFrames de cada página
    tabelas = []
    with pdfplumber.open(pdf_path) as pdf:
        for i, pagina in enumerate(pdf.pages):
            tabelas_pag = pagina.extract_tables()
            for tabela in tabelas_pag:
                # Converte a tabela em DataFrame, remove linhas vazias
                df = pd.DataFrame(tabela)
                if not df.empty:
                    tabelas.append(df)
    return tabelas

def salvar_csv(tabelas, csv_path):
    # Se houver mais de uma tabela, concatena. Caso contrário, salva só uma.
    if len(tabelas) > 1:
        df_final = pd.concat(tabelas, ignore_index=True)
    else:
        df_final = tabelas[0]
    df_final.to_csv(csv_path, index=False, header=False)

if __name__ == "__main__":
    caminho_pdf = "conversor/stone.pdf"     # Caminho do PDF de entrada
    caminho_csv = "extrato_stone.csv"     # Caminho do CSV de saída

    tabelas = extrair_tabelas_pdf(caminho_pdf)
    if tabelas:
        salvar_csv(tabelas, caminho_csv)
        print(f"Arquivo CSV gerado com sucesso: {caminho_csv}")
    else:
        print("Nenhuma tabela encontrada no PDF.")