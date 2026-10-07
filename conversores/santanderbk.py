# conversor_santandermod1.py (CORRIGIDO)
import pdfplumber
import re
import pandas as pd
from tkinter import filedialog # Mantém apenas o filedialog
import os
# REMOVIDO: import tkinter as tk, messagebox

def extrair_dados(linha, data_corrente):
    match_valor = re.search(r"(\d{1,3}(?:\.\d{3})*,\d{2}-?)", linha)
    if not match_valor:
        return None

    valor_raw = match_valor.group(1)
    valor_index = linha.rfind(valor_raw)
    lancamento = linha[:valor_index].strip()

    doc_match = re.search(r"(\d{6,})(?:\s+|\s*-\s*)?" + re.escape(valor_raw), linha)
    documento = doc_match.group(1) if doc_match else ""

    historico_minusculo = lancamento.lower()
    palavras_negativas = ["boleto", "outros bancos", "aplicacao", "pix enviado", "transferência enviada","tarifa","comercial",
                          "tributo","estadual","esgoto","telefone","devolvido","cancelado","estorno","distribuidora","fornecedores",
                          "darf","celular","salario","bananas","ted enviada","pagsal","conta luz","agua",]
    valor_final_str = "" 

    for palavra in palavras_negativas:
        if palavra in historico_minusculo:
            valor_final_str = "-" + valor_raw.replace("-", "").rstrip("-")
            break
    else:
        tem_hifen = valor_raw.endswith("-")
        valor_final_str = "-" + valor_raw[:-1] if tem_hifen else valor_raw
    
    # Formata para o padrão do app principal
    valor_formatado_csv = valor_final_str.replace('.', '').replace(',', ',')
    
    return [data_corrente, lancamento, valor_formatado_csv, documento]

def preparar_linha(linhas, idx):
    linha = linhas[idx].strip().replace('\t', ' ')
    linhas_usadas = 1
    
    data_inicio_regex_lookahead = re.compile(r"^(\d{2}/\d{2}(?:/\d{2,4})?)\b")
    for offset in range(1, 3):
        if idx + offset < len(linhas):
            extra = linhas[idx + offset].strip().replace('\t', ' ')
            if not re.search(r"\d{1,3}(?:\.\d{3})*,\d{2}-?", linha) and \
               not data_inicio_regex_lookahead.match(extra) and \
               extra:
                linha += " " + extra
                linhas_usadas += 1
            else:
                break
        else:
            break

    linha = re.sub(r"(\d{6,})(\d{1,3}(?:\.\d{3})*,\d{2}-?)", r"\1 \2", linha)
    return linha, linhas_usadas

def processar_pdf(pdf_path):
    try:
        reader = pdfplumber(pdf_path)
        data = []
        current_date = ""
        start_extract = False
        data_inicio_regex = re.compile(r"^(\d{2}/\d{2}(?:/\d{2,4})?)\b")
        fim_conteudo = "EXTRATO CONSOLIDADO"

        for i, page in enumerate(reader.pages):
            texto = page.extract_text()
            if not texto:
                continue

            linhas = texto.split('\n')
            idx = 0
            while idx < len(linhas):
                linha_base = linhas[idx].strip()

                if "Movimentação" in linha_base:
                    start_extract = True
                    for skip_idx in range(idx + 1, min(idx + 4, len(linhas))):
                        if re.match(r"^\s*SALDO (ANTERIOR|EM \d{2}/\d{2}/\d{4})", linhas[skip_idx].strip().upper()):
                            idx = skip_idx + 1
                            break
                        if data_inicio_regex.match(linhas[skip_idx].strip()):
                            idx = skip_idx
                            break
                    else:
                        idx += 2
                    continue
                
                if not start_extract or (fim_conteudo in linha_base and not data_inicio_regex.match(linha_base)):
                    idx += 1
                    continue

                linha_completa, usadas = preparar_linha(linhas, idx)

                match_data = data_inicio_regex.match(linha_completa)
                if match_data:
                    current_date = match_data.group(1)
                    linha_completa = data_inicio_regex.sub('', linha_completa, 1).strip()

                if current_date:
                    entrada = extrair_dados(linha_completa, current_date)
                    if entrada:
                        data.append(entrada)

                idx += usadas

        if not data:
            raise UserWarning(f"Nenhuma transação encontrada ou extraída em:\n{os.path.basename(pdf_path)}")

        df = pd.DataFrame(data, columns=["Data", "Lançamento", "Valor", "Documento"])
        
        # Limpa o DataFrame
        df.drop_duplicates(inplace=True)
        df = df[~df['Lançamento'].str.contains("SALDO ANTERIOR", case=False, na=False)]
        df = df[~df['Lançamento'].str.match(r"^\s*SALDO EM \d{2}/\d{2}(?:/\d{2,4})?\s*$", case=False, na=False)]

        if df.empty:
            raise UserWarning(f"Nenhuma transação válida após limpeza em:\n{os.path.basename(pdf_path)}")

        csv_path = os.path.splitext(pdf_path)[0] + ".csv"
        # Salva o CSV com o valor já formatado como string
        df.to_csv(csv_path, index=False, sep=";", encoding="utf-8-sig")
        
        return csv_path

    except Exception as e:
        raise Exception(f"Erro ao processar o arquivo {os.path.basename(pdf_path)}:\n{e}")

def iniciar_processamento():
    """Função principal para orquestrar a seleção e processamento."""
    
    # Esta chamada será interceptada pelo patch
    caminho_pdf = filedialog.askopenfilename(
        title="Selecione o PDF do extrato Santander (Modo 1)",
        filetypes=[("Arquivos PDF", "*.pdf")]
    )
    
    if not caminho_pdf:
        raise UserWarning("Nenhum arquivo selecionado.")

    resultado_path = processar_pdf(caminho_pdf)
        
    if resultado_path:
        return resultado_path # Retorna o caminho do CSV salvo
    else:
        # Se processar_pdf falhou e não lançou exceção, lança agora
        raise Exception("Falha no processamento do PDF.")

# REMOVIDO: if __name__ == "__main__":