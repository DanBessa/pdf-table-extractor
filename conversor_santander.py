import re
import pandas as pd
from PyPDF2 import PdfReader
import tkinter as tk
from tkinter import filedialog, messagebox
import os

def selecionar_pdf():
    # Esta função já existe no seu código e está correta para selecionar múltiplos arquivos.
    # Ela será chamada por iniciar_extracao_santander().
    root = tk.Tk()
    root.withdraw()
    return filedialog.askopenfilenames(
        title="Selecione os PDFs do extrato Santander",
        filetypes=[("Arquivos PDF", "*.pdf")]
    )

def extrair_dados(linha, data_corrente):
    # Sua função extrair_dados (sem alterações nesta lógica)
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
                          "tributo","estadual","esgoto","telefone","devolvido","cancelado","estorno","distribuidora","fornecedores","darf","celular"]
    valor_final_str = "" # Variável para armazenar o valor formatado como string

    for palavra in palavras_negativas:
        if palavra in historico_minusculo:
            valor_final_str = "-" + valor_raw.replace("-", "").rstrip("-")
            break
    else:
        tem_hifen = valor_raw.endswith("-")
        valor_final_str = "-" + valor_raw[:-1] if tem_hifen else valor_raw
    
    # Normaliza o valor para o formato numérico com ponto decimal para consistência, se necessário,
    # ou mantém como string se preferir (Pandas pode inferir ao criar o DataFrame).
    # Se for manter como string para o CSV, valor_final_str já está pronto.

    return [data_corrente, lancamento, valor_final_str, documento]

def preparar_linha(linhas, idx):
    # Sua função preparar_linha (sem alterações nesta lógica)
    linha = linhas[idx].strip().replace('\t', ' ')
    linhas_usadas = 1

    # Tenta concatenar até 2 linhas seguintes se a linha atual não tiver um valor monetário aparente
    # e a próxima linha não começar com uma data (indicando uma nova transação)
    data_inicio_regex_lookahead = re.compile(r"^(\d{2}/\d{2}(?:/\d{2,4})?)\b")
    for offset in range(1, 3): # Verifica as próximas 2 linhas
        if idx + offset < len(linhas):
            extra = linhas[idx + offset].strip().replace('\t', ' ')
            # Se a linha atual NÃO tem valor E a linha extra NÃO parece ser uma nova transação (não começa com data)
            if not re.search(r"\d{1,3}(?:\.\d{3})*,\d{2}-?", linha) and \
               not data_inicio_regex_lookahead.match(extra) and \
               extra: # E a linha extra não está vazia
                linha += " " + extra
                linhas_usadas += 1
            else:
                break # Para se a linha extra for uma nova transação ou a linha atual já tiver valor
        else:
            break # Não há mais linhas para verificar

    linha = re.sub(r"(\d{6,})(\d{1,3}(?:\.\d{3})*,\d{2}-?)", r"\1 \2", linha) # Adiciona espaço entre doc e valor
    return linha, linhas_usadas


def processar_pdf(pdf_path):
    # Sua função processar_pdf (sem alterações significativas, apenas garante que usa as outras funções)
    try:
        reader = PdfReader(pdf_path)
        data = []
        current_date = ""
        start_extract = False
        data_inicio_regex = re.compile(r"^(\d{2}/\d{2}(?:/\d{2,4})?)\b")
        fim_conteudo = "EXTRATO CONSOLIDADO" # Pode precisar de ajuste para ser mais específico

        for i, page in enumerate(reader.pages):
            texto = page.extract_text()
            if not texto:
                continue

            linhas = texto.split('\n')
            idx = 0
            while idx < len(linhas):
                linha_base = linhas[idx].strip()

                if "Movimentação" in linha_base: # Idealmente, um regex mais específico para o cabeçalho da tabela
                    start_extract = True
                    # Pular cabeçalhos da tabela de movimentação, se houver mais de uma linha
                    # Ex: Data, Descrição, Nº Doc, Débitos, Créditos, Saldo
                    # Este loop de pular linhas pode precisar de ajuste
                    for skip_idx in range(idx + 1, min(idx + 4, len(linhas))):
                        if re.match(r"^\s*SALDO (ANTERIOR|EM \d{2}/\d{2}/\d{4})", linhas[skip_idx].strip().upper()):
                            idx = skip_idx +1
                            break
                        if data_inicio_regex.match(linhas[skip_idx].strip()): # Se já encontrar uma data, não pular mais
                            idx = skip_idx
                            break
                    else: # Se não encontrou saldo anterior ou data, pula um número fixo de linhas de cabeçalho
                        idx += 2 # Ajuste conforme necessário
                    continue
                
                if not start_extract or (fim_conteudo in linha_base and not data_inicio_regex.match(linha_base)):
                    idx += 1
                    continue

                linha_completa, usadas = preparar_linha(linhas, idx)

                match_data = data_inicio_regex.match(linha_completa)
                if match_data:
                    current_date = match_data.group(1)
                    # Remove a data do início da linha completa para não interferir na extração do lançamento
                    linha_completa = data_inicio_regex.sub('', linha_completa, 1).strip()


                # Só tenta extrair dados se houver uma data corrente (evita processar lixo antes da primeira data)
                if current_date:
                    entrada = extrair_dados(linha_completa, current_date)
                    if entrada:
                        data.append(entrada)

                idx += usadas

        if not data:
            messagebox.showwarning("Aviso", f"Nenhuma transação encontrada ou extraída em:\n{os.path.basename(pdf_path)}")
            return # Retorna para não tentar criar CSV vazio

        df = pd.DataFrame(data, columns=["Data", "Lançamento", "Valor", "Documento"])
        # Converte a coluna 'Valor' para numérico, tratando erros
        def converter_valor_para_numerico(valor_str):
            if isinstance(valor_str, (int, float)):
                return valor_str
            s = valor_str.replace('.', '').replace(',', '.')
            try:
                return float(s)
            except ValueError:
                return None # Ou 0.0, ou manter como string se a conversão falhar

        df["Valor"] = df["Valor"].apply(converter_valor_para_numerico)
        df.drop_duplicates(inplace=True)

        # Remove linhas onde 'Lançamento' é apenas 'SALDO ANTERIOR' ou similar, se não desejado
        df = df[~df['Lançamento'].str.contains("SALDO ANTERIOR", case=False, na=False)]
        df = df[~df['Lançamento'].str.match(r"^\s*SALDO EM \d{2}/\d{2}(?:/\d{2,4})?\s*$", case=False, na=False)]


        if df.empty:
            messagebox.showwarning("Aviso", f"Nenhuma transação válida após limpeza em:\n{os.path.basename(pdf_path)}")
            return

        csv_path = os.path.splitext(pdf_path)[0] + ".csv"
        df.to_csv(csv_path, index=False, sep=";", decimal=",", encoding="utf-8-sig") # utf-8-sig para Excel ler acentos corretamente
        #messagebox.showinfo("Sucesso", f"CSV salvo com sucesso em:\n{csv_path}")

    except Exception as e:
        messagebox.showerror("Erro no Processamento", f"Erro ao processar o arquivo {os.path.basename(pdf_path)}:\n{e}")

def iniciar_extracao_santander(): # Função orquestradora que será chamada pela GUI
    caminhos_pdf = selecionar_pdf() # Chama a função de seleção de arquivos deste script
    
    if not caminhos_pdf:
        # O usuário cancelou a seleção de arquivos, pode-se mostrar uma mensagem ou apenas retornar.
        # messagebox.showwarning("Seleção Cancelada", "Nenhum PDF do Santander foi selecionado.")
        return

    for caminho_pdf in caminhos_pdf:
        processar_pdf(caminho_pdf)
    
    # Mensagem geral após tentar processar todos os arquivos (opcional)
    # if caminhos_pdf: # Apenas se arquivos foram selecionados inicialmente
    #      messagebox.showinfo("Concluído", "Processamento dos extratos Santander finalizado. Verifique as mensagens para cada arquivo.")


if __name__ == "__main__":
    # Este bloco permite que o script seja executado de forma independente para testes.
    iniciar_extracao_santander()