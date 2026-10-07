import os
import re
import pandas as pd
import pdfplumber
from tkinter import filedialog

def limpar_valor(valor_str):
    """Remove a formatação monetária e converte para float numérico."""
    if not valor_str:
        return 0.0
    v = str(valor_str).strip()
    eh_negativo = "-" in v or v.endswith("D") or v.endswith("-")
    
    v_limpo = re.sub(r"[^\d,\.]", "", v)
    if "," in v_limpo and "." in v_limpo:
        v_limpo = v_limpo.replace(".", "").replace(",", ".")
    elif "," in v_limpo:
        v_limpo = v_limpo.replace(",", ".")
    
    try:
        val = float(v_limpo)
        return -abs(val) if eh_negativo else val
    except (ValueError, TypeError):
        return 0.0

def processar_pdf_mod3(pdf_path):
    """Extrai transações do extrato Santander Modelo 2 utilizando pdfplumber."""
    try:
        full_text = ""
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                texto_pagina = page.extract_text(layout=True) or page.extract_text() or ""
                if texto_pagina:
                    full_text += texto_pagina + "\n"

        if not full_text:
            raise UserWarning(f"Não foi possível extrair texto do arquivo:\n{os.path.basename(pdf_path)}")

        transacoes = []
        transacao_pendente = None

        padrao_data_inicio = re.compile(r"^(\d{2}/\d{2}/\d{4})\s+(.*)")
        padrao_monetario = re.compile(r"-?[\d\.]*,\d{2}")
        palavras_negativas = [
            "boleto", "pix enviado", "tarifa", "tributo", "pagamento", "conta",
            "devolvido", "cancelado", "estorno", "darf", "ted enviada",
            "iof", "juros", "fgts", "saque", "débito", "debito"
        ]
        
        ignorar = [
            "SALDO ANTERIOR", "SALDO FINAL", "SALDO DO DIA", "SALDO TOTAL",
            "SALDO DISPON", "TOTAL", "S A L D O"
        ]

        linhas = full_text.split('\n')

        for linha in linhas:
            linha = linha.strip()
            if not linha or any(term in linha.upper() for term in ignorar):
                continue

            match_data = padrao_data_inicio.match(linha)

            if match_data:
                if transacao_pendente:
                    transacao_pendente = None

                data_str = match_data.group(1)
                resto_linha = match_data.group(2).strip()

                valores_encontrados = padrao_monetario.findall(resto_linha)

                if len(valores_encontrados) > 0:
                    valor_str = valores_encontrados[-2] if len(valores_encontrados) >= 2 else valores_encontrados[-1]
                    posicao_valor = resto_linha.rfind(valor_str)
                    desc = resto_linha[:posicao_valor].strip()
                    desc = re.sub(r'\s+', ' ', desc.replace('|', ' ')).strip()

                    valor = limpar_valor(valor_str)
                    is_negativo = any(palavra in desc.lower() for palavra in palavras_negativas)
                    if is_negativo and valor > 0:
                        valor = -valor

                    tipo = "D" if valor < 0 else "C"
                    if valor != 0.0:
                        transacoes.append({
                            "Data": data_str,
                            "Histórico": desc if desc else "LANCAMENTO SANTANDER",
                            "Valor": valor,
                            "Tipo": tipo
                        })
                else:
                    transacao_pendente = {"Data": data_str, "Histórico": resto_linha}

            elif transacao_pendente:
                valores_encontrados = padrao_monetario.findall(linha)

                if len(valores_encontrados) > 0:
                    valor_str = valores_encontrados[0]
                    posicao_valor = linha.rfind(valor_str)
                    desc_complemento = linha[:posicao_valor].strip()

                    desc_final = f"{transacao_pendente['Histórico']} {desc_complemento}".strip()
                    desc_final = re.sub(r'\s+', ' ', desc_final.replace('|', ' ')).strip()
                    
                    valor = limpar_valor(valor_str)
                    is_negativo = any(palavra in desc_final.lower() for palavra in palavras_negativas)
                    if is_negativo and valor > 0:
                        valor = -valor

                    tipo = "D" if valor < 0 else "C"
                    if valor != 0.0:
                        transacoes.append({
                            "Data": transacao_pendente['Data'],
                            "Histórico": desc_final if desc_final else "LANCAMENTO SANTANDER",
                            "Valor": valor,
                            "Tipo": tipo
                        })
                    transacao_pendente = None

        if not transacoes:
            raise UserWarning(f"Nenhuma transação encontrada no formato esperado em:\n{os.path.basename(pdf_path)}")

        # Preserva todas as transações repetidas
        df = pd.DataFrame(transacoes)[["Data", "Histórico", "Valor", "Tipo"]]
        
        caminho_saida = os.path.splitext(pdf_path)[0] + ".xlsx"
        df.to_excel(caminho_saida, index=False)
        return caminho_saida

    except Exception as e:
        raise Exception(f"Erro ao processar o arquivo {os.path.basename(pdf_path)}:\n{e}")

def iniciar_processamento(caminho_pdf=None):
    """Função de entrada chamada pelo integrador principal."""
    if not caminho_pdf:
        caminho_pdf = filedialog.askopenfilename(
            title="Selecione o extrato Santander (Modelo 2)",
            filetypes=[("Arquivos PDF", "*.pdf")]
        )
    
    if not caminho_pdf:
        raise UserWarning("Nenhum arquivo selecionado.")

    if isinstance(caminho_pdf, (list, tuple)):
        caminho_pdf = caminho_pdf[0]

    return processar_pdf_mod3(caminho_pdf)

def main(caminho_pdf=None):
    return iniciar_processamento(caminho_pdf)

class PDFTableExtractor:
    def __init__(self, file_path, configs=None):
        self.file_path = file_path

    def start(self):
        return iniciar_processamento(self.file_path)