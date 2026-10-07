import pdfplumber
import re
import csv
import io
import os
# from tkinter import messagebox
from pdfminer.pdfparser import PDFSyntaxError 

# --- REGEX GLOBAIS ---
REGEX_PERIODO = re.compile(r"lançamentos período:\s*(\d{2}/(\d{2}/\d{4}))")
REGEX_DATA_EXTRACT = re.compile(r"(\d{2}\s*/\s*set)", re.IGNORECASE) # Nota: 'set' está fixo, pode ser um problema futuro
REGEX_VALOR_FIM = re.compile(r"\s(-?[\d\.]*,\d{2})$")
STOP_SIGNAL = "saldo da conta corrente"

# --- Funções Auxiliares ---

def limpar_valor_para_float(valor_str):
    """ Converte '1.234,56' para 1234.56 """
    valor_limpo = valor_str.replace(".", "").replace(",", ".")
    return float(valor_limpo)

def formatar_valor_para_csv(valor_float):
    """ Converte 1234.56 para '1234,56' (padrão CSV brasileiro) """
    return f"{valor_float:.2f}".replace('.', ',')

def extrair_periodo_pdf(pdf_object): 
    """ Vasculha a primeira página para encontrar Mês e Ano. """
    try:
        pagina_1 = pdf_object.pages[0]
        texto = pagina_1.extract_text(x_tolerance=1, y_tolerance=1, layout=True) 
        
        if not texto:
            # MUDANÇA: Substitui messagebox por raise
            raise UserWarning("Nenhum texto encontrado na página 1 do PDF.")

        match = REGEX_PERIODO.search(texto)
        
        if match:
            mes, ano = match.group(2).split('/')
            return mes, ano
        else:
            # MUDANÇA: Substitui messagebox por raise
            raise Exception("Erro de Período: Não foi possível encontrar a string 'lançamentos período: dd/MM/AAAA' no PDF.")
            
    except Exception as e:
        # MUDANÇA: Relança a exceção
        raise Exception(f"Erro ao tentar ler o período: {e}")

def processar_texto_pagina(texto, mes, ano, data_anterior, processando_lancamentos_ativo):
    """ Processa o texto puro de uma página, linha por linha (v11). """
    lancamentos_pagina = []
    
    for linha in texto.split('\n'):
        linha_limpa = linha.strip()
        linha_lower = linha_limpa.lower()

        if STOP_SIGNAL in linha_lower:
            processando_lancamentos_ativo = False
        
        if not processando_lancamentos_ativo:
            continue

        filtro_cabecalho_saldo = "saldo disponível" in linha_lower or "limite da conta" in linha_lower
        filtro_geral = "lançamentos" in linha_lower or "ag/origem" in linha_lower
        filtro_saldo = "SALDO ANTERIOR" in linha_limpa.upper() or "SALDO TOTAL" in linha_limpa.upper()

        if filtro_cabecalho_saldo or filtro_geral or filtro_saldo:
            continue

        match_valor = REGEX_VALOR_FIM.search(linha_limpa)
        
        if match_valor:
            valor_str = match_valor.group(1)
            valor = limpar_valor_para_float(valor_str)
            desc_com_data = linha_limpa[:match_valor.start()].strip()
            match_data = REGEX_DATA_EXTRACT.search(desc_com_data)
            
            data_atual = data_anterior
            descricao = ""
            
            if match_data:
                data_bruta = match_data.group(1)
                dia = re.sub(r'[^0-9]', '', data_bruta.split('/')[0]) # Limpa espaços
                data_atual = f"{dia}/{mes}" # Formato dd/MM
                data_anterior = data_atual
                
                descricao = REGEX_DATA_EXTRACT.sub("", desc_com_data).strip()
                descricao = re.sub(r"^\s*,\s*", "", descricao)
                descricao = re.sub(r"\s+", " ", descricao).strip()
            else:
                data_atual = data_anterior
                descricao = desc_com_data
            
            if descricao:
                lancamentos_pagina.append({
                    "Data": data_atual,
                    "Descricao": descricao,
                    "Valor": valor
                })
                
    return lancamentos_pagina, data_anterior, processando_lancamentos_ativo


# --- Função Principal (Para ser chamada pelo menu) ---

def iniciar_processamento(pdf_path):
    """
    Função principal que o menu 'menuestilizado.py' irá chamar.
    Recebe o pdf_path diretamente do menu (1 argumento).
    """
    
    if not pdf_path:
        # MUDANÇA: Substitui messagebox por raise
        raise UserWarning("Nenhum caminho de PDF foi recebido pelo módulo.")

    arquivo_csv_path = os.path.splitext(pdf_path)[0] + ".csv"
    
    lancamentos_completos = []
    data_anterior_global = None
    processando_lancamentos_global = True 

    try:
        with pdfplumber.open(pdf_path) as pdf:
            
            mes_extrato, ano_extrato = extrair_periodo_pdf(pdf)
            
            if not mes_extrato or not ano_extrato:
                # O erro já foi lançado por extrair_periodo_pdf
                return None

            for pagina in pdf.pages:
                texto_pagina = pagina.extract_text(layout=True) 
                
                if not texto_pagina:
                    continue
                
                novos_lancamentos, data_anterior_global, processando_lancamentos_global = processar_texto_pagina(
                    texto_pagina, 
                    mes_extrato, 
                    ano_extrato, 
                    data_anterior_global,
                    processando_lancamentos_global
                )
                
                lancamentos_completos.extend(novos_lancamentos)

                if not processando_lancamentos_global:
                    break 
        
        if lancamentos_completos:
            with io.open(arquivo_csv_path, 'w', encoding='utf-8', newline='') as f:
                fieldnames_csv = ["Data", "Historico", "Valor"]
                writer = csv.DictWriter(f, fieldnames=fieldnames_csv, delimiter=';')
                writer.writeheader()
                
                for lancamento in lancamentos_completos:
                    valor_formatado = formatar_valor_para_csv(lancamento['Valor'])
                    writer.writerow({
                        "Data": lancamento['Data'],
                        "Historico": lancamento['Descricao'],
                        "Valor": valor_formatado
                    })
            
            # Retorna o caminho do CSV para o app principal
            return arquivo_csv_path
        
        else:
            # MUDANÇA: Substitui messagebox por raise
            raise UserWarning(f"Nenhum lançamento encontrado em '{os.path.basename(pdf_path)}'.")

    except PDFSyntaxError:
        # MUDANÇA: Substitui messagebox por raise
        raise Exception(f"Erro de PDF: O arquivo '{os.path.basename(pdf_path)}' parece estar corrompido ou não é um PDF válido.")
    except PermissionError:
         # MUDANÇA: Substitui messagebox por raise
         raise Exception(f"Erro de Permissão: Não foi possível salvar o arquivo:\n{arquivo_csv_path}\n\nVerifique se o arquivo não está aberto em outro programa.")
    except Exception as e:
        # MUDANÇA: Substitui messagebox por raise
        raise Exception(f"Ocorreu um erro durante a extração:\n{e}")