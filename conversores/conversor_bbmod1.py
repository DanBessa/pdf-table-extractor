import os
import re
import traceback
from typing import Optional, List
import pandas as pd
import pdfplumber
from tkinter import filedialog

def extrair_formato_cac(caminho_pdf: str) -> Optional[pd.DataFrame]:
    transacoes: List[dict] = []
    
    # Captura números com ponto ou vírgula seguidos de parênteses: ex "13,33 (+)", "330,77 (-)", "61.48 (+)"
    padrao_valor = re.compile(r'(\d{1,3}(?:[.,]\d{3})*[.,]\d{2}|\d+[.,]\d{2})\s*\(\s*([^)]+)\s*\)')
    padrao_data = re.compile(r'\b(\d{2}/\d{2}/\d{4})\b')

    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            texto_completo = ""
            for pagina in pdf.pages:
                txt = pagina.extract_text(x_tolerance=2, y_tolerance=2) or ""
                texto_completo += txt + "\n"

            # Normaliza quebras acidentais de parênteses entre linhas (ex: "(+\n)")
            texto_completo = re.sub(r'\(\s*([^\)\n]+)\s*\n\s*\)', r'(\1)', texto_completo)

            data_atual = None
            buffer_desc = []

            for linha in texto_completo.split('\n'):
                l = linha.strip().replace('|', ' ')
                if not l:
                    continue

                l_upper = l.upper()
                if any(t in l_upper for t in [
                    "EXTRATO DE CONTA CORRENTE", "CLIENTE CAC", "AGÊNCIA:", "AGENCIA:", 
                    "DIA LOTE DOCUMENTO", "INFORMAÇÕES ADICIONAIS", "INFORMACOES ADICIONAIS",
                    "INFORMAÇÕES COMPLEMENTARES", "TOTAL APLICAÇÕES", "LIMITE ESPECIAL", 
                    "TAXA CHEQUE", "CUSTO EFETIVO", "SUJEITOS A CONFIRMAÇÃO"
                ]):
                    continue

                # Identifica data (ignora a linha técnica 00/00/0000)
                m_d = padrao_data.search(l)
                if m_d and m_d.group(1) != "00/00/0000":
                    data_atual = m_d.group(1)

                # Verifica se a linha tem valor com sinal entre parênteses
                m_v = padrao_valor.search(l)
                if m_v:
                    val_str = m_v.group(1)
                    sinal_str = m_v.group(2).strip()

                    # Descarta linhas de saldo acumulado / saldos diários
                    if any(s in l_upper for s in ["SALDO ANTERIOR", "SALDO DO DIA", "SALDO FINAL"]):
                        buffer_desc = []
                        continue
                    if re.match(r'^\s*SALDO\b', l_upper) and ('0,00' in val_str or '0.00' in val_str):
                        buffer_desc = []
                        continue

                    # Conversão numérica flexível (vírgula ou ponto)
                    v_limpo = val_str
                    if '.' in v_limpo and ',' in v_limpo:
                        v_limpo = v_limpo.replace('.', '').replace(',', '.')
                    elif ',' in v_limpo:
                        v_limpo = v_limpo.replace(',', '.')

                    try:
                        num = abs(float(v_limpo))
                    except ValueError:
                        num = 0.0

                    if num == 0.0:
                        buffer_desc = []
                        continue

                    # REGRA DE SINAL:
                    # Se tem '+', é Crédito (positivo, 'C').
                    # Se NÃO tem '+' (ou tem '-', '−', '–', '—'), é DÉBITO (negativo, 'D').
                    if '+' in sinal_str:
                        valor_final = num
                        tipo_final = "C"
                    else:
                        valor_final = -num
                        tipo_final = "D"

                    # Limpa a descrição da linha atual
                    l_sem_val = l.replace(m_v.group(0), "")
                    if m_d:
                        l_sem_val = l_sem_val.replace(m_d.group(1), "")
                    l_sem_val = l_sem_val.replace("00/00/0000", "").strip()
                    l_sem_val = re.sub(r'^\s*\d{4,6}\s+[\d\w]{4,20}\s*', '', l_sem_val)
                    l_sem_val = re.sub(r'^\s*\d{4,6}\s*', '', l_sem_val)

                    partes = buffer_desc + ([l_sem_val] if l_sem_val else [])
                    desc_completa = " ".join(partes).strip()
                    buffer_desc = []

                    desc_completa = re.sub(r'^\s*\d{4,6}\s+[\d\w]{4,20}\s*', '', desc_completa)
                    desc_completa = re.sub(r'^\s*\d{4,6}\s*', '', desc_completa)
                    desc_completa = re.sub(r'\s+', ' ', desc_completa).strip()

                    if not desc_completa:
                        desc_completa = "LANCAMENTO BANCARIO"

                    if data_atual:
                        transacoes.append({
                            "Data": data_atual,
                            "Histórico": desc_completa,
                            "Valor": valor_final,  # float negativo para saídas
                            "Tipo": tipo_final      # 'D' ou 'C'
                        })
                else:
                    # Linha sem valor -> acumula no buffer
                    l_desc = l
                    if m_d:
                        l_desc = l_desc.replace(m_d.group(1), "")
                    l_desc = l_desc.replace("00/00/0000", "").strip()
                    l_desc = re.sub(r'^\s*\d{4,6}\s+[\d\w]{4,20}\s*', '', l_desc)
                    l_desc = re.sub(r'^\s*\d{4,6}\s*', '', l_desc)
                    l_desc = re.sub(r'\s+', ' ', l_desc).strip()

                    if l_desc and not any(ign in l_desc.upper() for ign in ["SALDO DO DIA", "SALDO ANTERIOR", "00/00/0000"]):
                        buffer_desc.append(l_desc)

        if not transacoes:
            return None

        # Gera com as 4 colunas exigidas pelo menu unificado
        return pd.DataFrame(transacoes)[["Data", "Histórico", "Valor", "Tipo"]]

    except Exception as e:
        print(f"Erro ao processar PDF do BB: {e}")
        traceback.print_exc()
        raise e

def iniciar_processamento(pdf_path=None):
    if not pdf_path:
        pdf_path = filedialog.askopenfilename(
            title="Selecione o extrato (BB - Modelo 1)",
            filetypes=[("Arquivos PDF", "*.pdf")]
        )
    if not pdf_path:
        raise UserWarning("Nenhum arquivo selecionado.")

    try:
        df_transacoes = extrair_formato_cac(pdf_path)
        
        if df_transacoes is None or df_transacoes.empty:
            raise UserWarning(f"Nenhuma transação válida foi encontrada em {os.path.basename(pdf_path)}.")

        total_cred = len(df_transacoes[df_transacoes["Tipo"] == "C"])
        total_deb = len(df_transacoes[df_transacoes["Tipo"] == "D"])
        print(f"✅ BB Modelo 1: {len(df_transacoes)} lançamentos extraídos com sucesso.")
        print(f"   🟢 Entradas (+ / C): {total_cred}")
        print(f"   🔴 Saídas (- / D): {total_deb}")

        caminho_saida = os.path.splitext(pdf_path)[0] + ".xlsx"
        df_transacoes.to_excel(caminho_saida, index=False)
        return caminho_saida
        
    except Exception as e:
        traceback.print_exc()
        raise Exception(f"Erro ao processar '{os.path.basename(pdf_path)}':\n{e}")

if __name__ == "__main__":
    iniciar_processamento()