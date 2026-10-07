import os
import re
from datetime import datetime
import pandas as pd
import pdfplumber
from tkinter import filedialog
import traceback

ANO_FALLBACK = datetime.now().year

def limpar_valor_itau(val_str: str):
    if not val_str or pd.isna(val_str): return 0.0, "C"
    v = str(val_str).strip()
    eh_debito = any(c in v for c in ['-', '−', '–', '—']) or v.endswith("D") or v.endswith("d") or "SAÍDA" in v.upper()
    v_limpo = re.sub(r"[^\d,\.]", "", v)
    if "," in v_limpo and "." in v_limpo: v_limpo = v_limpo.replace(".", "").replace(",", ".")
    elif "," in v_limpo: v_limpo = v_limpo.replace(",", ".")
    try:
        val_float = float(v_limpo)
        if eh_debito and val_float > 0: val_float = -val_float
        elif not eh_debito and val_float < 0: val_float = abs(val_float)
        return val_float, "D" if val_float < 0 else "C"
    except ValueError:
        return 0.0, "C"

def limpar_historico_itau(texto: str) -> str:
    if not texto: return ""
    t = str(texto).replace('|', ' ').strip()
    termos_remover = [
        r"^Sispag\s+", r"^PIX ENVIADO\s+", r"^PIX RECEBIDO\s+", r"^BOLETO PAGO\s+",
        r"^PIX QRS?\s+", r"^PIX TRANSF\s+", r"^TED\s+[\d\.]+\s*", r"^DOC\s+[\d\.]+\s*",
        r"^TED\s+", r"^DOC\s+", r"^TAR\s+", r"^Tar\s+", r"^Tarifa\s+"
    ]
    for termo in termos_remover:
        t = re.sub(termo, "", t, flags=re.IGNORECASE).strip()
    t = re.sub(r'\b[A-Z]{2}\d{8,}\b', '', t)
    t = re.sub(r'(?<!\S)[-\/](?!\S)', '', t)
    return re.sub(r'\s+', ' ', t).strip()

def extrair_extrato_itau_consolidado(caminho_pdf: str) -> str:
    transacoes = []
    ignorar_termos = [
        "SALDO ANTERIOR", "SALDO ATUAL", "SALDO DISPON", "SALDO DO DIA", "SALDO EM CONTA",
        "EXTRATO DE CONTA CORRENTE", "EXTRATO MENSAL CONSOLIDADO", "BANCO ITAU", "BANCO ITAÚ",
        "ITAU UNIBANCO", "DATA LANÇAMENTOS", "DATA HISTÓRICO", "VALOR SALDO", "VALOR (R$)",
        "SALDO (R$)", "FOLHA", "PÁGINA", "PAGINA", "EXTRATO DE:", "(CRÉDITOS)", "(DÉBITOS)",
        "OUTRAS SAIDAS", "OUTRAS ENTRADAS", "CONTA CORRENTE E APLICAÇÕES", "ENTRADAS R$"
    ]
    
    legendas_pdf = [
        "A = agendamento", "ações movimentadas", "pela Bolsa de Valores", "crédito a compensar",
        "D = débito a compensar", "débito a compensar", "G = aplicação programada",
        "aplicação programada", "P = poupança automática", "poupança automática",
        "Para demais siglas, consulte as Notas", "demais siglas, consulte as Notas",
        "Explicativas no final do extrato"
    ]

    current_date = None
    ano_detectado = datetime.now().strftime("%Y")

    with pdfplumber.open(caminho_pdf) as pdf:
        texto_completo = ""
        for p in pdf.pages[:3]: texto_completo += (p.extract_text() or "") + "\n"
        m_ano = re.search(r'/\s*(202\d)\b', texto_completo) or re.search(r'\b(202\d)\b', texto_completo)
        if m_ano: ano_detectado = m_ano.group(1)

        em_movimentacao = True
        for page in pdf.pages:
            texto_pagina = page.extract_text(layout=True) or ""
            for linha in texto_pagina.split("\n"):
                linha_limpa = linha.strip()
                if not linha_limpa: continue
                    
                for leg in legendas_pdf:
                    linha_limpa = re.sub(re.escape(leg), "", linha_limpa, flags=re.IGNORECASE).strip()
                if not linha_limpa: continue

                l_upper = linha_limpa.upper()
                if any(x in l_upper for x in ["SALDO FINAL", "TOTALIZADOR", "CHEQUE ESPECIAL", "02. INVESTIMENTOS", "INDICADORES DE MERCADO", "NOTAS EXPLICATIVAS", "RESUMO MÊS", "MOVIMENTAÇÃO - APLICAÇÕES"]):
                    em_movimentacao = False
                if not em_movimentacao: continue

                if re.match(r'^\d{6}\s+[A-Z0-9]+\s+\d{2}/\d{2}/\d{4}', linha_limpa) or l_upper.startswith("TOTAL ") or l_upper == "TOTAL":
                    continue

                if any(term in l_upper for term in ignorar_termos):
                    m_d_ign = re.search(r"\b(\d{2}/\d{2}(?:/\d{4})?)\b", linha_limpa)
                    if m_d_ign: current_date = m_d_ign.group(1) if len(m_d_ign.group(1)) == 10 else f"{m_d_ign.group(1)}/{ano_detectado}"
                    continue

                m_data = re.search(r"^\b(\d{2}/\d{2}(?:/\d{4})?)\b", linha_limpa)
                if m_data: current_date = m_data.group(1) if len(m_data.group(1)) == 10 else f"{m_data.group(1)}/{ano_detectado}"

                val_pattern = r'([−–—\-]?\s*(?:R\$\s*)?[\d\.]*,\d{2}\s*[−–—\-]?[CD]?)'
                valores = [m.strip() for m in re.findall(val_pattern, linha_limpa, re.IGNORECASE) if re.search(r'\d+,\d{2}', m)]

                if not valores or not current_date:
                    if transacoes and current_date and not re.search(r'\d+,\d{2}', linha_limpa):
                        txt = linha_limpa
                        m_data_any = re.search(r"\b(\d{2}/\d{2}(?:/\d{4})?)\b", txt)
                        if m_data_any: txt = txt.replace(m_data_any.group(1), "")
                        txt_limpo = limpar_historico_itau(txt)
                        if txt_limpo and not txt_limpo.isdigit():
                            transacoes[-1]["Histórico"] = limpar_historico_itau(transacoes[-1]["Histórico"] + " " + txt_limpo)
                    continue

                val_str = valores[0]
                desc = linha_limpa.replace(m_data.group(0), "", 1) if m_data else linha_limpa
                for v in valores: desc = desc.replace(v, "")

                desc_limpa = limpar_historico_itau(desc)
                
                # ---> BLOQUEIO DE SALDOS DIÁRIOS <---
                if desc_limpa.upper().startswith("SALDO") or desc_limpa.upper() == "S A L D O":
                    continue

                val_num, tipo = limpar_valor_itau(val_str)
                if val_num != 0.0:
                    transacoes.append({
                        "Data": current_date,
                        "Histórico": desc_limpa if desc_limpa else "LANCAMENTO ITAU",
                        "Valor": val_num, "Tipo": tipo
                    })

    if not transacoes: raise ValueError(f"Nenhuma transação encontrada no extrato Itaú '{os.path.basename(caminho_pdf)}'.")
    df = pd.DataFrame(transacoes)[["Data", "Histórico", "Valor", "Tipo"]]

    if not df.empty:
        df['Mes'] = df['Data'].str[3:5]
        df = df[df['Mes'] == df['Mes'].mode()[0]].drop(columns=['Mes'])

    caminho_saida = os.path.splitext(caminho_pdf)[0] + ".xlsx"
    df.to_excel(caminho_saida, index=False)
    return caminho_saida

def iniciar_processamento(caminho_pdf=None):
    if not caminho_pdf:
        caminho_pdf = filedialog.askopenfilename(title="Selecione o extrato Itaú (Modelo 1)", filetypes=[("Arquivos PDF", "*.pdf")])
    if not caminho_pdf: raise UserWarning("Nenhum arquivo selecionado.")
    if isinstance(caminho_pdf, (list, tuple)): caminho_pdf = caminho_pdf[0]
        
    try:
        df_transacoes = pd.read_excel(extrair_extrato_itau_consolidado(caminho_pdf))
        print(f"✅ Itaú Modelo 1: {len(df_transacoes)} lançamentos extraídos (Somente mês vigente).")
        return os.path.splitext(caminho_pdf)[0] + ".xlsx"
    except Exception as e:
        traceback.print_exc()
        raise Exception(f"Erro ao processar '{os.path.basename(caminho_pdf)}':\n{e}")

main = iniciar_processamento

if __name__ == "__main__":
    iniciar_processamento()