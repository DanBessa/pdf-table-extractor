import os
import re
import traceback
import pandas as pd
import pdfplumber

MESES_MAP = {
    'JAN': '01', 'FEV': '02', 'MAR': '03', 'ABR': '04',
    'MAI': '05', 'JUN': '06', 'JUL': '07', 'AGO': '08',
    'SET': '09', 'OUT': '10', 'NOV': '11', 'DEZ': '12',
    'JANEIRO': '01', 'FEVEREIRO': '02', 'MARÇO': '03', 'MARCO': '03',
    'ABRIL': '04', 'MAIO': '05', 'JUNHO': '06', 'JULHO': '07',
    'AGOSTO': '08', 'SETEMBRO': '09', 'OUTUBRO': '10', 'NOVEMBRO': '11',
    'DEZEMBRO': '12'
}

TERMOS_IGNORAR = [
    "SALDO ANTERIOR", "SALDO TOTAL", "SALDO CONTA", "SALDO RENDE",
    "CHEQUE ESPECIAL", "LIMITE CONTRATADO", "LIMITE UTILIZADO",
    "LIMITE DISPONÍVEL", "LIMITE DISPONIVEL", "RENDIMENTO PREVISTO",
    "ENTRADAS E SAÍDAS", "ENTRADAS E SAIDAS", "AGÊNCIA:", "AGENCIA:",
    "CONTA:", "CLIENTE:", "PERÍODO:", "PERIODO:", "COMPLEMENTO:",
    "EXTRATO CONSOLIDADO", "DATA/HORA EMISSÃO", "DATA/HORA EMISSAO",
    "HTTPS://", "HTTP://", "BANESTES INTERNET BANKING", "DATA LANÇAMENTO"
]

def limpar_float(val_str: str) -> float:
    """Converte valores monetários brasileiros (ex: -1.332,87 ou 0,02) para float."""
    v = re.sub(r'[^\d,\.\-−–—]', '', str(val_str))
    v = v.replace('−', '-').replace('–', '-').replace('—', '-')
    if ',' in v and '.' in v:
        v = v.replace('.', '').replace(',', '.')
    elif ',' in v:
        v = v.replace(',', '.')
    try:
        return float(v)
    except Exception:
        return 0.0

def extrair_dados_do_pdf(caminho_pdf):
    print(f"Iniciando extração Banestes: {caminho_pdf}")
    transacoes = []
    
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            texto_completo = ""
            for p in pdf.pages:
                texto_completo += (p.extract_text() or "") + "\n"

            # 1. Identifica Ano e Mês Padrão no cabeçalho do extrato
            ano_atual = "2026"
            mes_atual = "01"

            m_periodo = re.search(r'PER[IÍ]ODO:\s*(\d{2})/(\d{2})/(\d{4})', texto_completo, re.IGNORECASE)
            if m_periodo:
                mes_atual = m_periodo.group(2)
                ano_atual = m_periodo.group(3)
            else:
                m_data = re.search(r'\b(\d{2})/(\d{2})/(\d{4})\b', texto_completo)
                if m_data:
                    mes_atual = m_data.group(2)
                    ano_atual = m_data.group(3)

            dia_atual = "01"

            # 2. Estratégia Híbrida: Tabelas Nativas + Texto
            for page in pdf.pages:
                tabelas = page.extract_tables() or []
                processou_tabela = False

                for tab in tabelas:
                    for linha in tab:
                        if not linha:
                            continue
                        
                        celulas = [str(c).strip() for c in linha if c is not None and str(c).strip()]
                        if not celulas:
                            continue

                        texto_linha = " ".join(celulas)

                        # Atualiza dia e mês se presentes
                        m_d = re.search(r'\b(\d{1,2})\b', celulas[0])
                        if m_d and len(celulas[0]) <= 8:
                            dia_atual = m_d.group(1).zfill(2)

                        for sigla, mes_num in MESES_MAP.items():
                            if sigla in celulas[0].upper():
                                mes_atual = mes_num
                                break

                        # Ignora saldos e cabeçalhos
                        if any(term in texto_linha.upper() for term in TERMOS_IGNORAR):
                            continue

                        # Procura valor monetário na última célula ou no texto da linha
                        m_val = re.search(r'([−–—\-]?\s*(?:R\$\s*)?[\d\.]*,\d{2}\s*[−–—\-]?)', celulas[-1])
                        if not m_val:
                            m_val = re.search(r'([−–—\-]?\s*(?:R\$\s*)?[\d\.]*,\d{2}\s*[−–—\-]?)', texto_linha)

                        if m_val:
                            val_str = m_val.group(1)
                            f_val = limpar_float(val_str)

                            # Remove dia, mês e valor da descrição
                            desc_celulas = [c for c in celulas if c != val_str and not re.match(r'^\d{1,2}$', c) and c.upper() not in MESES_MAP]
                            desc_bruta = " ".join(desc_celulas)
                            desc_limpa = re.sub(r'[↑✔Π目\[\]]', '', desc_bruta).replace(val_str, '').strip()
                            desc_limpa = re.sub(r'\s+', ' ', desc_limpa).strip()

                            if not desc_limpa or any(term in desc_limpa.upper() for term in TERMOS_IGNORAR):
                                continue

                            # Ajuste de sinal
                            if any(k in desc_limpa.upper() for k in ['PIX ENVIADO', 'PAGAMENTO', 'TARIFA', 'CESTA', 'DÉBITO', 'DEBITO']) and f_val > 0:
                                f_val = -f_val

                            if f_val != 0.0:
                                transacoes.append({
                                    "Data": f"{dia_atual}/{mes_atual}/{ano_atual}",
                                    "Histórico": desc_limpa,
                                    "Valor": f_val,
                                    "Tipo": "D" if f_val < 0 else "C"
                                })
                                processou_tabela = True

                # 3. Fallback: Leitura Linha a Linha (caso não detecte tabela)
                if not processou_tabela:
                    texto_pag = page.extract_text(layout=True) or ""
                    for linha in texto_pag.split("\n"):
                        linha_s = linha.strip()
                        if not linha_s or any(t in linha_s.upper() for t in TERMOS_IGNORAR):
                            continue

                        m_dia_txt = re.match(r'^(\d{1,2})\b', linha_s)
                        if m_dia_txt:
                            dia_atual = m_dia_txt.group(1).zfill(2)

                        for sigla, mes_num in MESES_MAP.items():
                            if re.search(rf'\b{sigla}\b', linha_s.upper()):
                                mes_atual = mes_num
                                break

                        val_match = re.findall(r'([−–—\-]?\s*(?:R\$\s*)?[\d\.]*,\d{2}\s*[−–—\-]?)', linha_s)
                        if val_match:
                            val_str = val_match[-1]
                            f_val = limpar_float(val_str)

                            desc_bruta = linha_s.replace(val_str, "")
                            if m_dia_txt:
                                desc_bruta = desc_bruta.replace(m_dia_txt.group(1), "", 1)

                            desc_limpa = re.sub(r'[↑✔Π目\[\]]', '', desc_bruta).strip()
                            desc_limpa = re.sub(r'\s+', ' ', desc_limpa).strip()

                            if desc_limpa and not any(t in desc_limpa.upper() for t in TERMOS_IGNORAR):
                                if any(k in desc_limpa.upper() for k in ['PIX ENVIADO', 'PAGAMENTO', 'TARIFA', 'CESTA', 'DÉBITO', 'DEBITO']) and f_val > 0:
                                    f_val = -f_val

                                if f_val != 0.0:
                                    transacoes.append({
                                        "Data": f"{dia_atual}/{mes_atual}/{ano_atual}",
                                        "Histórico": desc_limpa,
                                        "Valor": f_val,
                                        "Tipo": "D" if f_val < 0 else "C"
                                    })
                        elif transacoes and len(linha_s) > 3:
                            # Concatena nome de favorecido em linha quebrada
                            desc_extra = re.sub(r'[↑✔Π目\[\]]', '', linha_s).strip()
                            if desc_extra and not any(t in desc_extra.upper() for t in TERMOS_IGNORAR):
                                transacoes[-1]["Histórico"] = f"{transacoes[-1]['Histórico']} {desc_extra}".strip()

    except Exception as e:
        print(f"Erro no processamento do Banestes: {e}")
        traceback.print_exc()
        raise e

    if not transacoes:
        return pd.DataFrame()

    return pd.DataFrame(transacoes)[["Data", "Histórico", "Valor", "Tipo"]]

def iniciar_processamento(pdf_path):
    if not pdf_path:
        raise UserWarning("Nenhum arquivo selecionado.")

    try:
        df_transacoes = extrair_dados_do_pdf(pdf_path)
        
        if df_transacoes is None or df_transacoes.empty:
            raise UserWarning(f"Nenhuma transação válida foi encontrada em {os.path.basename(pdf_path)}.")

        caminho_saida = os.path.splitext(pdf_path)[0] + ".xlsx"
        df_transacoes.to_excel(caminho_saida, index=False)
        return caminho_saida
        
    except Exception as e:
        traceback.print_exc()
        raise Exception(f"Erro ao processar extrato Banestes:\n{e}")