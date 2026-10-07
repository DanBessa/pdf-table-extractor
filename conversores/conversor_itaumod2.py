import os
import re
import pandas as pd
import pdfplumber
from tkinter import filedialog
import traceback

def limpar_valor_itau(val_str: str):
    if not val_str or pd.isna(val_str):
        return 0.0, "C"

    v = str(val_str).strip()
    eh_debito = any(c in v for c in ['-', '−', '–', '—']) or v.endswith("D") or v.endswith("d")
    
    v_limpo = re.sub(r"[^\d,\.]", "", v)
    if "," in v_limpo and "." in v_limpo:
        v_limpo = v_limpo.replace(".", "").replace(",", ".")
    elif "," in v_limpo:
        v_limpo = v_limpo.replace(",", ".")

    try:
        val_float = float(v_limpo)
        if eh_debito and val_float > 0:
            val_float = -val_float
        elif not eh_debito and val_float < 0:
            val_float = abs(val_float)
        tipo = "D" if val_float < 0 else "C"
        return val_float, tipo
    except ValueError:
        return 0.0, "C"

def limpar_historico_mod2(desc_bruta: str) -> str:
    t = str(desc_bruta).replace('|', ' ')
    
    # 1. Deleta prefixos bancários gerados pelo Itaú Empresas
    termos_remover = [
        r"PIX\s+(?:RECEBIDO|ENVIADO)\s+[A-Z0-9/]+\b", # Remove "PIX RECEBIDO DOUGLAS16/06"
        r"COMPRA NO D[EÉ]BITO\s+",
        r"PAGAMENTOS?\s+PIX QR-CODE\s+",
        r"PAGAMENTOS?\s+TRIB COD BARRAS\s+",
        r"RECEBIMENTOS?\s+",
        r"TED\s+[\d\.]+\s*",
        r"DOC\s+[\d\.]+\s*",
        r"BOLETO PAGO\s+"
    ]
    for termo in termos_remover:
        t = re.sub(termo, "", t, flags=re.IGNORECASE)
        
    # 2. Deleta Documentos (CPF e CNPJ)
    t = re.sub(r'\b\d{3}\.\d{3}\.\d{3}-\d{2}\b', '', t)
    t = re.sub(r'\b\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\b', '', t)
    
    # 3. Deleta Códigos internos de maquininha (ex: SUPERMERCAD-001008, SABO- 001008)
    t = re.sub(r'\b[A-Z0-9]+\s*-\s*\d{6}\b', '', t)
    t = re.sub(r'(?<!\S)-\s*\d{6}\b', '', t)
    
    # 4. Limpa traços órfãos e formatação de texto
    t = re.sub(r'(?<!\S)[-\/](?!\S)', '', t)
    t = re.sub(r'\s+', ' ', t).strip()
    return t

def extrair_itaumod2(caminho_pdf: str) -> str:
    transacoes = []
    
    ignorar_termos = [
        "SALDO ANTERIOR", "SALDO FINAL", "SALDO ATUAL", "SALDO DISPONÍVEL", "SALDO TOTAL",
        "SALDO DO DIA", "SALDO EM CONTA", "EXTRATO DE CONTA CORRENTE",
        "Lançamentos do período", "Razão Social", "CNPJ/CPF", "Valor (R$)", "Saldo (R$)"
    ]

    with pdfplumber.open(caminho_pdf) as pdf:
        tx_atual = {"data": None, "linhas_desc": [], "valor": 0.0, "tipo": ""}
        
        def salvar_tx():
            if tx_atual["data"] and tx_atual["valor"] != 0.0:
                desc_bruta = " ".join(tx_atual["linhas_desc"]).strip()
                desc_limpa = limpar_historico_mod2(desc_bruta)
                
                transacoes.append({
                    "Data": tx_atual["data"],
                    "Histórico": desc_limpa if desc_limpa else "LANCAMENTO ITAU",
                    "Valor": tx_atual["valor"],
                    "Tipo": tx_atual["tipo"]
                })
        
        for page in pdf.pages:
            # Leitura espacial
            words = page.extract_words(x_tolerance=2, y_tolerance=3, keep_blank_chars=False)
            if not words: continue
            
            words_ordenadas = sorted(words, key=lambda w: (w["top"], w["x0"]))
            linhas_visuais = []
            for w in words_ordenadas:
                if not linhas_visuais or abs(w["top"] - linhas_visuais[-1]["top"]) > 3.5:
                    linhas_visuais.append({"top": w["top"], "words": [w]})
                else:
                    linhas_visuais[-1]["words"].append(w)
                    
            for item in linhas_visuais:
                linha_completa = " ".join(w["text"] for w in item["words"]).strip()
                linha_upper = linha_completa.upper()
                
                if any(term in linha_upper for term in ignorar_termos):
                    continue
                    
                # Procura a data preferencialmente na margem esquerda
                txt_left = " ".join(w["text"] for w in item["words"] if w["x0"] < 100).strip()
                m_data = re.search(r"\b(\d{2}/\d{2}/\d{4})\b", txt_left)
                if not m_data:
                    m_data = re.search(r"^\b(\d{2}/\d{2}/\d{4})\b", linha_completa)
                
                tem_data = bool(m_data)
                
                # CORTINA DE FERRO CONTRA SALDOS FALSOS: 
                # Ignoramos palavras capturadas na extrema direita (X > 520) que pertencem à coluna de Saldo
                palavras_valor = [w for w in item["words"] if w["x0"] < 520]
                linha_sem_saldo = " ".join(w["text"] for w in palavras_valor).strip()
                
                val_pattern = r'([−–—\-]?\s*(?:R\$\s*)?[\d\.]*,\d{2}\s*[−–—\-]?[CDcd]?)'
                valores = [m.strip() for m in re.findall(val_pattern, linha_sem_saldo) if re.search(r'\d+,\d{2}', m)]
                
                if tem_data:
                    if tx_atual["linhas_desc"]:
                        salvar_tx() # Guarda a transação anterior completa
                    tx_atual = {"data": m_data.group(1), "linhas_desc": [], "valor": 0.0, "tipo": ""}
                
                if valores and tx_atual["data"]:
                    v_str = valores[0]
                    v_num, v_tipo = limpar_valor_itau(v_str)
                    
                    if tx_atual["valor"] == 0.0 and v_num != 0.0:
                        tx_atual["valor"] = v_num
                        tx_atual["tipo"] = v_tipo
                
                # Monta a descrição
                desc = linha_completa
                if tem_data:
                    desc = desc.replace(m_data.group(1), "", 1)
                
                # Tira o(s) valor(es) financeiro(s) do meio do texto da descrição
                val_pattern_todos = r'([−–—\-]?\s*(?:R\$\s*)?[\d\.]*,\d{2}\s*[−–—\-]?[CDcd]?)'
                for v in re.findall(val_pattern_todos, linha_completa):
                    desc = desc.replace(v, "")
                
                desc = desc.strip()
                if desc and desc not in ["|", "-"]:
                    tx_atual["linhas_desc"].append(desc)
                    
        # Salva o último lançamento processado do extrato
        if tx_atual["linhas_desc"]:
            salvar_tx()

    if not transacoes:
        raise ValueError(f"Nenhuma transação encontrada no extrato Itaú (Modelo 2).")

    df = pd.DataFrame(transacoes)[["Data", "Histórico", "Valor", "Tipo"]]
    
    caminho_saida = os.path.splitext(caminho_pdf)[0] + ".xlsx"
    df.to_excel(caminho_saida, index=False)
    return caminho_saida

def iniciar_processamento(caminho_pdf=None):
    if not caminho_pdf:
        caminho_pdf = filedialog.askopenfilename(
            title="Selecione o extrato Itaú (Modelo 2)",
            filetypes=[("Arquivos PDF", "*.pdf")]
        )
    if not caminho_pdf:
        raise UserWarning("Nenhum arquivo selecionado.")
        
    if isinstance(caminho_pdf, (list, tuple)):
        caminho_pdf = caminho_pdf[0]
        
    try:
        df_transacoes = pd.read_excel(extrair_itaumod2(caminho_pdf))
        total_cred = len(df_transacoes[df_transacoes["Tipo"] == "C"])
        total_deb = len(df_transacoes[df_transacoes["Tipo"] == "D"])
        print(f"✅ Itaú Modelo 2: {len(df_transacoes)} lançamentos extraídos.")
        print(f"   🟢 Entradas (+ / C): {total_cred}")
        print(f"   🔴 Saídas (- / D): {total_deb}")
        return caminho_pdf.replace('.pdf', '.xlsx')
    except Exception as e:
        traceback.print_exc()
        raise Exception(f"Erro ao processar '{os.path.basename(caminho_pdf)}':\n{e}")

# Compatibilidade
main = iniciar_processamento

if __name__ == "__main__":
    iniciar_processamento()