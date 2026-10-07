import os
import re
import pandas as pd
import pdfplumber
from tkinter import filedialog
import traceback

def limpar_valor_cef(val_str: str):
    """Converte valor monetário da Caixa (com suporte a 'C', 'D' e '-') para float e tipo."""
    if not val_str or pd.isna(val_str):
        return 0.0, "C"
        
    v = str(val_str).strip().upper()
    eh_debito = "D" in v or "-" in v
    
    v_limpo = re.sub(r"[^\d,\.]", "", v)
    if "," in v_limpo and "." in v_limpo:
        v_limpo = v_limpo.replace(".", "").replace(",", ".")
    elif "," in v_limpo:
        v_limpo = v_limpo.replace(",", ".")
        
    try:
        val_float = float(v_limpo)
        if eh_debito and val_float > 0:
            val_float = -val_float
        tipo = "D" if val_float < 0 else "C"
        return val_float, tipo
    except ValueError:
        return 0.0, "C"

def extrair_cef_mod1(caminho_pdf: str) -> str:
    transacoes = []
    
    # Termos de bloqueio para garantir que não pegaremos linhas exclusivas de sumários ou cabeçalhos
    ignorar_termos = [
        "SALDO ANTERIOR", "SALDO FINAL", "SALDO DISPON", 
        "SALDO BLOQUEADO", "LIMITE", "SALDO DO DIA", "SALDO EM",
        "TOTAL", "EXTRATO DE", "NOME:", "CONTA:", "MÊS:"
    ]

    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            current_date = None
            
            for page in pdf.pages:
                # Usa extração de palavras por coordenadas para FORÇAR a ordem de leitura (Esquerda -> Direita)
                words = page.extract_words(x_tolerance=2, y_tolerance=3, keep_blank_chars=False)
                if not words:
                    continue

                # Agrupa palavras por linha visual usando o eixo Y, e ordena pelo eixo X
                words_ordenadas = sorted(words, key=lambda w: (w["top"], w["x0"]))
                linhas_visuais = []
                for w in words_ordenadas:
                    if not linhas_visuais or abs(w["top"] - linhas_visuais[-1]["top"]) > 3.5:
                        linhas_visuais.append({"top": w["top"], "words": [w]})
                    else:
                        linhas_visuais[-1]["words"].append(w)

                for item_linha in linhas_visuais:
                    palavras = item_linha["words"]
                    # linha_completa agora reflete EXATAMENTE o que você vê na tela (Valor primeiro, Saldo depois)
                    linha_completa = " ".join(w["text"] for w in palavras).strip()
                    linha_upper = linha_completa.upper()

                    # Ignora sumários e linhas exclusivas de saldo geral
                    if any(term in linha_upper for term in ignorar_termos):
                        continue

                    # 1. Identifica a Data (Padrão Caixa: DD/MM/AAAA)
                    # Olha nas primeiras palavras (mais à esquerda) para a data
                    txt_left = " ".join(w["text"] for w in palavras if w["x0"] < 150).strip()
                    m_data = re.search(r"\b(\d{2}/\d{2}/\d{4})\b", txt_left)
                    if m_data:
                        current_date = m_data.group(1)

                    # 2. Encontra todos os valores monetários na linha
                    val_pattern = r'([−–—\-]?\s*[\d\.]*,\d{2}\s*[−–—\-]?[CDcd]?)'
                    valores = [m.strip() for m in re.findall(val_pattern, linha_completa) if re.search(r'\d+,\d{2}', m)]

                    if not valores or not current_date:
                        continue

                    # MÁGICA VISUAL: Como forçamos a ordenação pelo eixo X, 
                    # valores[0] é GARANTIDAMENTE a coluna "Valor", e valores[-1] será a coluna "Saldo".
                    val_str = valores[0]

                    # 3. Monta o histórico limpando a data e TODOS os números de valores (para não sujar a string)
                    desc_bruta = linha_completa
                    if m_data:
                        desc_bruta = desc_bruta.replace(m_data.group(1), "", 1)
                    
                    for v in valores:
                        desc_bruta = desc_bruta.replace(v, "")

                    desc_limpa = re.sub(r'\s+', ' ', desc_bruta).strip()
                    desc_limpa = re.sub(r"^-+|-+$", "", desc_limpa).strip() # Limpa traços soltos
                    
                    # ---> BLOQUEIO DE SALDOS DIÁRIOS <---
                    # Se o que sobrou na descrição for a palavra "SALDO" ou "S A L D O", abortamos essa linha.
                    desc_upper = desc_limpa.upper()
                    if desc_upper.startswith("SALDO") or desc_upper == "S A L D O" or desc_upper == "SDO":
                        continue
                    
                    val_num, tipo = limpar_valor_cef(val_str)

                    if val_num != 0.0:
                        transacoes.append({
                            "Data": current_date,
                            "Histórico": desc_limpa if desc_limpa else "LANCAMENTO CAIXA",
                            "Valor": val_num,
                            "Tipo": tipo
                        })

        if not transacoes:
            raise ValueError(f"Nenhuma transação encontrada no extrato '{os.path.basename(caminho_pdf)}'.")

        # Transforma em Dataframe e salva como Excel, que se comunica perfeitamente com seu Gerador de TXT
        df = pd.DataFrame(transacoes)[["Data", "Histórico", "Valor", "Tipo"]]
        
        caminho_saida = os.path.splitext(caminho_pdf)[0] + ".xlsx"
        df.to_excel(caminho_saida, index=False)
        return caminho_saida

    except Exception as e:
        print(f"Erro ao processar Caixa Econômica Mod 1: {e}")
        traceback.print_exc()
        raise e

def main(pdf_path=None):
    if not pdf_path:
        pdf_path = filedialog.askopenfilename(
            title="Selecione o extrato da Caixa (Modelo 1)",
            filetypes=[("Arquivos PDF", "*.pdf")]
        )
    if not pdf_path:
        raise UserWarning("Nenhum arquivo selecionado.")
        
    if isinstance(pdf_path, (list, tuple)):
        pdf_path = pdf_path[0]
        
    try:
        return extrair_cef_mod1(pdf_path)
    except Exception as e:
        traceback.print_exc()
        raise Exception(f"Erro ao processar Caixa Econômica:\n{e}")

# Mantido para compatibilidade com o sistema principal
iniciar_processamento = main

if __name__ == "__main__":
    main()