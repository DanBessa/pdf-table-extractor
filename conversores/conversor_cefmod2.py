import os
import re
import pandas as pd
import pdfplumber

def formatar_valor(valor_str: str):
    if not valor_str:
        return 0.0, "C"
        
    v = str(valor_str).strip()
    eh_debito = "-" in v or v.endswith("D")
    
    v_limpo = re.sub(r"[^\d,\.]", "", v)
    if "," in v_limpo and "." in v_limpo:
        v_limpo = v_limpo.replace(".", "").replace(",", ".")
    elif "," in v_limpo:
        v_limpo = v_limpo.replace(",", ".")
        
    try:
        valor_float = float(v_limpo)
        if eh_debito and valor_float > 0:
            valor_float = -valor_float
        tipo = "D" if eh_debito else "C"
        return valor_float, tipo
    except ValueError:
        return 0.0, "C"

def formatar_saldo(saldo_str: str):
    if not saldo_str:
        return 0.0
    val, _ = formatar_valor(saldo_str)
    return abs(val)

def normalizar_espacos(texto: str) -> str:
    if not texto:
        return ""
    return re.sub(r"\s+", " ", str(texto)).strip()

class ExtratorCaixaMod1:
    def __init__(self, caminho_pdf):
        self.caminho_pdf = caminho_pdf
        self.pasta_saida = os.path.dirname(os.path.abspath(caminho_pdf))
        self.nome_base = os.path.splitext(os.path.basename(caminho_pdf))[0]

    def extrair(self) -> str:
        dados = []

        with pdfplumber.open(self.caminho_pdf) as pdf:
            for page in pdf.pages:
                tabelas = page.extract_tables()
                tabela_processada = False

                for tabela in tabelas:
                    for linha in tabela:
                        if not linha or len(linha) < 4:
                            continue
                        item = self._processar_linha(linha)
                        if item:
                            dados.append(item)
                            tabela_processada = True

                if not tabela_processada:
                    texto = page.extract_text() or ""
                    itens_texto = self._processar_fallback_texto(texto)
                    dados.extend(itens_texto)

        if not dados:
            return None

        df = pd.DataFrame(dados)
        # df = df.drop_duplicates().reset_index(drop=True)

        caminho_excel = os.path.join(self.pasta_saida, f"{self.nome_base}.xlsx")
        with pd.ExcelWriter(caminho_excel, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Extrato Caixa Mod1")

        return caminho_excel

    def _processar_linha(self, linha):
        try:
            col_data = str(linha[0] or "").strip()
            if not re.search(r"\d{2}/\d{2}/\d{4}", col_data):
                return None

            partes_data = col_data.split("\n")
            data_mov = partes_data[0].strip()
            data_efetiva = partes_data[1].strip() if len(partes_data) > 1 else ""

            doc_idx = 1
            if len(linha) >= 6 and re.search(r"\d{2}/\d{2}", str(linha[1] or "")):
                data_efetiva = normalizar_espacos(linha[1])
                doc_idx = 2

            documento = normalizar_espacos(linha[doc_idx])
            historico = normalizar_espacos(linha[doc_idx + 1])
            valor_bruto = str(linha[doc_idx + 2] or "").strip()
            saldo_bruto = str(linha[doc_idx + 3] or "").strip() if len(linha) > doc_idx + 3 else ""

            if "SALDO DIA" in historico.upper() or "DATA" in data_mov.upper():
                return None

            valor_num, tipo = formatar_valor(valor_bruto)
            saldo_num = formatar_saldo(saldo_bruto)

            return {
                "Data": data_mov,
                "Data Efetiva": data_efetiva,
                "Documento": documento,
                "Histórico": historico,
                "Valor": valor_num,
                "Tipo": tipo,
                "Saldo": saldo_num
            }
        except Exception:
            return None

    def _processar_fallback_texto(self, texto):
        itens = []
        blocos = re.split(r"(?=\n\d{2}/\d{2}/\d{4})", texto)

        for bloco in blocos:
            bloco = bloco.strip()
            if not bloco or not re.match(r"^\d{2}/\d{2}/\d{4}", bloco) or "SALDO DIA" in bloco.upper():
                continue

            partes = [p.strip() for p in bloco.split("\n") if p.strip()]
            if len(partes) < 3:
                continue

            data_mov = partes[0]
            valores = re.findall(r"(-?\s*R\$\s*[\d\.,]+(?:\s*[CD])?)", bloco)
            if not valores:
                continue

            valor_bruto = valores[-2] if len(valores) >= 2 else valores[-1]
            saldo_bruto = valores[-1] if len(valores) >= 2 else ""

            valor_num, tipo = formatar_valor(valor_bruto)
            saldo_num = formatar_saldo(saldo_bruto)

            miolo = bloco.replace(data_mov, "")
            for v in valores:
                miolo = miolo.replace(v, "")

            miolo = re.sub(r"about:blank|\d+/\d+|\d{2}/\d{2}/\d{4},\s*\d{2}:\d{2}|extrato_pdf", "", miolo)
            miolo = normalizar_espacos(miolo)

            doc_match = re.search(r"\b(\d{6})\b", miolo)
            documento = doc_match.group(1) if doc_match else ""
            historico = miolo.replace(documento, "").strip() if documento else miolo

            itens.append({
                "Data": data_mov,
                "Data Efetiva": "",
                "Documento": documento,
                "Histórico": historico,
                "Valor": valor_num,
                "Tipo": tipo,
                "Saldo": saldo_num
            })

        return itens

def iniciar_processamento(caminho_pdf):
    if isinstance(caminho_pdf, (list, tuple)):
        caminho_pdf = caminho_pdf[0]
    return ExtratorCaixaMod1(caminho_pdf).extrair()