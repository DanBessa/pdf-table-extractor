import os
import re
import sys
from datetime import datetime
import pandas as pd

def extrair_texto_pdf(caminho_pdf: str) -> str:
    """Lê todo o texto do PDF com fallback entre pdfplumber e pypdf."""
    texto_completo = ""
    try:
        import pypdf
        with open(caminho_pdf, "rb") as f:
            reader = pypdf.PdfReader(f)
            for page in reader.pages:
                t = page.extract_text() or ""
                texto_completo += t + "\n"
    except Exception:
        import pdfplumber
        with pdfplumber.open(caminho_pdf) as pdf:
            for page in pdf.pages:
                t = page.extract_text() or ""
                texto_completo += t + "\n"
    return texto_completo

def extrair_extrato_stone(caminho_pdf: str) -> str:
    full_text = extrair_texto_pdf(caminho_pdf)
    if not full_text.strip():
        raise ValueError(f"Não foi possível ler o texto do arquivo '{os.path.basename(caminho_pdf)}'.")

    # Normaliza caracteres especiais de fontes do Stone
    full_text = full_text.replace('\ue088', '-').replace('\ue092', ':')
    lines = [l.strip() for l in full_text.split('\n') if l.strip()]

    # Localiza o início de cada transação (Data + Tipo)
    tx_starts = []
    for idx, line in enumerate(lines):
        if re.match(r'^\d{2}/\d{2}/\d{2,4}\s+(?:Sa[íi]da|Entrada)', line, re.IGNORECASE):
            tx_starts.append(idx)

    if not tx_starts:
        raise ValueError(f"Nenhuma transação Stone identificada em '{os.path.basename(caminho_pdf)}'.")

    footer_markers = [
        "Informações do Comprovante", "Código da autenticação", "Ouvidoria",
        "Se nosso atendimento", "úteis, das", "horário de Brasília", "ouvidoria@stone.com.br",
        "CNPJ:", "Dúvidas?", "Regiões Metropolitanas", "Envie um", "Outras regiões",
        "meajuda@stone.com.br", "Extrato de conta corrente", "Emitido em", "Página",
        "Período:", "DATA TIPO DESCRIÇÃO VALOR SALDO CONTRAPARTE"
    ]

    transactions = []

    for i in range(len(tx_starts)):
        start_idx = tx_starts[i]
        end_idx = tx_starts[i + 1] if i + 1 < len(tx_starts) else len(lines)
        tx_lines = lines[start_idx:end_idx]

        filtered_lines = []
        for l in tx_lines:
            if any(marker in l for marker in ["Informações do Comprovante", "Código da autenticação", "Ouvidoria"]):
                break
            if any(term in l for term in footer_markers):
                continue
            if re.search(r'[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}', l, re.I):
                break
            filtered_lines.append(l)

        if not filtered_lines:
            continue

        first_line = filtered_lines[0]
        m_head = re.match(r'^(\d{2}/\d{2}/\d{2,4})\s+(Sa[íi]da|Entrada)\s*(.*)$', first_line, re.IGNORECASE)
        if not m_head:
            continue

        data_raw = m_head.group(1)
        tipo_raw = m_head.group(2)
        resto_first_line = m_head.group(3).strip()

        # Formata data para DD/MM/AAAA
        d_parts = data_raw.split('/')
        if len(d_parts) == 3 and len(d_parts[2]) == 2:
            data_fmt = f"{d_parts[0].zfill(2)}/{d_parts[1].zfill(2)}/20{d_parts[2]}"
        else:
            data_fmt = data_raw

        eh_saida = "SAÍDA" in tipo_raw.upper() or "SAIDA" in tipo_raw.upper()
        tipo_final = "D" if eh_saida else "C"

        remaining_lines = ([resto_first_line] if resto_first_line else []) + filtered_lines[1:]

        desc_parts = []
        val_num = 0.0
        saldo_num = 0.0
        contraparte_parts = []
        found_values = False

        for l in remaining_lines:
            r_matches = list(re.finditer(r'(-?\s*R\$\s*[\d\.,]+)', l))

            if r_matches:
                found_values = True
                text_before_r = l[:r_matches[0].start()].strip()
                if text_before_r:
                    desc_parts.append(text_before_r)

                v_str = r_matches[0].group(1)
                v_clean = re.sub(r'[^\d,\.]', '', v_str)
                if ',' in v_clean and '.' in v_clean:
                    v_clean = v_clean.replace('.', '').replace(',', '.')
                elif ',' in v_clean:
                    v_clean = v_clean.replace(',', '.')
                try:
                    val = float(v_clean)
                except ValueError:
                    val = 0.0
                val_num = -abs(val) if eh_saida or '-' in v_str else val

                if len(r_matches) >= 2:
                    s_str = r_matches[1].group(1)
                    s_clean = re.sub(r'[^\d,\.]', '', s_str)
                    if ',' in s_clean and '.' in s_clean:
                        s_clean = s_clean.replace('.', '').replace(',', '.')
                    elif ',' in s_clean:
                        s_clean = s_clean.replace(',', '.')
                    try:
                        saldo_num = float(s_clean)
                    except ValueError:
                        saldo_num = 0.0

                text_after_r = l[r_matches[-1].end():].strip()
                if text_after_r:
                    contraparte_parts.append(text_after_r)
            else:
                if not found_values:
                    desc_parts.append(l)
                else:
                    contraparte_parts.append(l)

        desc_final = re.sub(r'\s+', ' ', " ".join(desc_parts)).strip()
        contraparte_final = re.sub(r'\s+', ' ', " ".join(contraparte_parts)).strip()

        transactions.append({
            'Data': data_fmt,
            'Tipo': tipo_final,
            'Descrição': desc_final,
            'Valor': val_num,
            'Saldo': saldo_num,
            'Contraparte': contraparte_final
        })

    if not transactions:
        raise ValueError("Falha ao estruturar as transações do extrato Stone.")

    df = pd.DataFrame(transactions)
    caminho_saida = os.path.splitext(caminho_pdf)[0] + ".xlsx"
    df.to_excel(caminho_saida, index=False)
    return caminho_saida

# ==========================================
# 🚀 PONTOS DE ENTRADA DO MÓDULO
# ==========================================
class PDFTableExtractor:
    def __init__(self, file_path, configs=None):
        self.file_path = file_path

    def start(self):
        return extrair_extrato_stone(self.file_path)

def iniciar_processamento(caminho_pdf):
    if isinstance(caminho_pdf, (list, tuple)):
        caminho_pdf = caminho_pdf[0]
    return extrair_extrato_stone(caminho_pdf)

def main(caminho_pdf=None):
    if len(sys.argv) > 1 and not caminho_pdf:
        caminho_pdf = sys.argv[1]
    return iniciar_processamento(caminho_pdf)

if __name__ == "__main__":
    if len(sys.argv) > 1:
        res = main(sys.argv[1])
        print(f"Gerado com sucesso em: {res}")