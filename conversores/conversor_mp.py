import pdfplumber
import re
import csv
import os
from decimal import Decimal

def main(pdf_path):
    """
    Conversor de Extrato Mercado Pago — PDF para CSV adaptado para o ConversorApp.
    """
    if not pdf_path:
        return None

    # Regex e Configurações Internas do Conversor
    ANCHOR_RE = re.compile(
        r'^(\d{2}-\d{2}-\d{4})' 
        r'(?:\s+(.+?))?' 
        r'\s+(\d{10,})' 
        r'\s+R\$\s*(-?[\d.]+,\d{2})' 
        r'\s+R\$\s*(-?[\d.]+,\d{2})\s*$'
    )

    PRE_MARKERS = (
        'Pagamento com Código QR', 'Pagamento do crédito', 'Débito por dívida',
        'Reembolso', 'Dinheiro retido', 'Transferência Pix enviada',
        'Transferência Pix recebida', 'Bônus por envio', 'Liberação de dinheiro'
    )

    TERMINAL_RE = re.compile(r'^Data de gera[çc][ãa]o:\s*\d{2}-\d{2}-\d{4}', re.IGNORECASE)

    SKIP_PATTERNS = tuple(
        re.compile(p, re.IGNORECASE) for p in (
            r'^EXTRATO DE CONTA', r'^CPF/CNPJ', r'^Per[íi]odo:', r'^Entradas:',
            r'^Sa[íi]das:', r'^Saldo (inicial|final):', r'^DETALHE DOS MOVIMENTOS',
            r'^Agência:', r'^\d+\s*/\s*\d+\s*$', r'^Data\s+Descri[cç][ãa]o.*Saldo'
        )
    )

    def looks_doubled(line: str) -> bool:
        s = ''.join(c for c in line if not c.isspace())
        if len(s) < 8: return False
        pair_count = len(s) // 2
        if pair_count == 0: return False
        matches = sum(1 for i in range(0, pair_count * 2, 2) if s[i] == s[i + 1])
        return matches / pair_count >= 0.7

    def heal_doubled_chars(line: str) -> str:
        if not looks_doubled(line): return line
        return re.sub(r'(.)\1+', r'\1', line)

    def is_skip(line: str) -> bool:
        line = line.strip()
        if not line: return True
        healed = heal_doubled_chars(line)
        return any(p.search(healed) for p in SKIP_PATTERNS)

    def starts_with_pre_marker(line: str) -> bool:
        return any(line.startswith(m) for m in PRE_MARKERS)

    def parse_value_to_decimal(s: str):
        if not s: return Decimal("0")
        s = s.strip().lstrip('-').replace('.', '').replace(',', '.')
        try: return Decimal(s)
        except: return Decimal("0")

    # --- PROCESSAMENTO DO PDF ---
    all_lines = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text: continue
                for line in text.split('\n'):
                    line = heal_doubled_chars(line.strip())
                    if TERMINAL_RE.match(line):
                        break
                    if is_skip(line): continue
                    all_lines.append(line)
    except Exception as e:
        print(f"Erro ao ler o PDF do Mercado Pago: {e}")
        return None

    # --- LÓGICA DE ANCORAGEM E REMONTAGEM ---
    anchors = []
    for i, line in enumerate(all_lines):
        m = ANCHOR_RE.match(line)
        if m:
            is_neg = m.group(4).strip().startswith('-')
            val_dec = parse_value_to_decimal(m.group(4))
            if is_neg: val_dec = -val_dec

            anchors.append({
                'idx': i,
                'date': m.group(1).replace('-', '/'),
                'desc_on': (m.group(2) or '').strip(),
                'value': val_dec,
                'pre': [],
                'post': [],
            })

    if not anchors:
        return None

    for k, anchor in enumerate(anchors):
        prev_idx = anchors[k-1]['idx'] if k > 0 else -1
        floating = all_lines[prev_idx + 1 : anchor['idx']]
        in_pre_mode = False
        for line in floating:
            if not in_pre_mode and starts_with_pre_marker(line):
                in_pre_mode = True
            if in_pre_mode:
                anchor['pre'].append(line)
            elif k > 0:
                anchors[k-1]['post'].append(line)

    last = anchors[-1]
    for line in all_lines[last['idx']+1:]:
        if starts_with_pre_marker(line): break
        last['post'].append(line)

    # --- MONTAGEM E ESCRITA DO CSV ---
    transactions = []
    for a in anchors:
        parts = list(a['pre'])
        if a['desc_on']: parts.append(a['desc_on'])
        parts.extend(a['post'])
        desc = re.sub(r'\s+', ' ', ' '.join(parts)).strip()
        transactions.append((a['date'], desc, a['value']))

    csv_path = os.path.splitext(pdf_path)[0] + ".csv"
    
    try:
        with open(csv_path, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['Data', 'Descrição', 'Valor'])
            for date, desc, value in transactions:
                value_str = str(value).replace('.', ',')
                writer.writerow([date, desc, value_str])
        return csv_path
    except Exception as e:
        print(f"Erro ao salvar arquivo CSV: {e}")
        return None