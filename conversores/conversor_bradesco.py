import os
import re
import csv
from dataclasses import dataclass, asdict
from typing import List, Optional
import pandas as pd
import pdfplumber

@dataclass
class Transaction:
    date: str
    description: str
    value: float = 0.0
    type: str = 'C'
    
    def to_dict(self):
        return {
            'Data': self.date,
            'Histórico': self.description,
            'Valor': self.value,
            'Tipo': self.type
        }

def parse_br_number(value: str) -> Optional[float]:
    if not value:
        return None
    try:
        return float(value.strip().replace('.', '').replace(',', '.'))
    except Exception:
        return None

def is_transaction_type(text: str) -> bool:
    """Identifica se o texto é uma descrição de operação (gatilho de início)."""
    text = text.strip().upper()
    type_patterns = [
        r'^TED-TRANSF', r'^TRANSFERENCIA PIX', r'^TRANSF CC PARA', r'^TRANSF PGTO PIX',
        r'^PIX QR CODE', r'^PIX - ENVIADO', r'^PIX - RECEBIDO', r'^PIX ENVIADO', r'^PIX RECEBIDO',
        r'^CARTAO VISA', r'^CARTAO MASTER', r'^CIELO VDA', r'^CIELO MASTER', r'^CIELO VISA',
        r'^CIELO ANTECIPACAO', r'^CIELO AMEX', r'^VENDA CARTAO', r'^CABAL DEBITO',
        r'^STONE MASTER', r'^STONE VISA', r'^STONE ELO', r'^STONE AMEX',
        r'^VISA ANTECIPACAO', r'^MASTER ANTECIPACAO', r'^ELO ANTECIPACAO', r'^AMEX ANTECIPACAO',
        r'^PAGTO ELETRON', r'^PAGTO SALARIO', r'^PGTO SALARIO', r'^PAGTO ELETRONICO',
        r'^RECEBIMENTO FORNECEDOR', r'^ENCARGOS C GARANTIDA', r'^TARIFA BANCARIA',
        r'^TV POR ASSINATURA', r'^RENTAB\.INVEST', r'^CONTA DE AGUA', r'^CONTA DE LUZ',
        r'^CONTA DE ENERGIA', r'^PAGAMENTO PEDÁGIO', r'^PAGAMENTO PEDAGIO'
    ]
    return any(re.match(p, text) for p in type_patterns)

def is_orphan_continuation(text: str) -> bool:
    """Detecta continuações curtas (ex: 'LTDA', 'S/A')."""
    text = text.strip()
    if len(text) <= 5 and text.isupper() and text.replace(' ', '').isalpha():
        return not is_transaction_type(text)
    return False

def is_skip_line(line: str) -> bool:
    """Filtra cabeçalhos, rodapés e todas as linhas de saldo do PDF."""
    skip_patterns = [
        'Extrato Mensal', 'CNPJ:', 'Nome do usuário', 'Data da operação',
        'Folha', 'bradesco', 'net empresa', 'Agência | Conta',
        'Total Disponível', 'Extrato de:', 'Data Lançamento',
        'Crédito (R$)', 'Débito (R$)', 'Saldo (R$)', 'Os dados acima',
        'Últimos Lançamentos', 'Saldos Invest', 'Data Histórico',
        'SALDO INVEST', 'SALDO ANTERIOR', 'SALDO FINAL', 'SALDO ATUAL',
        'SALDO TOTAL', 'SALDO DISPONIVEL', 'SALDO DISPONÍVEL', 'SALDO DO DIA',
        'SALDO EM C/C', 'S A L D O', 'Dcto.', 'Valor (R$)', 'Investimento sem',
        '(A+B)', '(B)', 'TOTAL APLICAÇÃO', 'APLICACAO AUTOMATICA', 'RESGATE AUTOMATICO'
    ]
    stripped = line.strip()
    if not stripped or stripped.startswith('Total') or re.match(r'^--- Page \d+ ---$', stripped):
        return True
    return any(skip.upper() in line.upper() for skip in skip_patterns)

def is_value_line(line: str) -> bool:
    """Verifica se a linha contém valores monetários no final."""
    return bool(re.search(r'-?[\d.]+,\d{2}\s*$', line.strip()))

def extract_transactions(pdf_path: str) -> List[Transaction]:
    """Extrai transações tratando descrições de múltiplas linhas sem barras verticais (|)."""
    all_lines = []
    if pdf_path.lower().endswith('.txt'):
        with open(pdf_path, 'r', encoding='utf-8') as f:
            all_lines = f.read().split('\n')
    else:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text(layout=True)
                if text:
                    all_lines.extend(text.split('\n'))

    lines_data = []
    for i, line in enumerate(all_lines):
        skip = is_skip_line(line)
        value = is_value_line(line) and not skip
        date_match = re.match(r'^(\d{2}/\d{2}/\d{4})', line.strip())
        date = date_match.group(1) if date_match else None
        
        clean = line.strip()
        if date: 
            clean = clean[len(date):].strip()
        
        is_type = is_transaction_type(clean) if not skip and not value else False
        is_orphan = is_orphan_continuation(clean) if not skip and not value and not is_type else False
        
        lines_data.append({
            'idx': i, 'raw': line, 'clean': clean, 'skip': skip, 
            'value': value, 'type': is_type, 'date': date, 'orphan': is_orphan
        })

    value_indices = [i for i, ld in enumerate(lines_data) if ld['value']]
    transactions = []
    current_date = None
    
    saldo_terms = [
        'SALDO ANTERIOR', 'SALDO FINAL', 'SALDO DO DIA', 'SALDO ATUAL',
        'SALDO TOTAL', 'SALDO DISPON', 'SALDO INVEST', 'S A L D O'
    ]

    for vi, data_idx in enumerate(value_indices):
        ld = lines_data[data_idx]
        line = ld['clean']
        if ld['date']: 
            current_date = ld['date']
        
        if any(term in line.upper() for term in saldo_terms):
            continue

        numbers = re.findall(r'(-?[\d.]+,\d{2})', line)
        if not numbers: 
            continue
        
        text_part = line
        for num in numbers: 
            text_part = text_part.replace(num, ' ')
        
        text_part = ' '.join(text_part.split()).strip('- ').strip()
        
        prev_value_data_idx = value_indices[vi-1] if vi > 0 else -1
        trans_type = ''
        
        for j in range(data_idx - 1, prev_value_data_idx, -1):
            if j < 0: 
                break
            prev_ld = lines_data[j]
            if prev_ld['skip'] or prev_ld['value'] or prev_ld['orphan']: 
                continue
            if prev_ld['date']: 
                current_date = prev_ld['date']
            if prev_ld['type']:
                trans_type = prev_ld['clean']
                break
        
        next_value_data_idx = value_indices[vi+1] if vi < len(value_indices)-1 else len(lines_data)
        counterparty_parts = []
        for j in range(data_idx + 1, next_value_data_idx):
            next_ld = lines_data[j]
            if next_ld['skip'] or next_ld['value'] or next_ld['type'] or next_ld['date']: 
                break
            
            clean_text = next_ld['clean']
            if clean_text:
                if next_ld['orphan'] and counterparty_parts:
                    counterparty_parts[-1] += ' ' + clean_text
                elif next_ld['orphan'] and text_part:
                    text_part += ' ' + clean_text
                else:
                    counterparty_parts.append(clean_text)
        
        # Junta partes com espaço simples e remove qualquer '|'
        parts = [p for p in [trans_type, text_part] + counterparty_parts if p]
        seen = set()
        description = ' '.join([p for p in parts if not (p in seen or seen.add(p))])
        description = description.replace('|', ' ')
        description = re.sub(r'\s+', ' ', description).strip()
        
        if not description or 'Total' in description or any(term in description.upper() for term in saldo_terms): 
            continue
        
        val_num = 0.0
        tipo = "C"
        
        if len(numbers) >= 3:
            val_cred = parse_br_number(numbers[-3])
            val_deb = parse_br_number(numbers[-2])
            if val_deb and val_deb != 0:
                val_num = -abs(val_deb)
                tipo = "D"
            elif val_cred and val_cred != 0:
                val_num = abs(val_cred)
                tipo = "C"
        elif len(numbers) >= 2:
            val = parse_br_number(numbers[-2])
            if val is not None:
                if val < 0:
                    val_num = val
                    tipo = "D"
                else:
                    val_num = val
                    tipo = "C"
        else:
            continue

        if val_num == 0.0 or val_num is None:
            continue
            
        if current_date:
            transactions.append(Transaction(
                date=current_date,
                description=description,
                value=val_num,
                type=tipo
            ))
            
    return transactions

def save_to_excel(transactions: List[Transaction], output_path: str):
    """Salva os dados consolidados em planilha Excel."""
    data = [t.to_dict() for t in transactions]
    df = pd.DataFrame(data)[["Data", "Histórico", "Valor", "Tipo"]]
    df.to_excel(output_path, index=False)

def main(file_path=None):
    if not file_path:
        return None

    if isinstance(file_path, (list, tuple)):
        file_path = file_path[0]

    try:
        transactions = extract_transactions(file_path)
        if not transactions:
            return None

        output_file = os.path.splitext(file_path)[0] + ".xlsx"
        save_to_excel(transactions, output_file)
        return output_file
    except Exception as e:
        print(f"Erro no processamento Bradesco: {e}")
        return None

def iniciar_processamento(file_path):
    return main(file_path)

class PDFTableExtractor:
    def __init__(self, file_path, configs=None):
        self.file_path = file_path

    def start(self):
        return main(self.file_path)