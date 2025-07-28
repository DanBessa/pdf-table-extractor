import os
import re
import pdfplumber
from tkinter import filedialog
from xlwt import Workbook
import traceback

def extract_pdf_to_text(pdf_path):
    """Extrai o texto do PDF para um arquivo de texto temporário."""
    base_name = os.path.splitext(pdf_path)[0]
    output_path = base_name + "_temp.txt"
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            with open(output_path, 'w', encoding='utf-8') as txt_file:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        txt_file.write(text + '\n\n')
        return output_path
    except Exception as e:
        print(f"Ocorreu um erro ao extrair texto do Bradesco: {str(e)}")
        return None

def clean_statement(file_path: str):
    """Limpa o arquivo de texto, removendo cabeçalhos e rodapés."""
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            lines = [line.strip() for line in file.readlines() if line.strip()]
        
        if len(lines) > 5:
            lines = lines[5:]
        else:
            lines = []

        total_index = -1
        for i, line in enumerate(lines):
            if "total" in line.lower():
                total_index = i
                break
        if total_index != -1:
            lines = lines[:total_index]
            
        with open(file_path, 'w', encoding='utf-8') as file:
            file.write('\n'.join(lines))
    except Exception as e:
        print(f"Erro em clean_statement: {e}")

def mark_all_transaction_blocks(file_path: str):
    """Marca blocos de transação de 3 linhas para facilitar o processamento."""
    with open(file_path, 'r', encoding='utf-8') as file:
        lines = [line.strip() for line in file.readlines() if line.strip()]

    marked_lines = []
    i = 0
    while i < len(lines):
        current_line = lines[i]
        if i + 2 < len(lines):
            numbers_line = lines[i + 1]
            next_string_line = lines[i + 2]
            num_match = re.match(r'^(\d+)\s+(-?\d{1,3}(?:\.\d{3})*(?:,\d{2})?)\s+(-?\d{1,3}(?:\.\d{3})*(?:,\d{2})?)$', numbers_line)
            if num_match:
                marked_lines.append(f"*{current_line}")
                marked_lines.append(numbers_line)
                marked_lines.append(f"{next_string_line}*")
                i += 3
                continue
        marked_lines.append(current_line)
        i += 1
    with open(file_path, 'w', encoding='utf-8') as file:
        file.write('\n'.join(marked_lines))

def process_marked_blocks(file_path: str):
    """Combina os blocos marcados em uma única linha."""
    with open(file_path, 'r', encoding='utf-8') as file:
        lines = [line.strip() for line in file.readlines() if line.strip()]
    processed_lines = []
    i = 0
    while i < len(lines):
        current_line = lines[i]
        if current_line.startswith('*') and i + 2 < len(lines) and lines[i+2].endswith('*'):
            line1 = current_line[1:]
            line2 = lines[i+1]
            line3 = lines[i+2][:-1]
            concatenated = f"*{line1} {line3} {line2}*"
            processed_lines.append(concatenated)
            i += 3
        else:
            processed_lines.append(current_line)
            i += 1
    with open(file_path, 'w', encoding='utf-8') as file:
        file.write('\n'.join(processed_lines))

def first_exception(file_path: str):
    """Trata o primeiro caso de formatação excepcional."""
    with open(file_path, 'r', encoding='utf-8') as file:
        lines = [line.strip() for line in file.readlines() if line.strip()]
    processed_lines = []
    i = 0
    while i < len(lines):
        current_line = lines[i]
        if (not current_line.startswith('*')) and (i + 1 < len(lines)) and (not lines[i+1].startswith('*')):
            first_line_match = re.search(r'(\d+)\s+(-?\d{1,3}(?:\.\d{3})*(?:,\d{2})?)\s+(-?\d{1,3}(?:\.\d{3})*(?:,\d{2})?)$', current_line)
            second_line_match = re.search(r'(\d+)\s+(-?\d{1,3}(?:\.\d{3})*(?:,\d{2})?)\s+(-?\d{1,3}(?:\.\d{3})*(?:,\d{2})?)$', lines[i+1])
            if first_line_match and not second_line_match:
                desc_part = re.sub(r'\s+\d+\s+[-+]?\d{1,3}(?:\.\d{3})*(?:,\d{2})?\s+[-+]?\d{1,3}(?:\.\d{3})*(?:,\d{2})?$', '', current_line)
                numbers_part = first_line_match.group(0)
                new_line = f"{desc_part} {lines[i+1]} {numbers_part}"
                processed_lines.append(new_line)
                i += 2
                continue
        processed_lines.append(current_line)
        i += 1
    with open(file_path, 'w', encoding='utf-8') as file:
        file.write('\n'.join(processed_lines))

def second_exception(file_path: str):
    """Trata o segundo caso de formatação excepcional."""
    with open(file_path, 'r', encoding='utf-8') as file:
        lines = [line.strip() for line in file.readlines() if line.strip()]
    processed_lines = []
    i = 0
    while i < len(lines):
        current_line = lines[i]
        if not current_line.startswith('*') and i + 1 < len(lines) and not lines[i+1].startswith('*'):
            first_line_has_numbers = re.search(r'(\d+)\s+(-?\d{1,3}(?:\.\d{3})*(?:,\d{2})?)\s+(-?\d{1,3}(?:\.\d{3})*(?:,\d{2})?)$', current_line)
            second_line_has_numbers = re.search(r'(\d+)\s+(-?\d{1,3}(?:\.\d{3})*(?:,\d{2})?)\s+(-?\d{1,3}(?:\.\d{3})*(?:,\d{2})?)$', lines[i+1])
            if not first_line_has_numbers and second_line_has_numbers:
                concatenated = f"{current_line} {lines[i+1]}"
                processed_lines.append(concatenated)
                i += 2
                continue
        processed_lines.append(current_line)
        i += 1
    with open(file_path, 'w', encoding='utf-8') as file:
        file.write('\n'.join(processed_lines))

def propagate_and_format(file_path: str):
    """Garante que cada linha de transação tenha uma data."""
    with open(file_path, 'r', encoding='utf-8') as file:
        lines = [line.strip() for line in file.readlines() if line.strip()]
    cleaned_lines = [line.strip('*').strip() for line in lines]
    
    processed_lines = []
    current_date = None
    date_pattern = re.compile(r'^(\d{2}/\d{2}/\d{4})')

    for line in cleaned_lines:
        date_match = date_pattern.match(line)
        
        if date_match:
            current_date = date_match.group(1)
            processed_lines.append(line)
        elif current_date:
            processed_lines.append(f"{current_date} {line}")
        else:
            processed_lines.append(line)
            
    with open(file_path, 'w', encoding='utf-8') as file:
        file.write('\n'.join(processed_lines))

def txt_to_xls(input_file: str, output_file: str):
    """Converte o arquivo de texto final para uma planilha XLS."""
    with open(input_file, 'r', encoding='utf-8') as file:
        lines = [line.strip() for line in file.readlines() if line.strip()]

    wb = Workbook()
    ws = wb.add_sheet('Transacoes')
    headers = ['Data', 'Histórico', 'Dcto.', 'Valor', 'Saldo']
    for col, header in enumerate(headers):
        ws.write(0, col, header)

    row = 1
    date_pattern = re.compile(r'(\d{2}/\d{2}/\d{4})')
    
    for line in lines:
        date_match = date_pattern.search(line)
        if not date_match:
            continue

        current_date = date_match.group(1)
        remaining_text = line.replace(current_date, '', 1).strip()
        
        parts = remaining_text.split()
        if len(parts) < 3:
            continue
            
        saldo = parts[-1]
        valor = parts[-2]
        dcto = parts[-3]
        historico = ' '.join(parts[:-3])

        ws.write(row, 0, current_date)
        ws.write(row, 1, historico)
        ws.write(row, 2, dcto)
        ws.write(row, 3, valor)
        ws.write(row, 4, saldo)
        row += 1
        
    ws.col(0).width = 3000
    ws.col(1).width = 12000
    for col in range(2,5):
        ws.col(col).width = 4000
    wb.save(output_file)

# --- FUNÇÃO PRINCIPAL QUE O MENU CHAMA ---
def main():
    """
    Função principal que coordena todo o processo para o Bradesco.
    Retorna True em caso de sucesso e False em caso de cancelamento.
    """
    pdf_path = filedialog.askopenfilename(
        title="Selecione o extrato do Bradesco",
        filetypes=[("PDF files", "*.pdf")]
    )
    
    if not pdf_path:
        return False # Retorna False se o usuário cancelar

    temp_txt_path = ""
    try:
        temp_txt_path = extract_pdf_to_text(pdf_path)
        if not temp_txt_path:
            raise Exception("Falha ao extrair texto do PDF.")

        clean_statement(temp_txt_path)
        mark_all_transaction_blocks(temp_txt_path)
        process_marked_blocks(temp_txt_path)
        first_exception(temp_txt_path)
        second_exception(temp_txt_path)
        propagate_and_format(temp_txt_path)
        
        base_name = os.path.splitext(pdf_path)[0]
        output_xls_path = base_name + ".xls"

        txt_to_xls(temp_txt_path, output_xls_path)
        
        return True # Retorna True para indicar sucesso

    except Exception:
        traceback.print_exc()
        return False # Retorna False se ocorrer qualquer erro
    finally:
        if temp_txt_path and os.path.exists(temp_txt_path):
            try:
                os.remove(temp_txt_path)
            except OSError as e:
                print(f"Erro ao remover arquivo temporário: {e}")

# Bloco para permitir testar este script de forma independente
if __name__ == "__main__":
    main()