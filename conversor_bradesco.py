import os
import re
import pdfplumber
import tkinter as tk
from tkinter import filedialog
from xlwt import Workbook

def extract_pdf_to_text():
    root = tk.Tk()
    root.withdraw()
    pdf_path = filedialog.askopenfilename(
        title="Bradesco",
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
    )
    if not pdf_path:
        print("No file selected. Exiting.")
        return None

    output_path = os.path.join(os.path.dirname(pdf_path), "teste.txt")
    try:
        with pdfplumber.open(pdf_path) as pdf:
            with open(output_path, 'w', encoding='utf-8') as txt_file:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        txt_file.write(text + '\n\n')
        print(f"Successfully extracted text to: {output_path}")
        return output_path
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return None

def clean_statement(file_path: str):
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
    print(f"File '{file_path}' cleaned - removed first 5 lines and everything after 'Total'.")

def mark_all_transaction_blocks(file_path: str):
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
    print(f"File '{file_path}' updated with marked transaction blocks.")

def process_marked_blocks(file_path: str):
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
    print(f"File '{file_path}' processed with wrapped concatenated blocks.")

def first_exception(file_path: str):
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
    print(f"File '{file_path}' processed with unmarked sections reformatted.")

def second_exception(file_path: str):
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
    print(f"File '{file_path}' processed - unmarked sections concatenated.")

def propagate_and_format(file_path: str):
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
    print(f"File '{file_path}' processed - dates propagated correctly.")

# --- FUNÇÃO CORRIGIDA ---
def txt_to_xls(input_file: str, output_file: str):
    with open(input_file, 'r', encoding='utf-8') as file:
        lines = [line.strip() for line in file.readlines() if line.strip()]

    wb = Workbook()
    ws = wb.add_sheet('Transacoes')
    headers = ['Data', 'Histórico', 'Dcto.', 'Valor', 'Saldo']
    for col, header in enumerate(headers):
        ws.write(0, col, header)

    row = 1
    last_valid_date = None
    # Regex para ENCONTRAR uma data em qualquer lugar na linha
    date_pattern = re.compile(r'(\d{2}/\d{2}/\d{4})')
    
    for line in lines:
        current_line_date = ''
        remaining_text = line
        
        # Procura pelo padrão de data na linha
        match = date_pattern.search(line)
        
        if match:
            # Se encontrou, extrai a data
            current_line_date = match.group(1)
            last_valid_date = current_line_date
            # Remove a data da linha para criar o histórico, substituindo apenas a primeira ocorrência
            remaining_text = line.replace(current_line_date, '', 1).strip()
        elif last_valid_date:
            # Se não encontrou, usa a última data válida e o texto da linha inteira como histórico
            current_line_date = last_valid_date
            remaining_text = line

        # O resto da lógica opera sobre o 'remaining_text' já sem a data
        final_parts = remaining_text.split()
        valores = []
        for part in final_parts[-3:]:
            if re.match(r'^-?[\d.,]+$', part):
                clean_val = part.replace('.', '').replace(',', '.')
                try:
                    valores.append(float(clean_val) if '.' in clean_val else int(clean_val))
                except ValueError:
                    valores.append(part)
            else:
                valores.append(part)
                
        historico_parts = final_parts[:-3] if len(final_parts) > 3 else []
        historico = ' '.join(historico_parts)

        ws.write(row, 0, current_line_date)
        ws.write(row, 1, historico)
        for col in range(3):
            val = valores[col] if col < len(valores) else ''
            ws.write(row, col+2, val)
        row += 1
        
    ws.col(0).width = 3000
    ws.col(1).width = 12000
    for col in range(2,5):
        ws.col(col).width = 4000
    wb.save(output_file)
    print(f'Successfully exported to {output_file}')
# --- FIM DA FUNÇÃO CORRIGIDA ---

def main():
    file_path = extract_pdf_to_text()
    if not file_path:
        return
    clean_statement(file_path)
    mark_all_transaction_blocks(file_path)
    process_marked_blocks(file_path)
    first_exception(file_path)
    second_exception(file_path)
    propagate_and_format(file_path)
    txt_to_xls(file_path, os.path.join(os.path.dirname(file_path), 'transacoes.xls'))
    
if __name__ == "__main__":
    main()
