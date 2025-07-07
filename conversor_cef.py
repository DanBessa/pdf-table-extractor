# conversor_cef.py

import csv
import re
import io
import pdfplumber # Biblioteca para ler arquivos PDF
import tkinter as tk
from tkinter import filedialog
import os # Módulo para manipulação de caminhos de arquivo

def clean_field(field_text):
    """
    Limpa um único campo: remove espaços extras no início/fim e 
    substitui múltiplos espaços internos por um único espaço.
    """
    if field_text is None:
        return ""
    cleaned_text = re.sub(r'\s+', ' ', field_text)
    return cleaned_text.strip()

def extract_text_from_pdf(pdf_path):
    """
    Extrai texto de todas as páginas de um arquivo PDF.
    Retorna uma única string com todo o texto do PDF.
    """
    full_text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for i, page in enumerate(pdf.pages):
                page_text = page.extract_text()
                if page_text:
                    full_text += page_text + f"\n--- PAGE {i+1} ---\n"
        if not full_text:
            # Este print aparecerá no console. O app principal mostrará sua própria messagebox.
            print(f"Aviso (conversor_cef): Nenhum texto foi extraído do PDF: {pdf_path}.")
    except Exception as e:
        # Este print aparecerá no console. A exceção será propagada para o app principal.
        print(f"Erro (conversor_cef) ao ler o arquivo PDF {pdf_path}: {e}")
        return None # Retorna None para indicar falha na extração
    return full_text

def parse_and_extract_transactions_revised(text_content, csv_file_path):
    """
    Analisa o conteúdo de texto para encontrar tabelas de transações
    e as grava em um arquivo CSV, usando ';' como delimitador.
    """
    all_data_rows = []
    expected_header_text_line = "Data Mov. Nr. Doc. Histórico Valor Saldo"
    csv_column_headers = ["Data Mov.", "Nr. Doc.", "Histórico", "Valor", "Saldo"]
    header_added_to_csv = False
    
    lines = text_content.splitlines()
    is_capturing_transaction_data = False
    processed_successfully = False # Flag para indicar sucesso

    for line in lines:
        cleaned_line_content = clean_field(line)

        if not cleaned_line_content:
            continue

        if cleaned_line_content == expected_header_text_line:
            is_capturing_transaction_data = True
            if not header_added_to_csv:
                all_data_rows.append(csv_column_headers)
                header_added_to_csv = True
            continue

        if "--- PAGE" in line or \
           cleaned_line_content.startswith("Lançamentos do Dia") or \
           cleaned_line_content.startswith("SAC CAIXA"):
            is_capturing_transaction_data = False
            continue
            
        if not is_capturing_transaction_data:
            continue

        parts = cleaned_line_content.split()
        
        if len(parts) >= 6: 
            try:
                data_mov = parts[0]
                if not re.match(r'\d{2}/\d{2}/\d{4}', data_mov):
                    is_capturing_transaction_data = False
                    continue

                nr_doc = parts[1]
                saldo_letra = parts[-1]
                saldo_valor = parts[-2]
                valor_letra = parts[-3]
                valor_valor = parts[-4]
                historico_parts_list = parts[2:-4]
                historico = " ".join(historico_parts_list)
                valor_completo_csv = f"{valor_valor} {valor_letra}"
                saldo_completo_csv = f"{saldo_valor} {saldo_letra}"
                
                all_data_rows.append([data_mov, nr_doc, historico, valor_completo_csv, saldo_completo_csv])
            except IndexError:
                is_capturing_transaction_data = False
                continue
    
    if not header_added_to_csv or len(all_data_rows) <= 1:
        # Este print aparecerá no console. O app principal mostrará sua própria messagebox.
        print("Aviso (conversor_cef): Nenhuma tabela de transações com o cabeçalho esperado foi processada ou nenhum dado de transação foi encontrado.")
        return False # Indica falha

    try:
        with open(csv_file_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';') 
            writer.writerows(all_data_rows)
        # Este print aparecerá no console. O app principal mostrará sua própria messagebox de sucesso.
        print(f"Info (conversor_cef): Dados extraídos com sucesso para o arquivo: {csv_file_path}")
        processed_successfully = True
    except IOError as e:
        # Este print aparecerá no console. A exceção será propagada.
        print(f"Erro Crítico (conversor_cef): Não foi possível escrever no arquivo CSV: {csv_file_path}. Erro: {e}")
        # Re-raise a exceção para ser capturada pelo try-except do aplicativo principal
        raise IOError(f"Não foi possível escrever no arquivo CSV: {csv_file_path}. Detalhes: {e}") from e
    
    return processed_successfully

def iniciar_processamento_cef(): # Nome da função principal que o seu app chamará
    """
    Função principal para o conversor CEF: seleciona o PDF e processa.
    """
    # Tkinter root para o filedialog (pode ser omitido se o app principal já tiver um root)
    # No entanto, para manter o módulo mais independente, criamos um temporário.
    root_fd = tk.Tk()
    root_fd.withdraw() # Esconde a janela root do Tkinter

    pdf_input_path = filedialog.askopenfilename(
        parent=root_fd, # Garante que o diálogo seja modal em relação a este root temporário
        title="Selecione o arquivo PDF do extrato da Caixa Econômica Federal",
        filetypes=(("Arquivos PDF", "*.pdf"), ("Todos os arquivos", "*.*"))
    )
    root_fd.destroy() # Destrói o root temporário após o uso

    if not pdf_input_path:
        print("Info (conversor_cef): Nenhum arquivo PDF da CEF foi selecionado.")
        # Não precisa de messagebox aqui, o app principal pode lidar com a falta de ação
        # ou a ausência de um erro propagado.
        # Para que o app principal mostre um erro, podemos levantar uma exceção:
        # raise FileNotFoundError("Nenhum arquivo PDF selecionado.")
        return # Simplesmente retorna se nenhum arquivo for selecionado

    print(f"Info (conversor_cef): Arquivo PDF da CEF selecionado: {pdf_input_path}")
        
    pdf_directory = os.path.dirname(pdf_input_path)
    pdf_basename_no_ext = os.path.splitext(os.path.basename(pdf_input_path))[0]
    csv_output_full_path = os.path.join(pdf_directory, pdf_basename_no_ext + ".csv")
            
    # Mantendo o arquivo de debug para consistência, pode ser útil
    debug_text_file_path = os.path.join(pdf_directory, f"{pdf_basename_no_ext}_extracted_text_debug_cef.txt")

    print(f"Info (conversor_cef): Arquivo CSV da CEF será salvo como: {csv_output_full_path}")
    print(f"Info (conversor_cef): Iniciando processamento do PDF da CEF: {pdf_input_path}")
            
    extracted_text_content = extract_text_from_pdf(pdf_input_path)
            
    if extracted_text_content:
        try:
            with open(debug_text_file_path, "w", encoding="utf-8") as f_debug:
                f_debug.write(extracted_text_content)
            print(f"Info (conversor_cef): (Texto da CEF extraído salvo para depuração em: {debug_text_file_path})")
        except Exception as e_debug:
            print(f"Aviso (conversor_cef): (Erro ao salvar arquivo de debug da CEF: {e_debug})")

        # Chama a função de parse e verifica o resultado
        success = parse_and_extract_transactions_revised(extracted_text_content, csv_output_full_path)
        if not success:
            # Se parse_and_extract_transactions_revised retornou False (falha),
            # podemos levantar uma exceção para ser capturada pelo app principal e mostrar um erro.
            raise Exception("Falha ao processar os dados da tabela no extrato da CEF.")
        # Se chegou aqui, a função parse_and_extract... já imprimiu sucesso ou levantou IOError
    else:
        # extract_text_from_pdf já imprimiu um aviso ou erro.
        # Levanta uma exceção para ser capturada pelo app principal.
        raise Exception("Não foi possível extrair texto do PDF da CEF ou o arquivo está vazio.")

# O if __name__ == '__main__': permite testar este módulo de forma independente
if __name__ == '__main__':
    print("Rodando conversor_cef.py de forma autônoma para teste...")
    try:
        iniciar_processamento_cef()
        print("\nTeste autônomo de conversor_cef.py concluído com sucesso (verifique o console e o arquivo gerado).")
    except FileNotFoundError as e: # Exceção específica para arquivo não selecionado
         print(f"Teste autônomo encerrado: {e}")
    except Exception as e:
        print(f"\nOcorreu um erro durante o teste autônomo de conversor_cef.py: {e}")