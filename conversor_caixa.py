"""
Extrator de dados de extratos bancários em formato PDF.
Este script foi projetado para ser flexível e trabalhar com vários formatos de extratos bancários.
"""

import argparse
import csv
import os
import re
import sys
from datetime import datetime

import pandas as pd
import pdfplumber  # Biblioteca mais robusta para extrair texto de PDFs


def extract_text_from_pdf(pdf_path):
    """
    Extrai o texto completo de um arquivo PDF usando pdfplumber,
    que é mais robusto para extrair textos com formatação estruturada
    """
    all_text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text(x_tolerance=3, y_tolerance=3)
                if text:
                    all_text += text + "\n"
        
        if not all_text.strip():
            print("Aviso: O PDF parece estar vazio ou o texto não pôde ser extraído diretamente.")
            # Tente outra abordagem - extrair tabelas
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        for row in table:
                            all_text += " ".join([cell or "" for cell in row]) + "\n"
        
        return all_text
    except Exception as e:
        print(f"Erro ao extrair texto do PDF: {e}")
        return None


def identify_date_pattern(text):
    """
    Identifica o padrão de data usado no extrato bancário
    Suporta DD/MM/YYYY e outros formatos comuns
    """
    # Verifica o padrão DD/MM/YYYY (mais comum em extratos brasileiros)
    if re.search(r'\d{2}/\d{2}/\d{4}', text):
        return r'\d{2}/\d{2}/\d{4}'
    # Verifica o padrão MM/DD/YYYY
    elif re.search(r'\d{2}/\d{2}/\d{4}', text):
        return r'\d{2}/\d{2}/\d{4}'
    # Verifica o padrão YYYY-MM-DD
    elif re.search(r'\d{4}-\d{2}-\d{2}', text):
        return r'\d{4}-\d{2}-\d{2}'
    # Verifica o padrão YYYY/MM/DD
    elif re.search(r'\d{4}/\d{2}/\d{2}', text):
        return r'\d{4}/\d{2}/\d{2}'
    # Verifica o padrão DD.MM.YYYY
    elif re.search(r'\d{2}\.\d{2}\.\d{4}', text):
        return r'\d{2}\.\d{2}\.\d{4}'
    else:
        return None


def parse_date(date_str, date_pattern):
    """
    Converte a string de data para o formato datetime com base no padrão identificado
    """
    if date_pattern == r'\d{2}/\d{2}/\d{4}':  # DD/MM/YYYY
        return datetime.strptime(date_str, '%d/%m/%Y')
    elif date_pattern == r'\d{2}/\d{2}/\d{4}':  # MM/DD/YYYY
        return datetime.strptime(date_str, '%m/%d/%Y')
    elif date_pattern == r'\d{4}-\d{2}-\d{2}':  # YYYY-MM-DD
        return datetime.strptime(date_str, '%Y-%m-%d')
    elif date_pattern == r'\d{4}/\d{2}/\d{2}':  # YYYY/MM/DD
        return datetime.strptime(date_str, '%Y/%m/%d')
    elif date_pattern == r'\d{2}\.\d{2}\.\d{4}':  # DD.MM.YYYY
        return datetime.strptime(date_str, '%d.%m.%Y')
    else:
        try:
            # Tenta alguns formatos comuns
            for fmt in ('%d/%m/%Y', '%m/%d/%Y', '%Y-%m-%d', '%Y/%m/%d', '%d.%m.%Y'):
                try:
                    return datetime.strptime(date_str, fmt)
                except ValueError:
                    continue
            raise ValueError(f"Não foi possível converter a data: {date_str}")
        except Exception as e:
            print(f"Erro ao converter data '{date_str}': {e}")
            return None


def extract_transactions_from_text(text):
    """
    Extrai transações do texto do extrato bancário usando padrões mais flexíveis
    """
    # Detecta o padrão de data usado no documento
    date_pattern = identify_date_pattern(text)
    if not date_pattern:
        print("Não foi possível identificar o padrão de data no extrato.")
        return []

    transactions = []
    lines = text.split('\n')
    
    # Para diagnóstico, mostra algumas linhas
    print("\nPrimeiras 5 linhas do texto extraído:")
    for i, line in enumerate(lines[:5]):
        print(f"{i+1}: {line}")
    
    # Diferentes padrões para extrair transações
    patterns = [
        # Padrão 1: data, código, descrição, valor, saldo
        fr'({date_pattern})\s+(\d+|\w+)\s+(.*?)\s+([\d\.,]+\s*[DC]?)\s+([\d\.,]+\s*[DC]?)',
        
        # Padrão 2: data, descrição, valor, saldo (sem código)
        fr'({date_pattern})\s+(.*?)\s+([\d\.,]+\s*[DC]?)\s+([\d\.,]+\s*[DC]?)',
        
        # Padrão 3: data no início da linha seguida por qualquer conteúdo
        fr'({date_pattern})(.+)'
    ]
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        # Se a linha não tiver uma data, pule
        if not re.search(date_pattern, line):
            continue
            
        # Tenta cada padrão de extração
        for pattern_index, pattern in enumerate(patterns):
            match = re.search(pattern, line)
            if match:
                try:
                    if pattern_index == 0:  # Padrão 1
                        date = match.group(1)
                        description = match.group(3).strip()
                        # Determina qual grupo tem o saldo
                        if 'C' in match.group(5) or 'D' in match.group(5):
                            balance = match.group(5).strip()
                        else:
                            # Tenta inferir se não há indicador C/D
                            balance = match.group(5).strip()
                    
                    elif pattern_index == 1:  # Padrão 2
                        date = match.group(1)
                        description = match.group(2).strip()
                        balance = match.group(4).strip()
                    
                    elif pattern_index == 2:  # Padrão 3 - extração mais básica
                        date = match.group(1)
                        remaining_text = match.group(2).strip()
                        
                        # Tenta extrair o último número da linha como saldo
                        balance_match = re.search(r'([\d\.,]+\s*[DC]?)\s*$', remaining_text)
                        if balance_match:
                            balance = balance_match.group(1).strip()
                            # O que sobra pode ser a descrição
                            description_parts = remaining_text.rsplit(balance, 1)
                            description = description_parts[0].strip()
                        else:
                            # Se não conseguir extrair o saldo, usa o texto restante como descrição
                            description = remaining_text
                            balance = "Não identificado"
                    
                    # Filtra linhas que são cabeçalhos ou não são transações reais
                    skip_words = ["SALDO DIA", "SALDO ANTERIOR", "TOTAL", "Data Mov", "Histórico", "Valor"]
                    if any(word in description for word in skip_words):
                        continue
                        
                    transactions.append({
                        'Data': date,
                        'Histórico': description,
                        'Saldo': balance
                    })
                    
                    # Se encontrou uma correspondência, não precisa tentar os outros padrões
                    break
                        
                except Exception as e:
                    print(f"Erro ao processar linha '{line}': {e}")
                    continue
    
    print(f"Total de transações encontradas: {len(transactions)}")
    return transactions


def clean_monetary_value(value_str):
    """
    Limpa e converte um valor monetário para formato numérico
    """
    try:
        # Remove caracteres não numéricos, exceto pontos, vírgulas e sinais
        cleaned = re.sub(r'[^\d\.,\-+CD]', '', value_str)
        
        # Determina se é débito ou crédito
        is_debit = 'D' in value_str or '-' in value_str
        
        # Remove indicadores de débito/crédito
        cleaned = cleaned.replace('D', '').replace('C', '')
        
        # Remove pontos de milhar e substitui vírgula por ponto
        if ',' in cleaned and '.' in cleaned:
            # Formato brasileiro (1.234,56)
            cleaned = cleaned.replace('.', '')
            cleaned = cleaned.replace(',', '.')
        elif ',' in cleaned:
            # Vírgula como decimal (1234,56)
            cleaned = cleaned.replace(',', '.')
        
        # Converte para float
        value = float(cleaned)
        
        # Se for débito, torna o valor negativo (se já não for)
        if is_debit and value > 0:
            value = -value
            
        return value
    except Exception as e:
        print(f"Erro ao processar valor monetário '{value_str}': {e}")
        # Retorna None ou 0.0 para valores que não podem ser convertidos
        return None


def process_transactions(transactions, date_pattern):
    """
    Processa a lista de transações e cria um DataFrame
    """
    if not transactions:
        return None
        
    # Cria o DataFrame
    df = pd.DataFrame(transactions)
    
    # Tenta converter datas
    try:
        df['Data'] = df['Data'].apply(lambda x: parse_date(x, date_pattern))
    except Exception as e:
        print(f"Aviso: Não foi possível converter as datas: {e}")
        # Mantém as datas como strings se a conversão falhar
    
    # Tenta converter valores de saldo
    if 'Saldo' in df.columns:
        try:
            # Cria uma cópia da coluna original antes de tentar converter
            df['Saldo_Original'] = df['Saldo']
            df['Saldo'] = df['Saldo'].apply(clean_monetary_value)
            
            # Se houve valores que não puderam ser convertidos (None), restaura o valor original
            mask = df['Saldo'].isna()
            if mask.any():
                print(f"Aviso: {mask.sum()} valores de saldo não puderam ser convertidos.")
                # Opcionalmente: df.loc[mask, 'Saldo'] = df.loc[mask, 'Saldo_Original']
        except Exception as e:
            print(f"Aviso: Não foi possível converter os valores de saldo: {e}")
            # Restaura a coluna original se a conversão falhar completamente
            if 'Saldo_Original' in df.columns:
                df['Saldo'] = df['Saldo_Original']
        finally:
            # Remove a coluna temporária
            if 'Saldo_Original' in df.columns:
                df.drop('Saldo_Original', axis=1, inplace=True)
    
    return df


def main():
    """
    Função principal que coordena a extração e processamento dos dados
    """
    parser = argparse.ArgumentParser(description='Extrai transações de extratos bancários em PDF.')
    parser.add_argument('pdf_path', help='Caminho para o arquivo PDF do extrato bancário')
    parser.add_argument('--output', '-o', help='Nome do arquivo CSV de saída', 
                        default=None)
    parser.add_argument('--verbose', '-v', action='store_true', 
                        help='Exibe informações detalhadas durante o processamento')
    
    args = parser.parse_args()
    
    # Define o arquivo de saída se não for especificado
    if not args.output:
        base_name = os.path.splitext(os.path.basename(args.pdf_path))[0]
        args.output = f"{base_name}_processado.csv"
    
    print(f"Extraindo dados do arquivo: {args.pdf_path}")
    text = extract_text_from_pdf(args.pdf_path)
    
    if not text:
        print("Erro: Não foi possível extrair texto do PDF.")
        sys.exit(1)
    
    if args.verbose:
        print("\n--- Trecho do texto extraído ---")
        print(text[:500])
        print("--- Fim do trecho ---\n")
    
    date_pattern = identify_date_pattern(text)
    if not date_pattern:
        print("Erro: Não foi possível identificar o padrão de data no extrato.")
        sys.exit(1)
    
    print(f"Padrão de data identificado: {date_pattern}")
    transactions = extract_transactions_from_text(text)
    
    if not transactions:
        print("Erro: Nenhuma transação foi encontrada no texto extraído.")
        sys.exit(1)
    
    df = process_transactions(transactions, date_pattern)
    
    if df is None or df.empty:
        print("Erro: Não foi possível processar as transações.")
        sys.exit(1)
    
    # Reorganiza as colunas conforme solicitado
    if 'Data' in df.columns and 'Histórico' in df.columns and 'Saldo' in df.columns:
        df = df[['Data', 'Histórico', 'Saldo']]
    
    # Salva o resultado
    df.to_csv(args.output, index=False, quoting=csv.QUOTE_NONNUMERIC, encoding='utf-8-sig')
    print(f"Arquivo '{args.output}' gerado com sucesso!")
    
    # Mostra as primeiras linhas para verificação
    print("\nPrimeiras 5 linhas do resultado:")
    print(df.head(5).to_string())
    
    print(f"\nTotal de registros processados: {len(df)}")


if __name__ == "__main__":
    main()