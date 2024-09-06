import os
import camelot
import pandas as pd
import re
from unidecode import unidecode
from tkinter import Tk, simpledialog
from tkinter.filedialog import askopenfilename

class PDFTableExtractor:
    def __init__(self, file_path, configs):
        self.path = file_path
        self.csv_path = os.path.dirname(file_path)
        self.configs = configs

    def start(self):
        pages_with_tables = simpledialog.askstring("Extrair Informações", "Digite as páginas que contêm tabelas (ex: 1,2,4-6):")
        if not pages_with_tables or not pages_with_tables.strip():
            raise ValueError("Nenhuma página válida foi especificada.")
        
        page_numbers = self.parse_pages(pages_with_tables)
        header = pd.DataFrame()

        if '1' in page_numbers:
            header = self.get_table_data('page_1', '1')
            page_numbers.remove('1')
            header_csv_path = os.path.join(self.csv_path, "header_page_1.csv")
            header.to_csv(header_csv_path, sep=";", index=False, encoding='utf-8')
        
        main = pd.DataFrame()
        for i in range(0, len(page_numbers), 5):
            pages_block = page_numbers[i:i+5]
            block_data = self.get_table_data('page_2_end', ','.join(pages_block))
            main = pd.concat([main, block_data], ignore_index=True)
        
        if not main.empty and not header.empty:
            main = self.add_infos(header, main)
            main = self.sanitize_column_names(main)
            main = self.fill_empty_column(main, '144,262, 204,262, 157,757, 173,757')
            final_csv_path = self.save_csv(main)
            self.combine_and_align_csvs(header_csv_path, final_csv_path)

    def parse_pages(self, pages):
        page_numbers = []
        for part in pages.split(','):
            part = part.strip()
            if '-' in part:
                start, end = map(int, part.split('-'))
                page_numbers.extend(map(str, range(start, end + 1)))
            else:
                page_numbers.append(part)
        return page_numbers

    def get_table_data(self, page_key, pages):
        config = self.configs.get(page_key, {})
        tables = camelot.read_pdf(
            self.path,
            flavor=self.configs.get('flavor', 'stream'),
            table_areas=config.get('table_areas'),
            columns=config.get('columns'),
            pages=pages,
            strip_text=config.get('strip_text', '')
        )

        table_data = [self.fix_header(table.df) for table in tables if not table.df.empty]
        return pd.concat(table_data, ignore_index=True) if table_data else pd.DataFrame()

    def save_csv(self, df):
        base_name = os.path.splitext(os.path.basename(self.path))[0]
        path = os.path.join(self.csv_path, f"{base_name}.csv")
        df.to_csv(path, sep=";", index=False, encoding='utf-8')
        return path

    def add_infos(self, header, content):
        infos = header.iloc[0]
        df = pd.DataFrame([infos.values] * len(content), columns=header.columns)
        content = pd.concat([content.reset_index(drop=True), df.reset_index(drop=True)], axis=1)
        return content

    @staticmethod
    def fix_header(df):
        df.columns = df.iloc[0]
        df = df.drop(0).reset_index(drop=True)
        return df

    @staticmethod
    def sanitize_column_names(df):
        df.columns = df.columns.map(lambda x: unidecode(x))  
        df.columns = df.columns.map(lambda x: re.sub(r'[^\w\s]', '', x))  
        df.columns = df.columns.map(lambda x: x.replace(' ', '_'))  
        df.columns = df.columns.map(lambda x: x.lower())  
        
        df = df.loc[:, ~df.columns.str.contains(r'^Unnamed:\s*\d+', regex=True)]
        if 'data_de_insercao' in df.columns:
            df = df.drop('data_de_insercao', axis=1)
        df = df.dropna(axis=1, how='all')
        return df

    @staticmethod
    def fill_empty_column(df, column_name):
        if column_name not in df.columns:
            return df
        
        for index, row in df.iterrows():
            if pd.isna(row[column_name]):
                last_valid_index = row.last_valid_index()
                if last_valid_index and last_valid_index != column_name:
                    df.at[index, column_name] = row[last_valid_index]
        return df

    def combine_and_align_csvs(self, header_csv, final_csv):
        # Leia os arquivos CSV
        header_df = pd.read_csv(header_csv, sep=";", encoding='utf-8')
        final_df = pd.read_csv(final_csv, sep=";", encoding='utf-8')

        # Combine os DataFrames
        combined_df = pd.concat([header_df, final_df], ignore_index=True)

        # Remove qualquer coluna completamente vazia
        combined_df = combined_df.dropna(axis=1, how='all')

        # Define o caminho do CSV final alinhado
        final_aligned_path = os.path.join(self.csv_path, "extratoconsolidado.csv")

        # Corrige o problema de múltiplos delimitadores consecutivos
        combined_df.to_csv(final_aligned_path, sep=";", index=False, encoding='utf-8', na_rep='')

        # Remover múltiplos delimitadores consecutivos do arquivo CSV resultante
        with open(final_aligned_path, 'r', encoding='utf-8') as file:
            csv_content = file.read()

        # Substitui múltiplos delimitadores `;` consecutivos por um único delimitador
        csv_content = re.sub(r';{2,}', ';', csv_content)

        # Salva o conteúdo corrigido de volta ao arquivo
        with open(final_aligned_path, 'w', encoding='utf-8') as file:
            file.write(csv_content)

if __name__ == "__main__":
    root = Tk()
    root.withdraw()
    file_path = askopenfilename(filetypes=[("PDF files", "*.pdf")])
    root.destroy()

    if file_path:
        configs = {
            'flavor': 'stream',
            'page_1': {
                'table_areas': ['149,257, 552,21'],
                'columns': ['144,262, 204,262, 303,262, 351,262, 406,262, 418,262, 467,262, 506,262, 553,262'],
                'fix': True
            },
            'page_2_end': {
                'table_areas': ['151,760, 553,20'],
                'columns': ['157,757, 173,757, 269,757, 309,757, 363,757, 380,757, 470,757, 509,757, 545,757'],
                'fix': True
            }
        }
        
        extractor = PDFTableExtractor(file_path, configs)
        extractor.start()
