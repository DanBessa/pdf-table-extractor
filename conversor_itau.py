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
        main = pd.DataFrame()
        if '1' in page_numbers:
            header = self.get_table_data('page_1', '1')
            page_numbers.remove('1')

            if not header.empty:
                header = self.clean_data(header)
                main = pd.concat([main, header], ignore_index=True)

        for i in range(0, len(page_numbers), 5):
            pages_block = page_numbers[i:i+5]
            block_data = self.get_table_data('page_2_end', ','.join(pages_block))

            if not block_data.empty:
                block_data = self.clean_data(block_data)
                self.debug_dataframes(main, block_data)
                main = pd.concat([main, block_data], ignore_index=True)

        if not main.empty:
            main = self.sanitize_column_names(main)
            main = self.fill_empty_dates(main, 'data')
            main = self.remove_credit_debit_repeats(main)

            final_csv_path = self.save_csv(main)
            self.finalize_csv(final_csv_path)

    def clean_data(self, df):
        df = df.reset_index(drop=True)  # Garante que os índices sejam únicos
        df = df.loc[:, ~df.columns.duplicated()]  # Remove colunas duplicadas
        df = df.dropna(axis=1, how='all')  # Remove colunas vazias
        df.columns = df.columns.str.strip()  # Remove espaços extras dos nomes das colunas
        
        # Adiciona mais algumas limpezas específicas para as colunas
        df = df.reset_index(drop=True)
        df = df.loc[:, ~df.columns.duplicated()]
        df = df.dropna(axis=1, how='all')
        df.columns = df.columns.str.strip()

        if 'data' in df.columns:
            df['data'] = df['data'].str.strip()  # Remove espaços extras dos valores das colunas
            df['data'] = df['data'].str.strip()

        for column in df.columns:
            df[column] = df[column].apply(self.fix_hyphen)

        return df

    def fix_hyphen(self, value):
        if isinstance(value, str):
            value = value.strip()
            value = value.replace(".", "")
            value = re.sub(r'(\d+),(\d+)-$', r'-\1,\2', value)
        return value

    def debug_dataframes(self, df1, df2):
        print("DataFrame principal:")
        print("Colunas:", df1.columns)
        print("Índices:", df1.index)
        print("Exemplo de dados:")
        print(df1.head())

        print("\nDataFrame do bloco:")
        print("Colunas:", df2.columns)
        print("Índices:", df2.index)
        print("Exemplo de dados:")
        print(df2.head())

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

        table_data = [self.fix_header(table.df).reset_index(drop=True) for table in tables if not table.df.empty]
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
        df.columns = df.columns.map(lambda x: unidecode(str(x)))  
        df.columns = df.columns.map(lambda x: re.sub(r'[^\w\s]', '', x))  
        df.columns = df.columns.map(lambda x: x.replace(' ', '_'))  
        df.columns = df.columns.map(lambda x: x.lower())  

        df = df.loc[:, ~df.columns.duplicated()]  # Remove colunas duplicadas
        df = df.loc[:, ~df.columns.duplicated()]
        df = df.loc[:, ~df.columns.str.contains(r'^Unnamed:\s*\d+', regex=True)]
        if 'data_de_insercao' in df.columns:
            df = df.drop('data_de_insercao', axis=1)
        df = df.dropna(axis=1, how='all')
        return df

    def fill_empty_dates(self, df, date_column_name):
        if date_column_name in df.columns:
            df[date_column_name] = df[date_column_name].replace('', pd.NA)  # Substitui strings vazias por NA
            df[date_column_name] = df[date_column_name].ffill()  # Preenche valores vazios com a última data válida
            df[date_column_name] = df[date_column_name].replace('', pd.NA)
            df[date_column_name] = df[date_column_name].ffill()
        return df

    def remove_credit_debit_repeats(self, df):
        credit_column = 'credito'
        debit_column = 'debito'
        history_column = 'historico'

        if credit_column in df.columns:
            df[credit_column] = df[credit_column].replace('', pd.NA)
            df[credit_column] = df[credit_column].bfill()

        if debit_column in df.columns:
            df[debit_column] = df[debit_column].replace('', pd.NA)
            df[debit_column] = df[debit_column].bfill()

        return df

    def finalize_csv(self, final_csv):
        final_aligned_path = os.path.join(self.csv_path, "extratoconvertido.csv")

        with open(final_csv, 'r', encoding='utf-8') as file:
            csv_content = file.read()

        csv_content = re.sub(r';{2,}', ';', csv_content)

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
                'strip_text': ''
            },
            'page_2_end': {
                'table_areas': ['151,760, 553,20'],
                'columns': ['157,757, 173,757, 269,757, 309,757, 363,757, 380,757, 470,757, 509,757, 545,757'],
                'strip_text': ''
            }
        }

        extractor = PDFTableExtractor(file_path, configs)
        extractor.start()
