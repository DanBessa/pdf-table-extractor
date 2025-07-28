# conversor_ofx.py (CORRIGIDO PARA C6 BANK)

import os
from ofxparse import OfxParser
from openpyxl import Workbook
from tkinter import filedialog, messagebox
import re
import traceback

def processar_ofx(*args):
    """
    Função principal para converter arquivos OFX em uma planilha Excel,
    com pré-processamento para corrigir formatos de número (ponto decimal).
    """
    caminhos_dos_arquivos = filedialog.askopenfilenames(
        title="Selecione os arquivos OFX para conversão",
        filetypes=[("Arquivos OFX", "*.ofx"), ("Todos os arquivos", "*.*")]
    )

    if not caminhos_dos_arquivos:
        # Usamos UserWarning para que o menu principal trate como cancelamento
        raise UserWarning("Nenhum arquivo selecionado.")

    try:
        pasta_trabalho = Workbook()
        planilha_padrao = pasta_trabalho.active
        sheets_criadas = 0

        for caminho_arquivo in caminhos_dos_arquivos:
            nome_arquivo = os.path.basename(caminho_arquivo)
            
            # --- ETAPA DE CORREÇÃO ---
            # 1. Lê todo o conteúdo do arquivo para uma string
            with open(caminho_arquivo, 'r', encoding='utf-8', errors='ignore') as file_handle:
                conteudo_original = file_handle.read()

            # 2. Usa uma expressão regular para encontrar e corrigir os valores monetários
            # Procura por <TRNAMT> seguido de um número com ponto decimal e o substitui por vírgula
            conteudo_corrigido = re.sub(r'(<TRNAMT>)(-?[\d]+)\.(\d{2})', r'\1\2,\3', conteudo_original)
            
            # 3. Passa o conteúdo corrigido para a biblioteca ofxparse
            ofx = OfxParser.parse(conteudo_corrigido)
            
            nome_planilha = os.path.splitext(nome_arquivo)[0][:31]
            planilha = pasta_trabalho.create_sheet(title=nome_planilha)
            sheets_criadas += 1
            planilha.append(['Data', 'Descrição', 'Valor'])
            
            for transacao in ofx.account.statement.transactions:
                planilha.append([transacao.date.strftime('%d/%m/%Y'), transacao.memo, transacao.amount])
        
        if sheets_criadas > 0:
            pasta_trabalho.remove(planilha_padrao)
            caminho_salvamento = os.path.splitext(caminhos_dos_arquivos[0])[0] + ".xlsx"
            pasta_trabalho.save(caminho_salvamento)
            # A mensagem de sucesso agora é controlada pelo menu principal
        else:
            raise Exception("Nenhum arquivo OFX válido foi processado.")

    except Exception as e:
        # Propaga o erro para o menu principal, que mostrará a janela de erro detalhada
        traceback.print_exc() # Imprime o erro no console para depuração
        raise e