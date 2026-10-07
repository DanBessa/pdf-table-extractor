# menuprincipal.py (VERSÃO FINAL com correção de CACHE)

import os
import sys
import importlib
import customtkinter as ctk
from tkinter import messagebox, filedialog
from PIL import Image
from customtkinter import CTkImage
import traceback
import json
import base64
import datetime
import requests
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.exceptions import InvalidSignature

# Adiciona a subpasta 'conversores' ao caminho do Python
try:
    base_path = sys._MEIPASS
except AttributeError:
    base_path = os.path.dirname(os.path.abspath(__file__))

conversores_path = os.path.join(base_path, 'conversores')
if conversores_path not in sys.path:
    sys.path.insert(0, conversores_path)

# --- IDENTIDADE VISUAL ---
COLORS = {
    "background": "#1A1B26", "frame": "#2A2D3A", "accent": "#275555", "hover": "#237641",
    "text": "#F0F0F0", "disabled": "#4A4D5A", "success": "#2ECC71", "warning": "#F39C12", "error": "#E74C3C",
}
FONTS = {"title": ("Segoe UI", 30, "bold"), "button": ("Segoe UI", 12, "bold"), "status": ("Segoe UI", 12)}

# --- DICIONÁRIO DE CONFIGURAÇÃO DOS CONVERSORES ---
CONVERTERS = {
    "bb": {
        "nome": "Banco do Brasil", "icons": "bblogo.png", "aba": "pdf", "type": "model_choice",
        "model_config": {
            "titulo": "Seleção de Modelo BB",
            "label": "Selecione o modelo do extrato do Banco do Brasil:",
            "opcoes": {"modelo1": "Modelo 1 (Com Cabeçalho)", "modelo2": "Modelo 2"}
        }
    },
    "inter": { "nome": "Inter", "icons": "inter.png", "aba": "pdf", "type": "single_file", "module": "conversor_inter", "function": "iniciar_processamento" },
    "itau": { "nome": "Itaú", "icons": "itaulogo.png", "aba": "pdf", "type": "itau_special" },
    "sicoob": {
        "nome": "Sicoob", "icons": "sicoob2.png", "aba": "pdf", "type": "model_choice",
        "model_config": {
            "titulo": "Seleção de Modelo Sicoob",
            "label": "Selecione o modelo do extrato do Sicoob:",
            "opcoes": {"modelo1": "Modelo 1", "modelo2": "Modelo 2 (Quebras)"}
        }
    },
    "bradesco": { "nome": "Bradesco", "icons": "bradesco.png", "aba": "pdf", "type": "simple_run", "module": "conversor_bradesco", "function": "main" },
    "pagbank": { "nome": "PagBank", "icons": "pagbank.png", "aba": "pdf", "type": "multi_file", "module": "conversor_pagbank" },
    "santander": { "nome": "Santander", "icons": "santander.png", "aba": "pdf", "type": "simple_run", "module": "conversor_santander", "function": "iniciar_extracao_santander" },
    "cef": { "nome": "Caixa Econômica", "icons": "cef.png", "aba": "pdf", "type": "simple_run", "module": "conversor_cef", "function": "main" },
    "c6": { "nome": "C6 Bank", "icons": "c6logo.png", "aba": "pdf", "type": "simple_run", "module": "conversor_c6", "function": "iniciar_processamento" },
    "banestes": { "nome": "Banestes", "icons": "banestes.png", "aba": "pdf", "type": "simple_run", "module": "conversor_banestes", "function": "iniciar_processamento" },
    "paycash": { "nome": "PayCash", "icons": "paycash.png", "aba": "pdf", "type": "simple_run", "module": "conversor_paycash", "function": "iniciar_processamento" },
    "safra": { "nome": "Safra", "icons": "safra.png", "aba": "pdf", "type": "simple_run", "module": "conversor_safra", "function": "iniciar_processamento" },
    "ofx": { "nome": "Converter Arquivo(s) OFX para Excel", "icons": None, "aba": "ofx", "type": "ofx" },
}
# --- COLE O CONTEÚDO DA SUA CHAVE PÚBLICA AQUI ---
CHAVE_PUBLICA_PEM = """
-----BEGIN PUBLIC KEY-----
MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAlH4rwrpr4hdIlStJxsg+
I6Y8I/O19GmifJYGXBOBMLlRnRpndf9qMqZUhR2wq8Xq3s4kP6JGRLPp0paWlzJ6
veMbME1odu8adtfPsgt4mrPlxVdyb5FEmpS56O/fl+t0FlsIAH9cHKBWZ7wRI2cr
eScMJVyCHlUEYjAYH+nQkXJaOj7/G3QEAUv3FhJsiRbXw0sdg8sY14QsHh/1Fjw/
K+DACi0a3oLYqr/M7iPQHFP7oRcZa0mcMNuqcic6VLFS90cQSrv59otqnzSs5B0V
XO/3tee2cRPqZHaWqv6R2uUR8SSpuO2H4xYQDiEfZHFrR0PgGIjD043IfwbHCXlS
BwIDAQAB
-----END PUBLIC KEY-----
"""
# ---------------------------------------------------
# --- COLOQUE APENAS O ID DO SEU GIST AQUI ---
GIST_ID = "88012cdb74eaad2cf9dee06fe465d490"
# ---------------------------------------------------

# --- FUNÇÕES DE VERIFICAÇÃO DE LICENÇA ---
def carregar_chave_publica():
    return serialization.load_pem_public_key(CHAVE_PUBLICA_PEM.encode('utf-8'))

def verificar_licenca_online():
    """Verifica o status da licença usando a API do GitHub para evitar cache."""
    try:
        url_api = f"https://api.github.com/gists/{GIST_ID}"
        print(f"Buscando URL da API: {url_api}")
        
        response = requests.get(url_api, timeout=5)
        response.raise_for_status()
        
        gist_data = response.json()
        
        conteudo_str = gist_data['files']['config_licenca.json']['content']
        print(f"Conteúdo bruto recebido da API: {conteudo_str}")
        
        config = json.loads(conteudo_str)
        print(f"JSON decodificado: {config}")
        
        return config.get("licenca_obrigatoria", False)
        
    except requests.exceptions.RequestException as e:
        print(f"Aviso: Não foi possível verificar a licença online: {e}")
        return False # Falha "aberta" em caso de erro de rede
    except (KeyError, json.JSONDecodeError) as e:
        print(f"Erro ao processar a resposta da API: {e}")
        return False # Falha "aberta" se a resposta for inesperada
            
    except Exception as e:
        print(f"ERRO GERAL na verificação online: {e}")
        return False

def verificar_licenca_local(caminho_licenca):
    try:
        with open(caminho_licenca, 'r') as f:
            chave_ativacao = f.read()
        pacote_decodificado = base64.b64decode(chave_ativacao)
        pacote = json.loads(pacote_decodificado)
        mensagem = base64.b64decode(pacote['dados'])
        assinatura = base64.b64decode(pacote['assinatura'])
        public_key = serialization.load_pem_public_key(CHAVE_PUBLICA_PEM.encode('utf-8'))
        public_key.verify(assinatura, mensagem, padding.PSS(mgf=padding.MGF1(hashes.SHA256()), salt_length=padding.PSS.MAX_LENGTH), hashes.SHA256())
        dados_licenca = json.loads(mensagem)
        data_expiracao = datetime.date.fromisoformat(dados_licenca['expira_em'])
        if datetime.date.today() > data_expiracao:
            messagebox.showerror("Licença Expirada", f"Sua licença expirou em {data_expiracao.strftime('%d/%m/%Y')}.")
            return False
        return True
    except FileNotFoundError:
        return None
    except Exception:
        return False

def pedir_e_ativar_chave(caminho_licenca):
    chave = ctk.CTkInputDialog(text="Por favor, insira sua chave de ativação:", title="Ativação do Programa").get_input()
    if not chave:
        return False
    try:
        pacote_decodificado = base64.b64decode(chave)
        pacote = json.loads(pacote_decodificado)
        mensagem = base64.b64decode(pacote['dados'])
        assinatura = base64.b64decode(pacote['assinatura'])
        
        public_key = carregar_chave_publica()
        public_key.verify(assinatura, mensagem, padding.PSS(mgf=padding.MGF1(hashes.SHA256()), salt_length=padding.PSS.MAX_LENGTH), hashes.SHA256())
        
        dados_licenca = json.loads(mensagem)
        data_expiracao = datetime.date.fromisoformat(dados_licenca['expira_em'])

        if datetime.date.today() > data_expiracao:
            messagebox.showerror("Chave Inválida", "A chave de ativação inserida já está expirada.")
            return False

        with open(caminho_licenca, 'w') as f:
            f.write(chave)
        
        messagebox.showinfo("Sucesso", f"Programa ativado com sucesso! Sua licença é válida até {data_expiracao.strftime('%d/%m/%Y')}.")
        return True
    except Exception as e:
        messagebox.showerror("Chave Inválida", f"A chave de ativação inserida é inválida ou está corrompida.\n\nDetalhes: {e}")
        return False

# --- CLASSE DE BOTÃO PARA ESTILO CONSISTENTE ---
class ModernButton(ctk.CTkButton):
    def __init__(self, master, **kwargs):
        anchor = kwargs.pop('anchor', 'w')
        fg_color = kwargs.pop('fg_color', COLORS["accent"])
        hover_color = kwargs.pop('hover_color', COLORS["hover"])
        super().__init__(master=master, font=FONTS["button"], fg_color=fg_color, hover_color=hover_color,
                         text_color=COLORS["text"], height=55, corner_radius=10, anchor=anchor,
                         border_spacing=10, compound="left", cursor="hand2", **kwargs)

# --- CLASSE PRINCIPAL DA APLICAÇÃO ---
class ConversorApp(ctk.CTk):
    def __init__(self):
        super().__init__(fg_color=COLORS["background"])
        
        self.base_path = self._get_base_path()
        caminho_licenca = os.path.join(self.base_path, 'licenca.dat')
        
        status_local = verificar_licenca_local(caminho_licenca)
        
        if status_local is True:
            pass
        elif status_local is False:
            self.destroy()
            return
        elif status_local is None:
            if verificar_licenca_online():
                if not pedir_e_ativar_chave(caminho_licenca):
                    self.destroy()
                    return
        
        self.title("Conversor Bancário")
        self.geometry("550x650")
        self.resizable(False, False)
        
        self.icons = self._load_icons()
        self._create_widgets()
    
    def _get_base_path(self):
        try: return sys._MEIPASS
        except AttributeError: return os.path.dirname(os.path.abspath(__file__))

    def _load_icons(self):
        icons = {}
        pasta_icons = os.path.join(self.base_path, "icons")
        
        for key, config in CONVERTERS.items():
            if not config.get("icons"): continue
            try:
                path = os.path.join(pasta_icons, config['icons'])
                image = Image.open(path)
                icons[key] = CTkImage(dark_image=image, light_image=image, size=(28, 28))
            except Exception as e: print(f"Erro ao carregar ícone para '{key}': {e}")
        return icons

    # ... (O restante do seu código da classe e do programa principal continua o mesmo) ...

    def _create_widgets(self):
        container = ctk.CTkFrame(self, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=3, pady=3)
        titulo = ctk.CTkLabel(container, text="Conversor Bancário", font=FONTS["title"], text_color=COLORS["text"])
        titulo.pack(pady=(10, 25))
        tabs = ctk.CTkTabview(container, fg_color=COLORS["frame"], segmented_button_fg_color=COLORS["frame"],
                              segmented_button_selected_color=COLORS["accent"], segmented_button_selected_hover_color=COLORS["hover"],
                              segmented_button_unselected_color=COLORS["frame"], text_color=COLORS["text"],
                              border_width=2, border_color=COLORS["accent"])
        tabs.pack(fill="both", expand=True)
        tab_pdf = tabs.add("PDF")
        tab_ofx = tabs.add("OFX")
        self.frame_botoes_pdf = ctk.CTkFrame(tab_pdf, fg_color="transparent")
        self.frame_botoes_pdf.pack(pady=20, padx=20, fill="both", expand=True)
        self.frame_botoes_ofx = ctk.CTkFrame(tab_ofx, fg_color="transparent")
        self.frame_botoes_ofx.pack(pady=20, padx=20, fill="both", expand=True)

        pdf_buttons = {k: v for k, v in CONVERTERS.items() if v['aba'] == 'pdf'}
        ofx_buttons = {k: v for k, v in CONVERTERS.items() if v['aba'] == 'ofx'}

        for i, (key, config) in enumerate(pdf_buttons.items()):
            row, col = divmod(i, 2)
            btn = ModernButton(master=self.frame_botoes_pdf, text=config['nome'],
                               image=self.icons.get(key), command=lambda k=key: self.processar_conversao(k))
            if config.get('enabled') is False: btn.configure(state="disabled", fg_color=COLORS["disabled"])
            btn.grid(row=row, column=col, padx=15, pady=12, sticky="ew")
        self.frame_botoes_pdf.grid_columnconfigure((0, 1), weight=1)

        for key, config in ofx_buttons.items():
            btn = ModernButton(master=self.frame_botoes_ofx, text=config['nome'], image=self.icons.get(key),
                               command=lambda k=key: self.processar_conversao(k),
                               anchor="center", fg_color="#27AE60", hover_color="#2ECC71")
            btn.pack(fill="x", padx=10, pady=10)
            
        self.status_label = ctk.CTkLabel(container, text="Pronto para iniciar.", font=FONTS["status"], text_color=COLORS["text"])
        self.status_label.pack(pady=(20, 0), side="bottom", fill="x")

    def _set_buttons_state(self, new_state: str):
        for frame in [self.frame_botoes_pdf, self.frame_botoes_ofx]:
            for widget in frame.winfo_children():
                if isinstance(widget, ctk.CTkButton):
                    if new_state == "normal" and widget.cget("fg_color") == COLORS["disabled"]: continue
                    widget.configure(state=new_state)

    def update_status(self, message, color=None):
        self.status_label.configure(text=message, text_color=color or COLORS["text"])
        self.update_idletasks()

    def processar_conversao(self, key):
        self._set_buttons_state("disabled")
        self.update_status(f"Iniciando: {CONVERTERS[key]['nome']}", COLORS["warning"])
        
        try:
            success = self.run_converter(key)
            if success:
                self.update_status("Processo concluído com sucesso!", COLORS["success"])
                messagebox.showinfo("Sucesso", f"Conversão de '{CONVERTERS[key]['nome']}' concluída com sucesso!")
        except UserWarning as e:
            self.update_status("Operação cancelada pelo usuário.", COLORS["warning"])
            if str(e): messagebox.showwarning("Operação Cancelada", str(e))
        except Exception:
            error_details = traceback.format_exc()
            messagebox.showerror("Erro Crítico na Execução", f"Ocorreu um erro inesperado:\n\n{error_details}")
            self.update_status("Ocorreu um erro crítico.", COLORS["error"])
        finally:
            self._set_buttons_state("normal")

    def run_converter(self, key):
        config = CONVERTERS[key]
        handler_type = config.get("type")
        
        if handler_type == "ofx":
            return self._run_ofx_converter()

        self.withdraw()
        try:
            if handler_type == "model_choice":
                return self._run_model_choice_converter(key, config)
            elif handler_type == "single_file":
                return self._run_single_file_converter(key, config)
            elif handler_type == "multi_file":
                return self._run_multi_file_converter(key, config)
            elif handler_type == "itau_special":
                return self._run_itau_converter(key, config)
            elif handler_type == "simple_run":
                return self._run_simple_converter(key, config)
        finally:
            if self.winfo_exists():
                self.deiconify()

    def _run_ofx_converter(self):
        from ofxparse import OfxParser
        from openpyxl import Workbook
        caminhos = filedialog.askopenfilenames(title="Selecione os arquivos OFX", filetypes=[("Arquivos OFX", "*.ofx")])
        if not caminhos: raise UserWarning("")
        wb = Workbook(); wb.remove(wb.active)
        for path in caminhos:
            self.update_status(f"Processando {os.path.basename(path)}...")
            with open(path, 'r', encoding='latin-1', errors='ignore') as f:
                ofx = OfxParser.parse(f)
            ws = wb.create_sheet(title=os.path.splitext(os.path.basename(path))[0][:31])
            ws.append(['Data', 'Descrição', 'Valor'])
            for t in ofx.account.statement.transactions: ws.append([t.date.strftime('%d/%m/%Y'), t.memo, t.amount])
        caminho_salvamento = os.path.splitext(caminhos[0])[0] + ".xlsx"
        wb.save(caminho_salvamento)
        return True

    def _run_model_choice_converter(self, key, config):
        modelo = self._escolher_modelo(config['model_config'])
        if not modelo: raise UserWarning("")
        module_name = f"conversor_{key}mod{modelo[-1]}"
        module = importlib.import_module(module_name)
        return module.iniciar_processamento()

    def _run_single_file_converter(self, key, config):
        path = filedialog.askopenfilename(title=f"Selecione o PDF do {config['nome']}", filetypes=[("PDF files", "*.pdf")])
        if not path: raise UserWarning("")
        module = importlib.import_module(config['module'])
        func = getattr(module, config['function'])
        func(path)
        return True

    def _run_multi_file_converter(self, key, config):
        module = importlib.import_module(config['module'])
        paths = module.selecionar_pdfs()
        if not paths: raise UserWarning("")
        for path in paths: module.extrair_texto_pdf(path)
        return True
    
    def _run_itau_converter(self, key, config):
        path = filedialog.askopenfilename(title="Selecione o PDF do Itaú", filetypes=[("PDF files", "*.pdf")])
        if not path: raise UserWarning("")
        itau_configs = {'flavor': 'stream', 'page_1': {'table_areas': ['149,257, 552,21'], 'columns': ['144,262, 204,262, 303,262, 351,262, 406,262, 418,262, 467,262, 506,262, 553,262']}, 'page_2_end': {'table_areas': ['151,760, 553,20'], 'columns': ['157,757, 173,757, 269,757, 309,757, 363,757, 380,757, 470,757, 509,757, 545,757']}}
        module = importlib.import_module("conversor_itau")
        extractor = module.PDFTableExtractor(path, itau_configs)
        extractor.start()
        return True

    def _run_simple_converter(self, key, config):
        module = importlib.import_module(config['module'])
        func = getattr(module, config['function'])
        success = func() 
        if not success: raise UserWarning("")
        return True

    def _escolher_modelo(self, config_modelo):
        modelo_selecionado = ctk.StringVar()
        janela = ctk.CTkToplevel(self)
        janela.title(config_modelo['titulo']); janela.geometry("450x180"); janela.resizable(False, False)
        janela.transient(self); janela.grab_set()
        def selecionar_e_fechar(modelo):
            modelo_selecionado.set(modelo)
            janela.destroy()
        ctk.CTkLabel(janela, text=config_modelo['label'], font=("Segoe UI", 14)).pack(pady=20)
        frame_botoes = ctk.CTkFrame(janela, fg_color="transparent")
        frame_botoes.pack(pady=10)
        for key, text in config_modelo['opcoes'].items():
            ctk.CTkButton(frame_botoes, text=text, command=lambda m=key: selecionar_e_fechar(m), width=180).pack(side="left", padx=10, pady=10)
        self.wait_window(janela)
        return modelo_selecionado.get()

if __name__ == "__main__":
    app = ConversorApp()
    if app.winfo_exists():
        app.mainloop()