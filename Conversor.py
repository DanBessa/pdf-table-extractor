# menuprincipal.py (VERSÃO FINAL, OTIMIZADA E COMPLETA)

import os
import sys
import importlib
import customtkinter as ctk
from tkinter import messagebox, filedialog
from PIL import Image
from customtkinter import CTkImage
import traceback

# --- IDENTIDADE VISUAL ---
COLORS = {
    "background": "#1A1B26", "frame": "#2A2D3A", "accent": "#78add0", "hover": "#5DADE2",
    "text": "#F0F0F0", "disabled": "#4A4D5A", "success": "#2ECC71", "warning": "#F39C12", "error": "#E74C3C",
}
FONTS = {"title": ("Segoe UI", 30, "bold"), "button": ("Segoe UI", 12, "bold"), "status": ("Segoe UI", 12)}

# --- DICIONÁRIO DE CONFIGURAÇÃO DOS CONVERSORES ---
CONVERTERS = {
    "bb": {
        "nome": "Banco do Brasil", "icone": "bblogo.png", "aba": "pdf", "type": "model_choice",
        "model_config": {
            "titulo": "Seleção de Modelo BB",
            "label": "Selecione o modelo do extrato do Banco do Brasil:",
            "opcoes": {"modelo1": "Modelo 1 (Com Cabeçalho)", "modelo2": "Modelo 2"}
        }
    },
    "inter": { "nome": "Inter", "icone": "inter.png", "aba": "pdf", "type": "single_file", "module": "conversor_inter", "function": "iniciar_processamento" },
    "itau": { "nome": "Itaú", "icone": "itaulogo.png", "aba": "pdf", "type": "itau_special" },
    "sicoob": {
        "nome": "Sicoob", "icone": "sicoob2.png", "aba": "pdf", "type": "model_choice",
        "model_config": {
            "titulo": "Seleção de Modelo Sicoob",
            "label": "Selecione o modelo do extrato do Sicoob:",
            "opcoes": {"modelo1": "Modelo 1", "modelo2": "Modelo 2 (Quebras)"}
        }
    },
    "bradesco": { "nome": "Bradesco", "icone": "bradesco.png", "aba": "pdf", "type": "simple_run", "module": "conversor_bradesco", "function": "main" },
    "pagbank": { "nome": "PagBank", "icone": "pagbank.png", "aba": "pdf", "type": "multi_file", "module": "conversor_pagbank" },
    "santander": { "nome": "Santander", "icone": "santander.png", "aba": "pdf", "type": "simple_run", "module": "conversor_santander", "function": "iniciar_extracao_santander" },
    "cef": { "nome": "Caixa Econômica", "icone": "cef.png", "aba": "pdf", "type": "simple_run", "module": "conversor_cef", "function": "main" },
    "c6": { "nome": "C6 Bank", "icone": "c6logo.png", "aba": "pdf", "type": "simple_run", "module": "conversor_c6", "function": "iniciar_processamento" },
    "banestes": { "nome": "Banestes", "icone": "banestes.png", "aba": "pdf", "type": "simple_run", "module": "conversor_banestes", "function": "iniciar_processamento" }, #"enabled": False },
    "stone": { "nome": "Stone", "icone": "stone.png", "aba": "pdf", "type": "simple_run", "module": "conversor_stone", "function": "iniciar_processamento" },# "enabled": False },
    "ofx": { "nome": "Converter Arquivo(s) OFX para Excel", "icone": None, "aba": "ofx", "type": "ofx" },
}

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
        self.title("Conversor Bancário")
        self.geometry("550x650")
        self.resizable(False, False)
        
        self.base_path = self._get_base_path()
        self.icons = self._load_icons()
        self._create_widgets()

    def _get_base_path(self):
        try: return sys._MEIPASS
        except AttributeError: return os.path.dirname(os.path.abspath(__file__))

    def _load_icons(self):
        icons = {}
        for key, config in CONVERTERS.items():
            if not config.get("icone"): continue
            try:
                path = os.path.join(self.base_path, config['icone'])
                image = Image.open(path)
                icons[key] = CTkImage(dark_image=image, light_image=image, size=(28, 28))
            except Exception as e: print(f"Erro ao carregar ícone para '{key}': {e}")
        return icons

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
        # Chama a função e propaga seu resultado (True/False/Exceção)
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
    app.mainloop()