import os
import sys
import customtkinter as ctk
from tkinter import messagebox, filedialog
from PIL import Image
from customtkinter import CTkImage

# --- Funções e Lógica Principal ---

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("green")

def escolher_modelo_bb(app_root):
    """
    Cria uma nova janela para o usuário escolher entre os modelos de extrato do BB.
    """
    modelo_selecionado = None
    
    janela_modelo = ctk.CTkToplevel(app_root)
    janela_modelo.title("Seleção de Modelo BB")
    janela_modelo.geometry("400x180")
    janela_modelo.resizable(False, False)
    janela_modelo.transient(app_root)
    janela_modelo.grab_set()

    def selecionar_e_fechar(modelo):
        nonlocal modelo_selecionado
        modelo_selecionado = modelo
        janela_modelo.destroy()

    label = ctk.CTkLabel(janela_modelo, text="Selecione o modelo do extrato do Banco do Brasil:", font=("Segoe UI", 14))
    label.pack(pady=20)

    frame_botoes_modelo = ctk.CTkFrame(janela_modelo)
    frame_botoes_modelo.pack(pady=10)
    
    btn_mod1 = ctk.CTkButton(frame_botoes_modelo, text="Modelo 1 (Com Cabeçalho)", command=lambda: selecionar_e_fechar("modelo1"), width=150)
    btn_mod1.pack(side="left", padx=10, pady=10)
    
    btn_mod2 = ctk.CTkButton(frame_botoes_modelo, text="Modelo 2", command=lambda: selecionar_e_fechar("modelo2"), width=150)
    btn_mod2.pack(side="left", padx=10, pady=10)
    
    app_root.wait_window(janela_modelo)
    
    return modelo_selecionado

def escolher_modelo_sicoob(app_root):
    """
    Cria uma nova janela para o usuário escolher entre os modelos de extrato do Sicoob.
    """
    modelo_selecionado = None
    
    janela_modelo = ctk.CTkToplevel(app_root)
    janela_modelo.title("Seleção de Modelo Sicoob")
    janela_modelo.geometry("400x180")
    janela_modelo.resizable(False, False)
    janela_modelo.transient(app_root)
    janela_modelo.grab_set()

    def selecionar_e_fechar(modelo):
        nonlocal modelo_selecionado
        modelo_selecionado = modelo
        janela_modelo.destroy()

    label = ctk.CTkLabel(janela_modelo, text="Selecione o modelo do extrato do Banco Sicoob:", font=("Segoe UI", 14))
    label.pack(pady=20)

    frame_botoes_modelo = ctk.CTkFrame(janela_modelo)
    frame_botoes_modelo.pack(pady=10)
    
    btn_mod1 = ctk.CTkButton(frame_botoes_modelo, text="Modelo 1", command=lambda: selecionar_e_fechar("modelo1"), width=150)
    btn_mod1.pack(side="left", padx=10, pady=10)
    
    btn_mod2 = ctk.CTkButton(frame_botoes_modelo, text="Modelo 2 (Quebras)", command=lambda: selecionar_e_fechar("modelo2"), width=150)
    btn_mod2.pack(side="left", padx=10, pady=10)
    
    app_root.wait_window(janela_modelo)
    
    return modelo_selecionado

def carregar_imagem(caminho, tamanho):
    try:
        imagem = Image.open(caminho)
        return CTkImage(dark_image=imagem, light_image=imagem, size=tamanho)
    except Exception as e:
        print(f"Erro ao carregar imagem {caminho}: {e}")
        return None

def processar_banco_selecionado(banco, app_root):
    """
    Função central que gerencia qual script chamar e exibe as mensagens finais.
    """
    app_root.withdraw() # Esconde a janela principal enquanto o script do banco roda

    try:
        if banco == "bb":
            modelo = escolher_modelo_bb(app_root)
            if modelo == "modelo1":
                import conversor_bbmod1
                conversor_bbmod1.iniciar_processamento()
                messagebox.showinfo("Sucesso", "Extração do Banco do Brasil (Modelo 1) concluída!")
            elif modelo == "modelo2":
                import conversor_bbmod2
                conversor_bbmod2.iniciar_processamento()
                messagebox.showinfo("Sucesso", "Extração do Banco do Brasil (Modelo 2) concluída!")
            else:
                messagebox.showwarning("Aviso", "Nenhuma opção de modelo selecionada.")
        
        elif banco == "inter":
            import conversor_inter
            file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
            if file_path:
                conversor_inter.extrair_dados_inter(file_path)
                messagebox.showinfo("Sucesso", "Extração do Banco Inter concluída com sucesso!")
        
        elif banco == "sicoob":
            modelo = escolher_modelo_sicoob(app_root)
            import conversor_sicoobmod1
            if modelo == "modelo1":
                import conversor_sicoobmod1
                conversor_sicoobmod1.iniciar_processamento()
                messagebox.showinfo("Sucesso", "Extração do Banco Sicoob (Modelo 1) concluída!")
            elif modelo == "modelo2":
                import conversor_sicoobmod2
                conversor_sicoobmod2.iniciar_processamento()
                messagebox.showinfo("Sucesso", "Extração do Banco Sicoob (Modelo 2) concluída!")
            else:
                messagebox.showwarning("Aviso", "Nenhuma opção de modelo selecionada.")
            
        elif banco == "bradesco":
            import conversor_bradesco
            conversor_bradesco.main()
            messagebox.showinfo("Sucesso", "Extração do Banco Bradesco concluída com sucesso!")

        elif banco == "pagbank":
            import conversor_pagbank
            
            # 1. Chama a função que agora retorna os caminhos dos arquivos
            pdf_paths = conversor_pagbank.selecionar_pdfs()
            
            # 2. Verifica se algum arquivo foi selecionado
            if pdf_paths:
                # 3. Itera sobre cada arquivo e chama a função de extração
                for path in pdf_paths:
                    conversor_pagbank.extrair_texto_pdf(path)
                
                # 4. Exibe uma única mensagem de sucesso após processar todos os arquivos
                messagebox.showinfo("Sucesso", f"Extração de {len(pdf_paths)} arquivo(s) do PagBank concluída com sucesso!")
        elif banco == "santander":
            import conversor_santander
            conversor_santander.iniciar_extracao_santander()
            messagebox.showinfo("Sucesso", "Extração do Santander concluída com sucesso!")

        elif banco == "cef":
            import conversor_cef
            conversor_cef.iniciar_processamento_cef()
            messagebox.showinfo("Sucesso", "Extração da Caixa Econômica Federal concluída com sucesso!")

        elif banco == "itau":
            import conversor_itau
            file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
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
                extractor = conversor_itau.PDFTableExtractor(file_path, configs)
                extractor.start()
                messagebox.showinfo("Sucesso", "Extração do Banco Itaú concluída com sucesso!")
        
        else:
             messagebox.showwarning("Aviso", f"Nenhuma ação definida para '{banco}'.")

    except UserWarning as e:
        messagebox.showwarning("Cancelado", str(e))
    except ImportError as e:
        messagebox.showerror("Erro de Importação", f"Não foi possível encontrar o script necessário: {e}.\n\nVerifique se o arquivo .py do conversor está na mesma pasta.")
    except Exception as e:
        messagebox.showerror("Erro na Execução", f"Ocorreu um erro durante o processamento:\n\n{e}")
    finally:
        app_root.destroy()
        sys.exit()

# --- Bloco de Execução Principal ---
if __name__ == "__main__":
    app = ctk.CTk()
    app.title("Conversor Bancário")
    app.geometry("600x450")
    app.resizable(False, False)

    if getattr(sys, 'frozen', False):
        caminho_base = sys._MEIPASS
    else:
        caminho_base = os.path.dirname(os.path.abspath(__file__))

    # Carregamento de todos os ícones
    caminho_bb = os.path.join(caminho_base, "bblogo.ico")
    caminho_itau = os.path.join(caminho_base, "itaulogo.png")
    caminho_inter = os.path.join(caminho_base, "inter.png")
    caminho_sicoob = os.path.join(caminho_base, "sicoob2.png")
    caminho_bradesco = os.path.join(caminho_base, "bradesco.png")
    caminho_pagbank = os.path.join(caminho_base, "pagbank.png")
    caminho_santander = os.path.join(caminho_base, "santander.png")
    caminho_cef = os.path.join(caminho_base, "cef.png")

    icone_bb = carregar_imagem(caminho_bb, (30, 30))
    icone_itau = carregar_imagem(caminho_itau, (30, 30))
    icone_inter = carregar_imagem(caminho_inter, (30, 30))
    icone_sicoob = carregar_imagem(caminho_sicoob, (30, 30))
    icone_bradesco = carregar_imagem(caminho_bradesco, (30, 30))
    icone_pagbank = carregar_imagem(caminho_pagbank, (30, 30))
    icone_santander = carregar_imagem(caminho_santander, (30, 30))
    icone_cef = carregar_imagem(caminho_cef, (30, 30))

    # --- Construção da Interface Gráfica ---
    titulo = ctk.CTkLabel(app, text="Conversor Bancário", font=("Segoe UI", 20, "bold"))
    titulo.pack(pady=10)

    tabs = ctk.CTkTabview(app, width=550, height=300)
    tabs.pack(pady=10)
    tab_bancos = tabs.add("Bancos")
    frame_botoes = ctk.CTkFrame(tab_bancos)
    frame_botoes.pack(pady=10)

    # Lista completa de botões
    botoes = [
        ("Banco do Brasil", icone_bb, "bb"),
        ("Banco Inter", icone_inter, "inter"),
        ("Itaú", icone_itau, "itau"),
        ("Sicoob", icone_sicoob, "sicoob"),
        ("Bradesco", icone_bradesco, "bradesco"),
        ("PagBank", icone_pagbank, "pagbank"),
        ("Santander", icone_santander, "santander"),
        ("Caixa Econômica", icone_cef, "cef"),
    ]

    for idx, (nome, icone, chave) in enumerate(botoes):
        linha = idx // 2
        coluna = idx % 2
        btn = ctk.CTkButton(
            master=frame_botoes,
            image=icone,
            text=nome,
            font=("Segoe UI", 13),
            width=240,
            anchor="w",
            compound="left",
            hover_color="#2D91F2",
            cursor="hand2",
            command=lambda b=chave: processar_banco_selecionado(b, app)
        )
        btn.grid(row=linha, column=coluna, padx=20, pady=10, sticky="w")

    app.mainloop()