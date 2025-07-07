import os
import sys
import customtkinter as ctk
from tkinter import messagebox, filedialog
from PIL import Image
from customtkinter import CTkImage

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("green")

executado = False

def escolher_banco():
    global banco_selecionado
    banco_selecionado = None

    def selecionar_banco(banco):
        global banco_selecionado
        banco_selecionado = banco
        janela.quit()
        janela.destroy()

    if getattr(sys, 'frozen', False):
        caminho_base = sys._MEIPASS
    else:
        caminho_base = os.path.dirname(os.path.abspath(__file__))

    janela = ctk.CTk()
    janela.title("Conversor bancário")
    janela.geometry("600x450")
    janela.resizable(False, False)

    def carregar_imagem(caminho, tamanho):
        try:
            imagem = Image.open(caminho)
            return CTkImage(dark_image=imagem, light_image=imagem, size=tamanho)
        except Exception as e:
            print(f"Erro ao carregar imagem {caminho}: {e}")
            return None

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

    titulo = ctk.CTkLabel(janela, text="Conversor Bancário", font=("Segoe UI", 20, "bold"))
    titulo.pack(pady=10)

    tabs = ctk.CTkTabview(janela, width=550, height=300)
    tabs.pack(pady=10)

    tab_bancos = tabs.add("Bancos")

    frame_botoes = ctk.CTkFrame(tab_bancos)
    frame_botoes.pack(pady=10)

    botoes = [
        ("Banco do Brasil", icone_bb, "bb"),
        ("Banco Inter", icone_inter, "inter"),
        ("Itaú", icone_itau, "itau"),
        ("Sicoob", icone_sicoob, "sicoob"),
        ("Bradesco", icone_bradesco, "bradesco"),
        ("PagBank", icone_pagbank, "pagbank"),
        ("Santander", icone_santander, "santander"),
        ("Caixa Economica Federal", icone_cef, "cef"),
    ]

    colunas = 2
    largura_botao = 240
    estilo_botao = {
        "font": ("Segoe UI", 13),
        "width": largura_botao,
        "anchor": "w",
        "cursor": "hand2",
        "hover_color": "#2D91F2",
        "master": frame_botoes,
    }

    for idx, (nome, icone, chave) in enumerate(botoes):
        linha = idx // colunas
        coluna = idx % colunas
        btn = ctk.CTkButton(
            image=icone,
            text=nome,
            compound="left",
            command=lambda b=chave: selecionar_banco(b),
            **estilo_botao
        )
        btn.grid(row=linha, column=coluna, padx=20, pady=10, sticky="w")

    janela.mainloop()
    return banco_selecionado


# Execução principal
if __name__ == "__main__":
    banco = escolher_banco()

    if banco == "bb":
        try:
            import conversor_bb
            conversor_bb.selecionar_pdfs()
            messagebox.showinfo("Sucesso", "Extração do Banco do Brasil concluída com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao processar o Banco do Brasil: {e}")
        finally:
            sys.exit()

    elif banco == "inter":
        try:
            import conversor_inter
            file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
            if file_path:
                conversor_inter.extrair_dados_inter(file_path)
                messagebox.showinfo("Sucesso", "Extração do Banco Inter concluída com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao processar o Banco Inter: {e}")
        finally:
            sys.exit()

    elif banco == "sicoob":
        try:
            import conversor_sicoob
            conversor_sicoob.selecionar_pdf()
            messagebox.showinfo("Sucesso", "Extração do Banco Sicoob concluída com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao processar o Banco Sicoob: {e}")
        finally:
            sys.exit()

    elif banco == "bradesco":
        try:
            import conversor_bradesco
            conversor_bradesco.main()
            messagebox.showinfo("Sucesso", "Extração do Banco Bradesco concluída com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao processar o Banco Bradesco: {e}")
        finally:
            sys.exit()


    elif banco == "pagbank":
        try:
            import conversor_pagbank
            pdf_path = conversor_pagbank.selecionar_pdfs()
            if pdf_path:
                conversor_pagbank.extrair_texto_pdf(pdf_path)
                messagebox.showinfo("Sucesso", "Extração do PagBank concluída com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao processar o PagBank: {e}")
        finally:
            sys.exit()

    elif banco == "santander":
        try:
            import conversor_santander
            conversor_santander.iniciar_extracao_santander()
            messagebox.showinfo("Sucesso", "Extração do Santander concluída com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao processar o Santander: {e}")
        finally:
            sys.exit()

    elif banco == "cef":
        try:
            import conversor_cef
            conversor_cef.iniciar_processamento_cef()
            messagebox.showinfo("Sucesso", "Extração da Caixa Economica Federal concluída com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao processar o Santander: {e}")
        finally:
            sys.exit()

    elif banco == "itau":
        try:
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
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao processar o Banco Itaú: {e}")
        finally:
            sys.exit()

    else:
        messagebox.showwarning("Aviso", "Nenhuma opção foi selecionada. O programa será encerrado.")
        sys.exit()
