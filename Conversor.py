import os
import sys
import customtkinter as ctk
from tkinter import messagebox, filedialog
from PIL import Image, ImageTk

# Configuração do tema do CustomTkinter
ctk.set_appearance_mode("System")  # Modo de aparência (System, Light, Dark)
ctk.set_default_color_theme("blue")  # Tema de cores

executado = False  # Variável de controle para evitar chamadas duplicadas

def escolher_banco():
    global banco_selecionado
    banco_selecionado = None

    def selecionar_banco(banco):
        global banco_selecionado
        banco_selecionado = banco
        janela.quit()

    # Verifica se o programa está sendo executado como um executável ou script
    if getattr(sys, 'frozen', False):
        caminho_base = sys._MEIPASS
    else:
        caminho_base = os.path.dirname(os.path.abspath(__file__))

    # Configuração da janela principal
    janela = ctk.CTk()
    janela.title("Conversor bancário")
    janela.geometry("300x200")
    janela.resizable(False, False)

    # Carrega o ícone da janela
    caminho_icone = os.path.join(caminho_base, "rclogo.ico")
    if os.path.exists(caminho_icone):
        try:
            janela.iconbitmap(caminho_icone)
        except Exception as e:
            print(f"Erro ao carregar ícone: {e}")

    # Função para carregar imagens dos botões
    def carregar_imagem(caminho, tamanho):
        try:
            imagem = Image.open(caminho).resize(tamanho, Image.LANCZOS)
            return ImageTk.PhotoImage(imagem)
        except Exception as e:
            print(f"Erro ao carregar imagem {caminho}: {e}")
            return None

    # Caminhos das imagens dos bancos
    caminho_bb = os.path.join(caminho_base, "bblogo.ico")
    caminho_itau = os.path.join(caminho_base, "itaulogo.png")
    caminho_inter = os.path.join(caminho_base, "inter.png")
    caminho_sicoob = os.path.join(caminho_base, "sicoob2.png")

    # Carrega as imagens
    icone_bb = carregar_imagem(caminho_bb, (20, 18))
    icone_itau = carregar_imagem(caminho_itau, (25, 15))
    icone_inter = carregar_imagem(caminho_inter, (25, 22))
    icone_sicoob = carregar_imagem(caminho_sicoob, (50, 50))

    # Título da janela
    titulo = ctk.CTkLabel(janela, text="Escolha o banco:", font=("Arial", 14))
    titulo.pack(pady=10)

    # Botões para selecionar o banco
    btn_bb = ctk.CTkButton(
        janela,
        text="",
        font=("Arial", 12),
        command=lambda: selecionar_banco("bb"),
        image=icone_bb,
        compound="left"
    )
    btn_bb.pack(pady=5)

    btn_inter = ctk.CTkButton(
        janela,
        text="",
        font=("Arial", 12),
        command=lambda: selecionar_banco("inter"),
        image=icone_inter,
        compound="left"
    )
    btn_inter.pack(pady=5)

    btn_itau = ctk.CTkButton(
        janela,
        text="",
        font=("Arial", 12),
        command=lambda: selecionar_banco("itau"),
        image=icone_itau,
        compound="left"
    )
    btn_itau.pack(pady=5)
    
    btn_sicoob = ctk.CTkButton(
        janela,
        text="",
        font=("Arial", 12),
        command=lambda: selecionar_banco("sicoob"),
        image=icone_sicoob,
        compound="left"
    )
    btn_sicoob.pack(pady=5)

    janela.mainloop()
    return banco_selecionado

# Função principal
if __name__ == "__main__":
    banco = escolher_banco()

    if banco == "bb":
        try:
            import conversor_bb
            conversor_bb.selecionar_pdfs()
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao processar o Banco do Brasil: {e}")
        finally:
            sys.exit()

    elif banco == "inter":
        try:
            import conversor_inter
            file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
            if file_path:
                conversor_inter.extrair_dados_inter(file_path)  # Agora passamos o caminho do arquivo
                messagebox.showinfo("Sucesso", "Conversão concluída com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao processar o Banco Inter: {e}")
        finally:
            sys.exit()

    elif banco == "sicoob":
        try:
            import conversor_sicoob
            conversor_sicoob.selecionar_pdf()
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao processar o Banco Sicoob: {e}")
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
                messagebox.showinfo("Sucesso", "Conversão concluída com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao processar o Banco Itaú: {e}")
        finally:
            sys.exit()

    else:
        messagebox.showwarning("Aviso", "Nenhuma opção foi selecionada. O programa será encerrado.")
        sys.exit()