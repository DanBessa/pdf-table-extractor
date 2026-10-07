import sys
import os
import time
import subprocess

def main():
    if len(sys.argv) < 4:
        print("Uso: updater.exe <caminho_atual> <caminho_novo> <arquivo_controle>")
        return
    
    current_exe = sys.argv[1]
    new_exe = sys.argv[2]
    control_file = sys.argv[3]
    
    try:
        # Aguarda o fechamento do programa principal
        print("Aguardando o fechamento do programa principal...")
        time.sleep(2)
        
        # Substitui o executável atual pelo novo
        print(f"Substituindo '{current_exe}' por '{new_exe}'...")
        os.replace(new_exe, current_exe)
        print("Atualização concluída com sucesso.")
        
    except Exception as e:
        print(f"Erro durante a atualização: {e}")
    finally:
        # Remove o arquivo de controle
        if os.path.exists(control_file):
            try:
                os.remove(control_file)
                print(f"Arquivo de controle '{control_file}' removido com sucesso.")
            except Exception as e:
                print(f"Erro ao remover o arquivo de controle: {e}")

if __name__ == "__main__":
    main()