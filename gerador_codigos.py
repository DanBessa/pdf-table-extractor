import datetime
import json
import base64
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.hazmat.primitives import serialization

import os

# Pega o diretório exato onde este script está salvo
diretriz_atual = os.path.dirname(os.path.abspath(__file__))
caminho_pem = os.path.join(diretriz_atual, "private_key.pem")

def gerar_chave(dias_validade=30):
    with open(caminho_pem, "rb") as key_file: # Usando o caminho absoluto
        private_key = serialization.load_pem_private_key(
            key_file.read(),
            password=b"85957176Ca@",
        )
    # ... (resto do código continua igual)
    # Define a data de expiração
    data_expiracao = datetime.date.today() + datetime.timedelta(days=dias_validade)
    
    # Cria os dados da licença (pode adicionar mais informações se quiser)
    dados_licenca = {
        "expira_em": data_expiracao.isoformat(),
        "info_usuario": "001", # Opcional: pode colocar o nome do cliente aqui
        "cnpj_cliente": "12.345.678/0001-99"
    }
    
    # Converte os dados para bytes
    mensagem = json.dumps(dados_licenca).encode('utf-8')

    # Assina os dados com sua chave privada
    assinatura = private_key.sign(
        mensagem,
        padding.PSS(
            mgf=padding.MGF1(hashes.SHA256()),
            salt_length=padding.PSS.MAX_LENGTH
        ),
        hashes.SHA256()
    )

    # Junta a mensagem e a assinatura
    pacote_final = {
        "dados": base64.b64encode(mensagem).decode('utf-8'),
        "assinatura": base64.b64encode(assinatura).decode('utf-8')
    }

    # Codifica tudo em base64 para criar uma chave de ativação única e "limpa"
    chave_ativacao = base64.b64encode(json.dumps(pacote_final).encode('utf-8')).decode('utf-8')
    
    print("--- CHAVE DE ATIVAÇÃO GERADA ---")
    print(chave_ativacao)
    print(f"\nVálida por {dias_validade} dias. Expira em: {data_expiracao.strftime('%d/%m/%Y')}")

if __name__ == "__main__":
    gerar_chave(dias_validade=30)