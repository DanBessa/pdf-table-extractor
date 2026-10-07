# gerar_chaves_sem_senha.py
from cryptography.hazmat.primitives.asymmetric import rsa
from cryptography.hazmat.primitives import serialization

# 1. Gera nova chave RSA 2048 bits
chave_privada = rsa.generate_private_key(
    public_exponent=65537,
    key_size=2048
)

# 2. Exporta chave privada SEM SENHA (NoEncryption)
priv_pem = chave_privada.private_bytes(
    encoding=serialization.Encoding.PEM,
    format=serialization.PrivateFormat.PKCS8,
    encryption_algorithm=serialization.NoEncryption()
)

# 3. Exporta chave pública
pub_pem = chave_privada.public_key().public_bytes(
    encoding=serialization.Encoding.PEM,
    format=serialization.PublicFormat.SubjectPublicKeyInfo
)

with open("chave_privada.pem", "wb") as f:
    f.write(priv_pem)

with open("chave_publica.pem", "wb") as f:
    f.write(pub_pem)

print("✅ Par de chaves criado sem senha!")
print("\n=== COPIE ESTA CHAVE PÚBLICA PARA O CÓDIGO DO CONVERSOR (CHAVE_PUBLICA_PEM) ===")
print(pub_pem.decode('utf-8'))