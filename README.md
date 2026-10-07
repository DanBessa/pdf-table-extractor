\# PDF Table Extractor \& Conversor Bancário



Sistema automatizado desenvolvido em Python para extrair tabelas de extratos bancários consolidados em formato PDF e convertê-los em arquivos estruturados (CSV/Excel). O projeto foi desenhado especificamente para facilitar o dia a dia de analistas contábeis e contadores, gerando arquivos perfeitamente compatíveis com os sistemas de contabilidade \*\*Alterdata\*\* e \*\*Domínio\*\*.



\## 🚀 Funcionalidades



\- \*\*Extração Inteligente:\*\* Processa PDFs de múltiplas páginas com layouts e delimitações de tabelas customizáveis.

\- \*\*Suporte Multi-Banco:\*\* Módulos de conversão dedicados para os principais bancos (Itaú, Banco do Brasil, Bradesco, Santander, Caixa, Banestes, C6, Inter, PagBank, Sicoob, Sicredi, Stone, Safra, Mercado Pago, PayCash).

\- \*\*Sanitização de Dados:\*\* Tratamento automático de nomes de colunas, remoção de vazios e alinhamento de delimitadores consecutivas.

\- \*\*Interface Gráfica Integrada:\*\* Menu estilizado em Tkinter com seletor de arquivos nativo para facilitar a operação do usuário.

\- \*\*Módulo de Licenciamento:\*\* Controle de ativação local via arquivos de licença criptografados (`licenca.dat`).

\- \*\*Atualizador Automático:\*\* Sistema integrado (`updater.py`) para checagem e download automático de novas versões do software.



\## 📦 Estrutura do Projeto



```text

pdf-table-extractor/

├── conversores/             # Scripts individuais de conversão por banco

├── icons/                  # Recursos visuais e logotipos das instituições

├── acionar\_licenca.py      # Rotina de ativação do software

├── menu\_de\_testes.py       # Interface principal de gerenciamento

├── updater.py              # Script responsável pelas atualizações

└── rubricas\_folha.json     # Configurações de mapeamento contábil

```



\## 🔧 Requisitos e Instalação



Para rodar o projeto localmente, certifique-se de ter o \*\*Python 3.10 ou superior\*\* instalado.



1\. Instale as dependências necessárias:

&#x20;  ```bash

&#x20;  pip install camelot-py\[pdf] pandas unidecode tkinter flet

&#x20;  ```

2\. Caso utilize o empacotamento com UPX para os binários:

&#x20;  Certifique-se de manter o executável do `upx.exe` mapeado na pasta raiz.



\## 💻 Como Usar



1\. Inicie a interface principal do sistema:

&#x20;  ```bash

&#x20;  python menu\_de\_testes.py

&#x20;  ```

2\. Na janela que for exibida, selecione o arquivo PDF do extrato bancário.

3\. Informe o intervalo de páginas que deseja processar (Ex: `1,2,4-6`).

4\. O arquivo formatado será gerado automaticamente no mesmo diretório do arquivo original.



\## 📄 Licença



Este projeto é de uso restrito e comercial. A execução depende da ativação de uma licença válida gerenciada pelo arquivo `licenca.dat`.



