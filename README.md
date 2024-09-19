Envio de E-mails Automatizado com Tkinter
Este projeto é uma aplicação Python que permite enviar e-mails automatizados usando o Tkinter para a interface gráfica. O usuário pode selecionar o mês e o tipo de e-mail a ser enviado (antes ou após o dia 10) e o e-mail será enviado para destinatários predefinidos com as informações formatadas. Configurado para solicitação de documentos para a área contábil
## Funcionalidades
- Selecionar o Mês: Escolha o mês para o qual a documentação deve ser solicitada.
- Escolher Tipo de E-mail: Selecione entre "EMAIL 1 - Antes do dia 10" ou "EMAIL 2 - Após o dia 10".
- Envio de E-mails: O e-mail será enviado com formatação HTML, incluindo destinatários principais, cópias e cópias ocultas.
- Interface Gráfica: Utiliza Tkinter para criar uma interface gráfica simples.
## Pré-requisitos
- Python 3.x: Certifique-se de que o Python 3 está instalado no seu sistema.
- Bibliotecas Python: O código utiliza as bibliotecas smtplib, email, e tkinter.
Estas são incluídas na instalação padrão do Python, mas podem precisar ser instaladas em alguns ambientes.

## Configuração

1. Clone o reposítorio
```Clone
git clone https://github.com/seuusuario/nome-do-repositorio.git
cd nome-do-repositorio
```
2. Atualize as Configurações de E-mail:

- Substitua 'seuemail@hotmail.com' pelo seu endereço de e-mail.
- Substitua 'SuaSenha' pela sua senha de e-mail.
- Atualize as listas destinatarios_principais, destinatarios_cc e destinatarios_bcc com os endereços de e-mail apropriados.
3.Servidor SMTP:

O código está configurado para usar o servidor SMTP do Office 365 (smtp.office365.com). Se você estiver usando um servidor diferente, altere a configuração conforme necessário.
## Como usar
1. Execute o script.

```bash
python send_email.py
```

2.Interaja com a Interface Gráfica:

- Selecione o Mês: Escolha o mês desejado para a solicitação de documentação.
- Escolha o Tipo de E-mail: Selecione entre "EMAIL 1 - Antes do dia 10" ou "EMAIL 2 - Após o dia 10".
- Clique em "Enviar": O e-mail será enviado e a janela do Tkinter será fechada após o sucesso.

## Exemplo de Email
- Email 1 - Antes do dia 10
```
<html>
<body>
    <p>Prezado cliente,</p>
    <p>No intuito de mantermos a contabilidade de sua empresa em dia, solicito que envie a documentação ref. ao mês de <b>Agosto</b>, sendo eles:</p>
    <ul>
        <li>Documentos ou Planilha Financeira preenchida com todas as despesas incorridas no mês.</li>
        <li>Extratos bancários nos formatos PDF e OFX.</li>
        <li>Caso tenham contratado algum tipo de empréstimo, seguro (gentileza enviar documento pertinente).</li>
    </ul>
    <p><span style="color: red;">Importante lembrar que precisamos dessa documentação até o dia 10 para que tenhamos tempo hábil para escriturar e oferecer relatórios tempestivos caso seja necessário.</span></p>
    <p><u>Qualquer dúvida, encontramo-nos à disposição.</u></p>
    <p><u>Gentileza confirmar o recebimento deste.</u></p>
    <p>At.te.</p>
</body>
</html>
```
- Email 2 - Após o dia 10
```
<html>
<body>
    <p>Prezado cliente,</p>
    <p>No início do mês solicitamos o envio da documentação contábil referente ao período de <b>Agosto</b> para devida escrituração, porém não tivemos nenhum retorno sobre.</p>
    <p>Segue novamente a relação de documentos:</p>
    <ul>
        <li>Extratos bancários nos formatos PDF e OFX.</li>
        <li>Documentos ou Planilha Financeira preenchida com todas as despesas incorridas no mês – Caso necessite, entre em contato para auxílio no preenchimento.</li>
        <li>Caso tenham contratado algum tipo de empréstimo, seguro – gentileza enviar documento pertinente.</li>
    </ul>
    <p><span style="color: red;">Importante lembrar que precisamos dessa documentação até o dia 10 de cada mês para que tenhamos tempo hábil para escriturar e oferecer relatórios tempestivos caso seja necessário.</span></p>
    <p><u>Qualquer dúvida, encontramo-nos à disposição.</u></p>
    <p><u>Gentileza confirmar o recebimento deste.</u></p>
    <p>At.te.</p>
</body>
</html>
```
## Contribuições
Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou enviar pull requests com melhorias.
