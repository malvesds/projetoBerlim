# Projeto Berlim

**Projeto Berlim** é um sistema automatizado para o envio de e-mails em massa com suporte a campos dinâmicos e personalização do corpo da mensagem, desenvolvido em Python. Ele permite que as mensagens sejam adaptadas com base em informações individuais dos destinatários, facilitando o envio de campanhas de e-mail para diferentes grupos de pessoas.

## Funcionalidades

- Envio de e-mails em massa de forma automatizada.
- Suporte a campos dinâmicos nos e-mails (ex.: nome, empresa, ambiente).
- Personalização do corpo do e-mail com base nos dados de cada destinatário.
- Controle de fluxo de envio para múltiplos destinatários.
- Criação de conteúdos condicionais com base nos dados dos destinatários.
  
## Tecnologias Utilizadas

- **Python**: Linguagem de programação principal.
- **smtplib**: Biblioteca para envio de e-mails via SMTP.
- **pandas**: Para manipulação de dados e integração com arquivos Excel.
- **EmailMessage**: Para construção e formatação de e-mails HTML.
- **logging**: Para monitoramento e controle de erros durante o processo de envio.

## Instalação

1. Clone o repositório:

  ```
  git clone https://github.com/seuusuario/projeto-berlim.git
  ```

2. Instale as dependências necessárias

  ```
  pip install pandas
  ```

3. Configure as variáveis de e-mail no arquivo Python (mainBerlim.py):
  
  ```
  EMAIL_ADDRESS = 'seuemail@gmail.com'
  EMAIL_PASSWORD = 'sua_senha_de_aplicativo'
  ```

## Uso

1. Carregue o arquivo Excel com a lista de destinatários na mesma pasta que o script.

2. Execute o script principal para iniciar o envio de e-mails:

  ```
  python mainBerlim.py
  ```

O sistema enviará e-mails para cada destinatário, personalizando o conteúdo com base nas informações fornecidas no Excel.
