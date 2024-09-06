import pandas as pd
import smtplib
from email.message import EmailMessage
import logging

# Configurar logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def carregar_lista_destinatarios(arquivo_excel):
    """Carrega a lista de destinatários a partir de um arquivo Excel."""
    try:
        dblist = pd.read_excel(arquivo_excel)
        return dblist.to_dict('records')
    except Exception as e:
        logging.error(f"Erro ao carregar lista de destinatários: {e}")
        return []

def enviar_email(person, server, email_address):
    """Envia um e-mail para um destinatário específico."""
    try:
        msg = EmailMessage()
        msg['Subject'] = 'Email Teste'
        msg['From'] = email_address
        msg['To'] = person['Email']

        # Criação do corpo do e-mail em HTML
        html_content = f"""
        <html>
            <head>
                <style>
                    body {{ font-family: Arial, sans-serif; }}
                    .header {{ background-color: #f4f4f4; padding: 10px; text-align: center; }}
                    .content {{ margin: 20px; }}
                    .footer {{ background-color: #f4f4f4; padding: 10px; text-align: center; }}
                    a {{ color: #1a73e8; text-decoration: none; }}
                    a:hover {{ text-decoration: underline; }}
                </style>
            </head>
            <body>
                <div class="header">
                  <h1>Olá {person['Nome']}!</h1>
                </div>
                <div class="content">
                  <p>
                    Estamos felizes em informar que sua empresa, {person['Empresa']}, foi selecionada para participar do nosso evento especial!
                  </p>
        """

        if person['Diretor'] == "Sim":
            html_content += '<p>Segue link para preenchimento do formulário <a href="#">Clique aqui</a></p>'
        else:
            html_content += '<p>Procure o diretor de sua empresa.</p>'

        html_content += """
                    <p>Salve este contato para receber as próximas comunicações!</p>
                </div>
                <div class="footer">
                    <p>
                      Atenciosamente, </br>
                      Equipe Projeto Berlim.
                    </p>
                </div>
            </body>
        </html>
        """

        msg.set_content(html_content, subtype='html')

        # Enviando o e-mail
        server.send_message(msg)
        logging.info(f'E-mail enviado para {person["Email"]} com sucesso!')
    except Exception as e:
        logging.error(f'Erro ao enviar e-mail para {person["Email"]}: {e}')

def conectar_servidor(email_address, email_password):
    """Conecta ao servidor de e-mail SMTP."""
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(email_address, email_password)
        logging.info("Conectado ao servidor de e-mail.")
        return server
    except Exception as e:
        logging.error(f'Erro ao conectar ao servidor de e-mail: {e}')
        return None

def main():
    # Defina as credenciais do seu e-mail
    EMAIL_ADDRESS = 'seuemail@aqui.com'
    EMAIL_PASSWORD = 'sua_senha_aqui'

    # Carregar a lista de destinatários
    destinatarios = carregar_lista_destinatarios('Projeto Berlim - DB Lista.xlsx')

    # Conectar ao servidor de e-mail
    server = conectar_servidor(EMAIL_ADDRESS, EMAIL_PASSWORD)

    if server:
        # Enviar e-mails para todos os destinatários
        for person in destinatarios:
            enviar_email(person, server, EMAIL_ADDRESS)

        # Fechar a conexão com o servidor
        server.quit()

if __name__ == "__main__":
    main()
