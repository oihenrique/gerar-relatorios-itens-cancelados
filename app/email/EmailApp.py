import os.path
from time import sleep

from app.config.Credentials import Credentials
from app.email.EmailServiceManager import EmailServiceManager
from app.utils.SystemService import SystemService


def send_email(date):
    """
    Send emails with attachments to a list of recipients.

    Parameters:
        date (str): The date used for folder and email subject naming.
    """
    credentials = Credentials()

    sender_email = credentials.sender_email
    sender_password = credentials.sender_password

    email_service = EmailServiceManager(sender_email, sender_password)

    attach_folder = SystemService.get_attach_folder(date)

    email_list = email_service.get_email_list('./resources/app/config/lista_emails_teste.JSON')
    file_list = email_service.get_attach_list(attach_folder)
    email_service.set_message('./resources/app/config/email_body.txt')

    if len(file_list) == 0:
        print('Attach path not found')
        return

    base_path = os.path.abspath(attach_folder)

    for store in email_list:
        for file_name in file_list:
            if str(store) in file_name:
                path = os.path.join(base_path, file_name)
                email_service.set_attachment_path(path)
                email_service.set_attachment()
                email_service.set_recipient_email(email_list.get(store))
                subject = f'Solicitação de contagem de estoque - {store}'
                email_service.send_email(subject)
                print(f'Email enviado para loja {store}\n'
                      f'Email: {email_list.get(store)}\n'
                      f'Título: {subject}\n'
                      f'Anexo: {path}\n')

                break

        sleep(18)
