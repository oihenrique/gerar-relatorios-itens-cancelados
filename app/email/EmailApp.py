import os.path
from time import sleep

from app.config.Credentials import Credentials
from app.email.EmailServiceManager import EmailServiceManager


def send_email():
    credentials = Credentials()

    sender_email = credentials.sender_email
    sender_password = credentials.sender_password

    email_service = EmailServiceManager(sender_email, sender_password)

    email_list = email_service.get_email_list('app/utils/lista_emails_teste.JSON')
    file_list = email_service.get_attach_list('excel')
    email_service.set_message('app/email/email_body.txt')

    base_path = os.path.abspath('excel')

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


if __name__ == "__main__":
    send_email()
