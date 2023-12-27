import json
import os.path
from os import listdir
from os.path import basename
from email.mime.application import MIMEApplication

from app.email.EmailData import EmailData
from app.email.EmailSender import send_email_smtp


class EmailServiceManager:
    """
    A class for managing email-related operations, such as setting email content, attachments, and sending emails.
    """

    def __init__(self, sender_email, sender_password):
        """
        Initialize an instance of the EmailServiceManager class.

        Parameters:
            sender_email (str): The sender's email address.
            sender_password (str): The sender's email password.
        """
        self._email = EmailData()
        self._email.sender_email = sender_email
        self._email.sender_password = sender_password

    def set_message(self, text_file_path):
        """
        Set the email message content from a text file.

        Parameters:
            text_file_path (str): The path to the text file containing the email message.
        """
        absolute_path = os.path.abspath(text_file_path)
        with open(absolute_path, 'r', encoding='utf-8') as file:
            self._email.message = file.read()

    def set_attachment_path(self, file_path):
        """
        Set the path to the email attachment file.

        Parameters:
            file_path (str): The path to the attachment file.
        """
        self._email.attachment = file_path

    def set_attachment(self):
        """
        Set the email attachment using the specified attachment file path.
        """
        new_name = self._email.attachment.split('/')[-1]

        with open(self._email.attachment, "rb") as file:
            part = MIMEApplication(
                file.read(),
                Name=basename(new_name)
            )
        # After the file is closed
        part['Content-Disposition'] = 'attachment; filename="%s"' % basename(new_name)

        self._email.attachment = part

    def set_recipient_email(self, email):
        """
        Set the recipient's email address.

        Parameters:
            email (str): The recipient's email address.
        """
        self._email.recipient_email = email

    def get_attach_list(self, folder_path):
        """
        Get the list of attachment files in the specified folder.

        Parameters:
            folder_path (str): The path to the folder containing attachment files.

        Returns:
            list: The list of attachment files.
        """
        absolute_path = os.path.abspath(folder_path)

        file_list = []
        for i in listdir(absolute_path):
            if '.xlsx' in i:
                file_list.append(i)
        file_list.sort()

        self._email.attach_list = file_list
        return self._email.attach_list

    def get_email_list(self, json_file_path):
        """
        Get the email list from the specified JSON file.

        Parameters:
            json_file_path (str): The path to the JSON file containing the email list.

        Returns:
            dict: The email list.
        """
        absolute_path = os.path.abspath(json_file_path)
        with open(absolute_path, 'r') as dados:
            self._email.email_list = json.load(dados)

        return self._email.email_list

    def send_email(self, subject):
        """
        Send the email with the specified subject.

        Parameters:
            subject (str): The subject of the email.
        """
        self._email.subject = subject
        send_email_smtp(self._email)
