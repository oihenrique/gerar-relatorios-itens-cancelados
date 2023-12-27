class EmailData:
    """
    A class to store email-related data such as sender email, recipient email, subject, message, and attachments.
    """

    def __init__(self):
        """
        Initialize an instance of the EmailData class.
        """
        self._body_text_path = None
        self._sender_email = ''
        self._sender_password = ''
        self._recipient_email = ''
        self._subject = ''
        self._message = ''
        self._email_list = []
        self._attach_list = []
        self._attachment = ''

    @property
    def body_text_path(self):
        return self._body_text_path

    @body_text_path.setter
    def body_text_path(self, value):
        self._body_text_path = value

    @property
    def sender_email(self):
        return self._sender_email

    @sender_email.setter
    def sender_email(self, value):
        self._sender_email = value

    @property
    def sender_password(self):
        return self._sender_password

    @sender_password.setter
    def sender_password(self, value):
        self._sender_password = value

    @property
    def recipient_email(self):
        return self._recipient_email

    @recipient_email.setter
    def recipient_email(self, value):
        self._recipient_email = value

    @property
    def subject(self):
        return self._subject

    @subject.setter
    def subject(self, value):
        self._subject = value

    @property
    def message(self):
        return self._message

    @message.setter
    def message(self, value):
        self._message = value

    @property
    def email_list(self):
        return self._email_list

    @email_list.setter
    def email_list(self, value):
        self._email_list = value

    @property
    def attach_list(self):
        return self._attach_list

    @attach_list.setter
    def attach_list(self, value):
        self._attach_list = value

    @property
    def attachment(self):
        return self._attachment

    @attachment.setter
    def attachment(self, value):
        self._attachment = value
