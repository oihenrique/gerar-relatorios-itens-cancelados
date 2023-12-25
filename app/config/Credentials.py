from decouple import config


class Credentials:
    def __init__(self):
        self._sender_email = config('EMAIL_USERNAME')
        self._sender_password = config('EMAIL_PASSWORD')

    @property
    def sender_email(self):
        return self._sender_email

    @property
    def sender_password(self):
        return self._sender_password
