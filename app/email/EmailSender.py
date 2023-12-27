import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


def send_email_smtp(email):
    """
    Send an email using SMTP with the provided EmailData object.

    Parameters:
        email (EmailData): An EmailData object containing email-related information.
    """
    message = MIMEMultipart()
    message['From'] = email.sender_email
    message['To'] = email.recipient_email
    message['Subject'] = email.subject

    html_content = email.message

    html_part = MIMEText(html_content, 'html')

    message.attach(html_part)
    message.attach(email.attachment)

    smtp_server = 'smtp-mail.outlook.com'
    smtp_port = 587

    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()

        server.login(email.sender_email, email.sender_password)
        server.sendmail(email.sender_email, email.recipient_email, message.as_string())

        server.quit()
    except smtplib.SMTPException as e:
        print('Error sending email', str(e))
