import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

class Email:
    def __init__(self, sender_name: str, complete_name: str, email_name: str, body: str, attachments: list, subject: str):
        """Inicializa o objeto Email com os parâmetros fornecidos."""
        self.sender_name = sender_name
        self.complete_name = complete_name
        self.email_name = email_name
        self.body = body
        self.attachments = attachments
        self.subject = subject


    def __repr__(self):
        """Representação textual do objeto Email."""
        return (f"Email(sender_name={self.sender_name}, complete_name={self.complete_name}, "
                f"email_name={self.email_name}, body={self.body[:30]}..., "
                f"attachments={len(self.attachments)} attachments)")

def create_email(dados: dict):
    """
    Converte um dicionário em objeto Email.
    """

    return Email(
        sender_name="AutoDoc",
        complete_name="AutoDoc PPCOMP",
        email_name="maquinas902@gmail.com",
        subject= f"Documentos de Defesa {dados.get('nome_completo_aluno')}",
        body= f"Documentos criados de forma autônoma pelo AutoDoc",
    )

class EmailSender:
    def __init__(self):
        self.email_address = "maquinas902@gmail.com"
        self.password = "vjmm jseq feuk amhv"
        self.connection = None

    def connect(self):
        '''Conecta ao servidor SMTP para enviar emails.'''
        self.connection = smtplib.SMTP('smtp.gmail.com', 587)
        self.connection.starttls()
        self.connection.login(self.email_address, self.password)
        print("Conectado ao servidor de envio")

    def send_email(self, email_obj: Email, destinatario: str):
        '''Envia um email a partir de um objeto Email e inclui o nome completo do remetente.'''
        msg = MIMEMultipart()
        msg['From'] = self.email_address
        msg['To'] = destinatario
        msg['Subject'] = email_obj.subject

        # Adiciona o corpo do email, incluindo o nome completo do aluno
        body_content = email_obj.body
        msg.attach(MIMEText(body_content, 'plain'))

        # Adiciona os anexos
        if email_obj.attachments:
            for attachment_path in email_obj.attachments:
                part = MIMEBase('application', 'vnd.openxmlformats-officedocument.wordprocessingml.document')
                
                try:
                    with open(attachment_path, 'rb') as attachment:
                        part.set_payload(attachment.read())
                        encoders.encode_base64(part)
                        filename = os.path.basename(attachment_path)
                        part.add_header('Content-Disposition', f'attachment; filename= {filename}')
                        msg.attach(part)
                except Exception as e:
                    print("Erro ao Anexar")

        # Envia o email
        self.connection.sendmail(self.email_address, destinatario, msg.as_string())
        print(f"Email enviado para {destinatario}")

    def disconnect(self):
        '''Desconecta do servidor SMTP.'''
        if self.connection:
            self.connection.quit()

