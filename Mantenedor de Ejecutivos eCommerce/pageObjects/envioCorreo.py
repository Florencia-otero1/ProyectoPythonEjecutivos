import datetime
import glob
import os
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

class EnvioCorreo:
    def __init__(self, smtp_server, smtp_port, smtp_username, smtp_password, filePath, fileName):
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.smtp_username = smtp_username
        self.smtp_password = smtp_password
        self.filePath = filePath
        self.fileName = fileName
    def enviar_mail(subject, body, to_emails, smtp_server, smtp_port, smtp_username, smtp_password, file_path=None, html=False):
            try:
                # Crea el mensaje de correo electrónico
                msg = MIMEMultipart()
                msg['From'] = smtp_username
                msg['To'] = ', '.join(to_emails)
                msg['Subject'] = subject
                hora_actual = datetime.datetime.now()

                # Agrega el cuerpo del correo
                mime_type = "html" if html else "plain"
                msg.attach(MIMEText(body, mime_type))

                # Si se proporciona un archivo, lo adjunta
                if file_path:
                    if os.path.isfile(file_path):
                        with open(file_path, "rb") as attachment:
                            part = MIMEBase('application', 'octet-stream')
                            part.set_payload(attachment.read())
                            encoders.encode_base64(part)
                            part.add_header(
                                'Content-Disposition',
                                f'attachment; filename="{os.path.basename(file_path)}"'
                            )
                            msg.attach(part)
                        print(f"Archivo adjuntado: {file_path}")
                    else:
                        print(f"El archivo no existe: {file_path}")

                # Configura el servidor SMTP dentro de un bloque 'with'
                with smtplib.SMTP(smtp_server, smtp_port) as server:
                    server.starttls()
                    server.login(smtp_username, smtp_password)

                    # Envía el correo
                    server.sendmail(smtp_username, to_emails, msg.as_string())
                    print(f"{hora_actual}: Correo enviado con éxito.")

            except Exception as e:
                print(f"Ocurrió un error al enviar el correo: {str(e)}")