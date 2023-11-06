import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import datetime
import os

class Email:

    type_notification = None

    # Dirección de correo electrónico del remitente
    sender_email = 'soporteqnex@gmail.com'
    # Contraseña de la cuenta del remitente
    password = 'sgekxwsvedocwgtp'

    def __init__(self, type):
        self.type_notification = type
        
    
    def saludar(self):
        # Obtener la fecha actual
        fecha_actual = datetime.date.today()
        print(fecha_actual)

    def send(self, title,data,to_addr,cc_email):

        currentDate = datetime.date.today()
                
        # Abrir el archivo HTML y leer su contenido
        pathTemplate = os.getcwd() + '/importador/util/templates/default.html'
        with open(pathTemplate, 'r') as f:
            html_body = f.read()

        # Crear un mensaje
        message = MIMEMultipart()
        message['From'] = self.sender_email
        message['To'] = to_addr
       
        message['Subject'] = 'BPO ROBOT-REPORTES-SERV.IMPORTACION['+str(currentDate)+']: Resultado de importacion de datos'

        # Agregar el contenido del mensaje
        body = ''
        if(self.type_notification == 'text'):
            body = data
        
        html_body = html_body.format(content=body,title=title)
        message.attach(MIMEText(html_body, 'html'))

        # Conectar al servidor SMTP de Gmail
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(self.sender_email, self.password)

        # Enviar el mensaje
        text = message.as_string()
        toaddrs = [to_addr] + cc_email['cc_list'] + cc_email['bcc_list'] 
        server.sendmail(self.sender_email, toaddrs, text)

        # Cerrar la conexión con el servidor SMTP
        server.quit()

