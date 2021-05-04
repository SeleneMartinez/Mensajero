import smtplib
import socket
import ssl
import email
from email.message import EmailMessage
import pandas as pd
import settings

encuesta = 'https://docs.google.com/forms/d/e/1FAIpQLSfM2qeqKzj7Xlg4S_NwLagtw71FHgOBt2IQHdtsXxhSj3kkmA/viewform'
data = pd.read_excel('mails.xlsx')
df = pd.DataFrame(data)
contactos = dict(zip(df.MAIL,df.NOMBRE))





def send_mail(usuario,contraseña,de,para,msg):
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(usuario,contraseña)
        msg['From'] = de
        msg['To'] = para
        msg.set_content(body)
    
        try:
            smtp.send_message(msg)
        except:
            with open('no_enviado.txt') as f:
                linea = 'no neviado a: {} \n'.format(para)
        finally:
            del msg['To']
    


for mail, nombre in contactos.items():
    msg = EmailMessage()
    msg['Subject'] = 'Primeras semanas de clase'
    body ="Hola {}, cómo estás?  \n\nEstamos arrancando la cursada, y ya pasaron las charlas de bienvenida, inscripciones, tener acceso al campus, hacer cambios de curso y saber cómo son las clases de cada curso. \nPor eso queremos aprovechar este momento de menos movimiento para saber cómo estás, cómo estás llevando la cursada, si tenés o no dificultades para conectarte, para organizar tu cursada, para conocer compañeros/as o alguna otra inquietud que quieras charlar en este momento. \n\n\n Dejamos, además, una encuesta muy breve y anónima, con el objetivo de identificar algunas de las dificultades en relación al estudio. Por esto te pido que por favor te tomes cinco minutos para responder.\n{}\n\n\nSaludos".format(nombre.capitalize(),encuesta)
    msg.set_content(body)
    send_mail(settings.EMAIL, settings.EMAIL_PASSWORD, settings.EMAIL, mail, msg)