import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.message import EmailMessage
from numpy import NaN
import pandas as pd
import settings

encuesta = 'https://docs.google.com/forms/d/e/1FAIpQLSfM2qeqKzj7Xlg4S_NwLagtw71FHgOBt2IQHdtsXxhSj3kkmA/viewform'
data = pd.read_excel('mails.xlsx')
df = pd.DataFrame(data)
df.dropna()
contactos = dict(zip(df.MAIL,df.NOMBRE))

email_content = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>BIENVENIDA - Programa de Tutorias FIUBA</title>
</head>
<body>
    <p>Hola {} , cómo estás?</p>
    <p>Soy estudiante de Ingeniería en Informática y te voy a estar acompañando durante el primer año de cursada. Estoy para cualquier duda que tengas sobre como funciona la FIUBA, lo que necesites para poder adaptarte al ritmo de estudio de la facu, o dudas sobre la carrera. 
        La idea de este mail es que tengas algunos links y actividades para que no te pierdas nada importante en el arranque de cursada.</p>
    <p>Cualquier cosa que necesites, o dudas sobre esta información que te estoy pasando, me podés escribir por mail o por el chat de consultas!</p>
    <p>¡¡Éxitos en tu primer cuatrimestre en la FIUBA!!</p>

    <p>|| <b>ACTIVIDADES QUE NO TE PODÉS PERDER</b> ||</p>
    <p><b>CHARLA DE BIENVENIDA</b> Jueves 26/08 - 10:00 </p>
    <p><b>CAFÉ DE L@S TUTORES</b> Viernes 27/08 - 18:00 hs <br>Para sacarte las dudas que te queden de la charla de bienvenida. Además haremos una trivia donde sortearemos unas materas!!</p>
    <p><b>MÁS TALLERES</b> A realizar desde el Lunes 23/08 <br>+ Info en: <a href="http://www.fi.uba.ar/es/node/4443">http://www.fi.uba.ar/es/node/4443</a>
    </p>
    <br>
    <p style ="font-family: sans-serif; font-size: small; font-weight: bold;"> 
        <b>Selene Anahi Martinez </b> <br>
        <b>Secretaria de Inclusión, Género, Bienestar y Articulación Social</b> <br>
        <span style="color:rgb(102, 102, 102); font-size: smaller;">
            Facultad de Ingenería - UBA <br>
            +541131472609 | <a href="mailto:semartinez@fi.uba.ar">semartinez@fi.uba.ar</a><br>
            <a href="http://fi.uba.ar">www.ingenieria.uba.ar</a>
            </span>
    </p>
    
</body>
</html>
"""
text = """Hola {} , cómo estás?
Soy estudiante de Ingeniería en Informática y te voy a estar acompañando durante el primer año de cursada. Estoy para cualquier duda que tengas sobre como funciona la FIUBA, lo que necesites para poder adaptarte al ritmo de estudio de la facu, o dudas sobre la carrera. 
La idea de este mail es que tengas algunos links y actividades para que no te pierdas nada importante en el arranque de cursada.

Cualquier cosa que necesites, o dudas sobre esta información que te estoy pasando, me podés escribir por mail o por el chat de consultas!

¡¡Éxitos en tu primer cuatrimestre en la FIUBA!!

|| ACTIVIDADES QUE NO TE PODÉS PERDER ||
CHARLA DE BIENVENIDA Jueves 26/08 - 10:00 hsCAFÉ DE L@S TUTORES Viernes 27/08 - 18:00 hs
Para sacarte las dudas que te queden de la charla de bienvenida. Además haremos una trivia donde sortearemos unas materas!!MÁS TALLERES A realizar desde el Lunes 23/08+ Info en: http://www.fi.uba.ar/es/node/4443
Selene Anahi Martinez
Secretaria de Inclusión, Género, Bienestar y Articulación Social
Facultad de Ingenería - UBA
+541131472609 | semartinez@fi.uba.ar
www.ingenieria.uba.ar"""




def send_mail(usuario,contraseña,de,para,msg):
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(usuario,contraseña)
    
        try:
            smtp.sendmail(de,para,msg.as_string())
        except:
            with open('no_enviado.txt',"a") as f:
                linea = 'no neviado a: {} \n'.format(para)
                print("no se mando a {}".format(para))
        finally:
            del msg['To']
    


for mail, nombre in contactos.items():
    msg = MIMEMultipart('alternative')
    msg['Subject'] = "BIENVENIDA - Programa de Tutoría"
    msg['From'] = mail
    msg['To'] = nombre
    if( isinstance(nombre,str)):
        nombreCapitalize = nombre.capitalize()
        part1 = MIMEText(text.format(nombreCapitalize), 'plain')
        part2 = MIMEText(email_content.format(nombreCapitalize), 'html')
        msg.attach(part1)
        msg.attach(part2)
        send_mail(settings.EMAIL, settings.EMAIL_PASSWORD, settings.EMAIL, mail, msg)