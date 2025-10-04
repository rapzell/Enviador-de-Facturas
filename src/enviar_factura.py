import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

def enviar_factura(remitente, password, destinatario, asunto, cuerpo, ruta_pdf):
    """Envía una factura por correo electrónico y devuelve True si tiene éxito, False si falla."""
    try:
        msg = MIMEMultipart()
        msg['From'] = remitente
        msg['To'] = destinatario
        msg['Subject'] = asunto

        msg.attach(MIMEText(cuerpo, 'plain', 'utf-8'))

        with open(ruta_pdf, "rb") as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
        
        encoders.encode_base64(part)
        part.add_header(
            'Content-Disposition',
            f"attachment; filename= {os.path.basename(ruta_pdf)}",
        )
        msg.attach(part)

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(remitente, password)
        texto = msg.as_string()
        server.sendmail(remitente, destinatario, texto)
        server.quit()
        return True  # Éxito
    except Exception as e:
        # El error específico se registrará en el hilo principal
        return False # Fallo
    """
    Envía un correo electrónico con una factura PDF adjunta.

    Args:
        remitente (str): Dirección de correo del remitente.
        password (str): Contraseña de aplicación del remitente.
        destinatario (str): Dirección de correo del destinatario.
        ruta_factura (str): Ruta al archivo PDF que se adjuntará.

    Returns:
        tuple: (bool, str|None) donde bool es True si el envío fue exitoso, False en caso contrario,
               y str es un mensaje de error si ocurrió uno.
    """
    try:
        nombre_archivo = os.path.basename(ruta_factura)
        asunto = nombre_archivo

        cuerpo_mensaje = """Saludos...

--
De acuerdo con el Reglamento UE 2016/679 relativo a la protección de las personas físicas en
lo que respecta al tratamiento de datos personales y a la libre circulación de estos datos
(RGPD) y la Ley Orgánica 3/2018 de protección de Datos Personales y Garantía de Derechos
Digitales (LOPD-GDD), le informamos que los datos de contacto utilizados para la presente
comunicación están incluidos en un registro titularidad de LIMPIEZAS MAYLIN SL, con la
finalidad de llevar a cabo la gestión contable y fiscal de la empresa. La causa que legitima este
tratamiento de datos es el consentimiento. Estos datos serán transmitidos a organismos
publicos, bancos, cajas de ahorro, cajas rurales y entidades privadas que prestan los servicios
de asesoramiento laboral, fiscal y contable. Los datos no serán transmitidos a terceros salvo
autorización expresa u obligación legal. Los datos proporcionados se conservarán mientras se
mantenga la relación profesional o durante los años necesarios para cumplir con las
obligaciones legales.
Puede ejercer los derechos de acceso, rectificación, supresión (derecho al olvido), limitación en
el tratamiento, portabilidad y oposición enviando una solicitud por escrito, acompañada de una
fotocopia de su DNI a la siguiente dirección: URB.PARTE ATLANTICO, BL8 , 11406, JEREZ
DE LA FRONTERA (Cádiz), o a través de la dirección de correo electrónico:
limpiezasmaylinsl@gmail.com"""

        msg = MIMEMultipart()
        msg['From'] = remitente
        msg['To'] = destinatario
        msg['Subject'] = asunto

        msg.attach(MIMEText(cuerpo_mensaje, 'plain'))

        with open(ruta_factura, "rb") as f:
            part = MIMEApplication(f.read(), Name=nombre_archivo)
        part['Content-Disposition'] = f'attachment; filename="{nombre_archivo}"'
        msg.attach(part)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(remitente, password)
            smtp.send_message(msg)
        
        return True, None

    except Exception as e:
        return False, str(e)

    # Adjuntar los archivos PDF
    if not lista_archivos_pdf_paths: # Verificar si la lista está vacía
        return False, "No se proporcionaron archivos PDF para adjuntar."

    for ruta_pdf_individual in lista_archivos_pdf_paths:
        if not ruta_pdf_individual or not os.path.exists(ruta_pdf_individual):
            # Considerar si continuar y adjuntar los demás, o fallar todo el envío.
            # Por ahora, se registrará un error y se intentará continuar con los demás.
            # Si se quiere que falle todo, se debe retornar (False, mensaje_error) aquí.
            print(f"ADVERTENCIA: Archivo PDF no encontrado o no accesible, se omitirá: {ruta_pdf_individual}") # O usar un log_callback si estuviera disponible
            continue # Saltar este archivo y continuar con el siguiente
        
        try:
            with open(ruta_pdf_individual, 'rb') as f:
                data = f.read()
                filename = os.path.basename(ruta_pdf_individual)
                msg.add_attachment(
                    data,
                    maintype='application',
                    subtype='pdf', 
                    filename=filename
                )
        except Exception as e:
            # Similar al caso anterior, decidir si fallar todo o solo este adjunto.
            # Por ahora, se retorna un error general si falla la adjunción de CUALQUIER archivo.
            return False, f"Error al adjuntar el archivo PDF '{ruta_pdf_individual}': {str(e)}"

    # Configuración SMTP (Gmail por defecto)
    smtp_host = 'smtp.gmail.com'
    smtp_port = 587 # Puerto para TLS
    smtp_user = remitente_email
    smtp_password = remitente_password

    # Enviar correo
    try:
        with smtplib.SMTP(smtp_host, smtp_port) as server:
            server.ehlo() # Saludo extendido al servidor
            server.starttls() # Iniciar TLS
            server.ehlo() # Saludo extendido de nuevo después de TLS
            server.login(smtp_user, smtp_password)
            server.send_message(msg)
        return True, None
    except smtplib.SMTPAuthenticationError:
        return False, "Error de autenticación SMTP. Verifique el correo y la contraseña del remitente."
    except smtplib.SMTPConnectError:
        return False, f"Error al conectar al servidor SMTP ({smtp_host}:{smtp_port}). Verifique la conexión o la configuración del host/puerto."
    except smtplib.SMTPServerDisconnected:
        return False, "El servidor SMTP se desconectó inesperadamente."
    except Exception as e:
        return False, f"Error general al enviar correo: {str(e)}"

