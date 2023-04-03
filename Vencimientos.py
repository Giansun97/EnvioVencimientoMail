import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import os


pd.set_option('display.max_columns', None)
pd.reset_option('display.max_colwidth')
pd.set_option('display.width', 1000)

# Definir la ruta del archivo Excel y leerlo con pandas
excel_file = './Exportación - Agenda - Del 01 al 30 de Abril de 2023.xlsx'
# Leer data frame a partir de la celda 9
df_vencimientos = pd.read_excel(excel_file,
                                skiprows=8)

# Eliminar los registros que con tienen NaN en la columna
df_vencimientos = df_vencimientos[df_vencimientos['CLIENTE'].notna()]

# print(df_vencimientos)

# Obtener una lista de los clientes únicos
clientes = df_vencimientos['CUIT'].unique()

# Iterar sobre cada cliente y enviar los vencimientos correspondientes
for cliente in clientes:
    # Filtrar los datos del DataFrame por el CUIT del cliente
    df_cliente = df_vencimientos[df_vencimientos['CUIT'] == cliente]

    # Configurar los detalles del correo electrónico
    from_address = 'glorenzo@pilarmep.com.ar'
    to_address = df_cliente['Mail'].iloc[0]  # Utilizar el primer correo electrónico de la lista
    subject = f'Calendarios de vencimiento impositivos - {cliente}'
    body = f'Buenos dias, los vencimientos correspondientes a este mes son los siguientes:' \
           f'\n\n{df_cliente.to_html(index=False)}' \
           f'Saludos'

    # Construir el mensaje
    msg = MIMEMultipart()
    msg['From'] = from_address
    msg['To'] = to_address
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'html'))

    # Filtrar el archivo Excel solo para el cliente correspondiente y adjuntarlo al correo electrónico
    df_cliente.to_excel('vencimientos_cliente.xlsx', index=False)
    with open('vencimientos_cliente.xlsx', 'rb') as f:
        attachment = MIMEApplication(f.read(), _subtype='xlsx')
        attachment.add_header('Content-Disposition', 'attachment', filename='vencimientos_cliente.xlsx')
        msg.attach(attachment)

    # Enviar el correo electrónico utilizando SMTP
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(from_address, os.environ['GMAIL_PASSWORD_MEP'])
    server.sendmail(from_address, to_address, msg.as_string())
    server.quit()
