# Enviar correos electrónicos con Python y Pandas
Este código utiliza Python y Pandas para enviar correos electrónicos con archivos adjuntos a partir de un archivo Excel.

## Requisitos
Este código requiere las siguientes bibliotecas de Python:

1. pandas
2. smtplib
3. email
4. os

## Uso
- Asegúrate de tener un archivo Excel con los vencimientos correspondientes para cada cliente (yo los descargo desde errepar pero poddrian copiar el formato y usar sus vencimientos). 
- Asegúrate de que el archivo contenga una columna para el CUIT del cliente, una columna para el correo electrónico del cliente y una columna para el nombre del cliente.
- Actualiza la ruta del archivo Excel en el código Python.
- Vas a tener que actualizar el Email desde el cual se envian los emails y configurar la contraseña desde la terminal utilizando el siguiente comando
setx GMAIL_PASSWORD "your_gmail_password"
- Ejecuta el código Python.
- El código Python enviará un correo electrónico a cada cliente con sus vencimientos correspondientes adjuntos como un archivo Excel.

## Contribuciones
Las contribuciones son bienvenidas. Si te interesa contribuir a este proyecto, por favor crea un fork del repositorio y envía una solicitud de extracción.

Muchas gracias.
