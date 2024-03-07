Prueba de automatización de formulario web mediante scraping.
Herramientas: excel y python.
Librerias complementarias: selenium, xlwing, openpyxl, smtplib, datetime, time.
Automatización de mails: tiene por ejemplo mi correo, contraseña falsa y está adecuado al servicio de mensajeria Gmail.com

Consigna:
Utiliza Xlwings para abrir un archivo Excel proporcionado
Procesar la información con la siguiente regla:
- Si el estado del proceso (Columna J) es Regularizado, subir la información al formulario formulario web con la información correspondiente.
- Si el estado del proceso es Atrasado, enviar un mail al responsable. El mail debe indicar el proceso, el estado, la observación y la fecha de compromiso.
- Los estados Pendientes, deben ser ignorados.
