#Librerias
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as condicion
from selenium.webdriver.common.by import By
import xlwings as wx
import openpyxl
import smtplib
from email.mime.text import MIMEText
import datetime as date
import time

#Conexiones
path='d:/Usuario/Desktop/chrome_path/chromedriver-win64/chromedriver.exe'
servicioschrome= Service(path)
driver = webdriver.Chrome(service=servicioschrome)
driver.get("https://roc.myrb.io/s1/forms/M6I8P2PDOZFDBYYG")
wait = WebDriverWait(driver,10)
doc_excel="d:/Usuario/Desktop/Base Seguimiento Observ Auditoría al_30042021.xlsx"
app= wx.App(visible=False)
workbook= app.books.open(doc_excel)
hoja= workbook.sheets['Hoja1']


for i in range (2,46):
    #Cargo los datos
    proceso= hoja[f'A{i}'].value
    observacion=hoja[f'B{i}'].value
    tipo_riesgo= hoja[f'C{i}'].value
    severidad=hoja[f'D{i}'].value
    plan_accion=hoja[f'E{i}'].value
    fecha_compromiso=hoja[f'F{i}'].value
    responsable=hoja[f'G{i}'].value
    area=hoja[f'H{i}'].value
    correo=hoja[f'I{i}'].value
    estado=hoja[f'J{i}'].value
    #Cominezo a migrar a las cajas de texto del fomrulario
    if (estado == 'Regularizado'):
    #Proceso
        proc_menudropdown= wait.until(condicion.presence_of_element_located((By.ID, 'process')))
        time.sleep(0.01)
        if(proceso =='Operaciones'):
            Select(proc_menudropdown).select_by_value('operaciones')
            time.sleep(0.2)
        if(proceso =='Cuentas por Cobrar'):
            Select(proc_menudropdown).select_by_value('cuentas')
            time.sleep(0.2)
        if(proceso =='Riesgo'):
            Select(proc_menudropdown).select_by_value('riesgo')
            time.sleep(0.2)
        if(proceso =='TI'):
            Select(proc_menudropdown).select_by_value('ti')
            time.sleep(0.2)
        if(proceso =='Financiero'):
            Select(proc_menudropdown).select_by_value('financiero')
            time.sleep(0.2)
        if(proceso =='Continuidad Operacional'):
            Select(proc_menudropdown).select_by_value('continuidad')
            time.sleep(0.2)
        if(proceso =='Operaciones'):
            Select(proc_menudropdown).select_by_value('operaciones')
            time.sleep(0.2)    
        if(proceso =='Contabilidad'):
            Select(proc_menudropdown).select_by_value('contabilidad')
            time.sleep(0.2)
        if(proceso =='Gobierno Corp'):
            Select(proc_menudropdown).select_by_value('gobierno')
            time.sleep(0.2)
        time.sleep(0.5)
    #Riesgo
        tipo_riesgo_input=wait.until(condicion.presence_of_element_located((By.ID, 'tipo_riesgo'))) #Doble parentesis porque es una tupla, sino lo toma como 2 args distintos.
        tipo_riesgo_input.clear()
        tipo_riesgo_input.send_keys(tipo_riesgo)
        time.sleep(0.2)
    #Severidad
        severidad_menudropdown= wait.until(condicion.presence_of_element_located((By.ID, 'severidad')))
        if(severidad not in ['Medio', 'Alto']): #No se por qué no me toma la == así que uso una no pertenencia.
            Select(severidad_menudropdown).select_by_visible_text('Bajo')
            time.sleep(0.02)
        if(severidad=='Medio'):
            Select(severidad_menudropdown).select_by_visible_text('Medio')
            time.sleep(0.02)
        if(severidad=='Alto'):
            Select(severidad_menudropdown).select_by_visible_text('Alto')
            time.sleep(0.02)
    #Responsable
        resp_input= wait.until(condicion.presence_of_element_located((By.ID, 'res')))
        resp_input.clear()
        resp_input.send_keys(responsable)
        time.sleep(0.2)
    #Fecha
        fecha_compromiso_input=wait.until(condicion.presence_of_element_located((By.ID, 'date')))
        fecha_compromiso_str=fecha_compromiso.strftime("%d-%m-%Y")
        fecha_compromiso_input.clear()
        fecha_compromiso_input.send_keys(fecha_compromiso_str)
        time.sleep(0.2)
    #Observaciones
        obs_input=wait.until(condicion.presence_of_element_located((By.ID, 'obs'))) 
        obs_input.clear()
        obs_input.send_keys(observacion)
        time.sleep(0.2)
    #Enviar
        submit_input=wait.until(condicion.presence_of_element_located ((By.ID, 'submit')))
        submit_input.click()
        print("Datos enviados")
    elif (estado =='Atrasado'):
        smtp_server= "smtp.gmail.com"
        smtp_port= 587
        smtp_user= "gustavoge.ruiz@gmail.com"
        smtp_pass= "nnivstipnshiixdy"
        asunto= "Información auditoria"
        body= (f"{responsable}. Buenos días.\n\n El presente es para consultar las gestiones respecto del Proceso {proceso}. \n Para mayor información se le indican los siguientes datos" 
        f"\n\n Proceso: {proceso}."
        f"\n Estado: {estado}." 
        f"\n Observación: {observacion}."
        f"\n Fecha de compromiso: {fecha_compromiso}."
        f"\n\n Aguardamos sus novedades. Saludos.")
        
        destino_prueba= 'gustavoge.ruiz@gmail.com'
        mensaje= MIMEText(body)
        mensaje['Subject'] = asunto
        mensaje['From']=smtp_user
        mensaje['To'] = destino_prueba

        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_user, smtp_pass)
            server.sendmail(smtp_user, destino_prueba, mensaje.as_string())
        
        print (f"Correo enviado a {destino_prueba}")

    
    print ("Proceso terminado")
    driver.quit()
