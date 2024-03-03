#Librerias
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as condicion
from selenium.webdriver.common.by import By
import xlwings as wx
import openpyxl
import time

path='d:/Usuario/Desktop/chrome_path/chromedriver-win64/chromedriver.exe'
servicioschrome= Service(path)
driver = webdriver.Chrome(service=servicioschrome)
driver.get("https://roc.myrb.io/s1/forms/M6I8P2PDOZFDBYYG")
wait = WebDriverWait(driver,10)
obs_input=wait.until(condicion.presence_of_element_located((By.ID, 'obs')))

#doc_excel= "D:/Usuario/Desktop/python prueba/db/Base Seguimiento Observ Auditoría al_30042021.xlsx"
doc_excel="d:/Usuario/Desktop/Base Seguimiento Observ Auditoría al_30042021.xlsx"
app= wx.App(visible=False)
workbook= app.books.open(doc_excel)
hoja= workbook.sheets['Hoja1']

for i in range (2,46):
    proceso= hoja[f'A{i}'].value
    observacion=hoja[f'B{i}'].value
    riesgo= hoja[f'C{i}'].value
    severidad=hoja[f'D{i}'].value
    plan_accion=hoja[f'E{i}'].value
    fecha_compromiso=hoja[f'F{i}'].value
    responsable=hoja[f'G{i}'].value
    area=hoja[f'H{i}'].value
    correo=hoja[f'I{i}'].value
    estado=hoja[f'J{i}'].value
    obs_input.clear()
    obs_input.send_keys(observacion)
    time.sleep(2)
    print("Datos enviados")


driver.quit()