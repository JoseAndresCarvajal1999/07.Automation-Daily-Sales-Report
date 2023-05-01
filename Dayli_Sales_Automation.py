from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from time import sleep
import os 
import datetime
import pandas as pd 
import re 
#%%
file_path = r'ScraperJivo'
os.chdir(file_path)
file_path = r'ScraperJivo\Ventas y Entregas'
#%%
conf_driver = {
                  "download.prompt_for_download": False,
                  "download.directory_upgrade": True,
                  "safebrowsing_for_trusted_sources_enabled": False,
                  "download.default_directory": file_path}
 
chromeOptions = webdriver.ChromeOptions()
chromeOptions.add_argument("--disable-notifications")
chromeOptions.add_argument('--no-sandbox')
chromeOptions.add_argument('--verbose')
chromeOptions.add_experimental_option("prefs", conf_driver)

chromeOptions.add_argument('--disable-gpu')
chromeOptions.add_argument('--disable-software-rasterizer')
chromeOptions.add_argument('--headless')
chromeOptions.add_argument('window-size=1920x1080')

#chromeOptions.add_argument("--start-maximized")
#chromeOptions.add_argument("--start-fullscreen")
chromeOptions.add_argument("--no-proxy-server")
chromeOptions.add_argument("--proxy-server='direct://'")
chromeOptions.add_argument("--proxy-bypass-list=*") 

    
driver = webdriver.Chrome(executable_path='chromedriver', chrome_options = chromeOptions)
page = 'https://app.jivosite.com/login?dh=jivochat.com.co%2F&ewv=1&form_url=index&lang=es&pricelist_id=129&utm_campaign=direct&utm_source=direct'
driver.get(page)
sleep(5)
#Credenciales para entras al Jivo 
driver.find_element_by_css_selector('#app-layout > section > div > div.wrapper__ICGOp > div.email__sVPKs.container__yz0HL.mousetrap.input-group.ym-disable-keys > input').send_keys(os.environ['User'])
driver.find_element_by_css_selector('#app-layout > section > div > div.wrapper__ICGOp > div.password__nB9LV.container__yz0HL.mousetrap.input-group.ym-disable-keys > input').send_keys(os.environ['Password'])
driver.find_element_by_css_selector('#app-layout > section > div > div.wrapper__ICGOp > div.btnContainer__SRTSV.loginWrapper__Zyoxv > button').click()
sleep(10)
#Pasar a la parte del CRM
driver.find_element_by_css_selector('#app-layout > aside > div > div.sideMenu__s3wnX > div.groupTop__RayWm > div > div:nth-child(2) > div.icon__NfwSQ > div.ico__qjExc.crm__lu43y').click()
sleep(5)
#Pasar a la parte de tratos Horizontales 
driver.find_element_by_css_selector('#app-layout > section > div > div:nth-child(2) > div.wrapper__pbup1 > div.toolbar___KwrP > div.icons__iB0Jg > div.icon__hA9AG.list-view__dIhgZ').click()
#Aplicar Filtros
driver.find_element_by_css_selector('#app-layout > section > div > div:nth-child(2) > div.wrapper__RIvdU.sectionHeader > div.navigation__Mp7vy > div > div.filterButton__VQlck > div > div > div.content__gEkAw').click()
#Fecha de inicio del trato 
driver.find_element_by_css_selector('#app-layout > div > div > div.inner__d3_V0 > div > div:nth-child(3) > div > div:nth-child(1) > div > button').click()
#Tratos del dia anterior 
#driver.find_element_by_css_selector('#app-layout > div > div > div.inner__d3_V0 > div > div:nth-child(3) > div > div.dropdown-menu.select-decor.scroll-container > ul > li:nth-child(3) > span').click()
#Tratos del dia 
driver.find_element_by_css_selector('#app-layout > div > div > div.inner__d3_V0 > div > div:nth-child(3) > div > div.dropdown-menu.select-decor.scroll-container > ul > li:nth-child(2) > span').click()
#Tratos del Ultimo mes 
#driver.find_element_by_css_selector('#app-layout > div > div > div.inner__d3_V0 > div > div:nth-child(3) > div > div.dropdown-menu.select-decor.scroll-container > ul > li:nth-child(6) > span').click()
#Exportar Datos 
driver.find_element_by_css_selector('#app-header > div.toolbar__NXz4s').click()
sleep(5)
#Descargar en Excel 
driver.find_element_by_css_selector('#app-header > div.toolbar__NXz4s > div > div > div.dropdown > ul > li:nth-child(1)').click()
sleep(20)
driver.find_element_by_css_selector('#notifyContainer > div > div.okBtn__WRANv').click()
sleep(5)
driver.quit()
#%%
downloaded_file_path = max([file_path + "/" + f for f in os.listdir(file_path)],key=os.path.getctime)
dt_m = datetime.datetime.fromtimestamp(os.path.getmtime(downloaded_file_path))
os.chdir(file_path)
new_name =  str(dt_m.year) +' -'+ str(dt_m.month) +'-'+ str(dt_m.day)
downloaded_file_frame = pd.read_excel(downloaded_file_path)
Nombre = downloaded_file_frame['Nombre del cliente']
Fecha_inicio = pd.to_datetime(downloaded_file_frame['Fecha de inicio del trato']).dt.date
Fecha_entrega = pd.to_datetime(downloaded_file_frame['Fecha cierre del trato ']).dt.date
Telefono = downloaded_file_frame['Número de teléfono'].apply(lambda x: str(x)[2:] if str(x)[0:2]=='57' else str(x))
Valor_Total = downloaded_file_frame['Monto del trato ']

Materias_lista = []
Tipo_Tareas = []
Tutores_lista = []
Medios_pago = []
valores_tutores = []
Primeros_pagos = []

for resumen in downloaded_file_frame['Comentario']:
    Materia = r'Materia:([\s\S]*?)\.'
    Materias_lista.append(re.search(Materia,resumen).group(1))
    Tipo_tarea =r'Tipo de tarea:([\s\S]*?)\.'
    Tipo_Tareas.append(re.search(Tipo_tarea,resumen).group(1))
    Medio_pago = r'Medio de pago:([\s\S]*?)\.'
    Medios_pago.append(re.search(Medio_pago,resumen).group(1).replace(' ',''))
    aux1 = resumen.replace('Valor Tutor','Valor')
    Tutor = r'Tutor:([\s\S]*?)\.'
    Tutores_lista.append(re.search(Tutor,aux1).group(1))
    Primer_Pago = r'pago: (\d+.\d+)'
    Primeros_pagos.append(int(re.findall(Primer_Pago,aux1)[0].replace('.','')))
    Valor_tutor = r'Valor: (\d+.\d+)'
    valores_tutores.append(int(re.findall(Valor_tutor,aux1)[0].replace('.','')))
    #valores_tutores.append(re.search(Valor_Tutor,aux1))

Resumen = pd.DataFrame()
Resumen['Fecha Venta'] = Fecha_inicio
Resumen['Fecha Entrega'] = Fecha_entrega
Resumen['Cliente'] = Nombre
Resumen['Celular'] = Telefono
Resumen['Materia'] = Materias_lista
Resumen['Tipo Tarea'] = Tipo_Tareas
Resumen['Valor Pago'] = Valor_Total
Resumen['Valor Tutor'] = valores_tutores
Resumen['Primer Pago'] = Primeros_pagos
Resumen['Tutor'] = Tutores_lista
Resumen['Medio De Pago'] = Medios_pago
Resumen['Venta de:'] = downloaded_file_frame['Encargado del trato ']

os.remove(downloaded_file_path)
new_name  = '2022-11-25'
Resumen.to_excel(new_name+'.xlsx')

