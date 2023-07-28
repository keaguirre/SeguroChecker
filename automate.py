#ref https://learn.microsoft.com/en-us/microsoft-edge/webdriver-chromium/?tabs=python
#ref 2 https://www.youtube.com/watch?v=FaVEISsLld0
from selenium import webdriver
import pandas as pd
from selenium.webdriver.common.by import By

import time

edge_options = webdriver.EdgeOptions()
edge_options.add_argument("--disable-features=EdgeExperimentsWithFeatures,EdgeExperimentsWithFeaturesCollection")

driver = webdriver.Edge(options=edge_options)

driver.implicitly_wait(5) #Al usar implicitly_wait(), el WebDriver esperará hasta 5 segundos antes de intentar encontrar cualquier elemento en la página.
driver.get('https://covid.aach.cl/') #abre la web

#me pidio este modulo ImportError: Missing optional dependency 'openpyxl'.  Use pip or conda to install openpyxl.
df = pd.read_excel('bbdd.xlsx')
for row, datos in df.iterrows():
    # print ('row: ',row) #print index
    # print (f'datos:\n {datos.RUT}')
    #print (datos['RUT'])
    print (datos.RUT)
    #Ingresar rut al form
    driver.find_element(By.ID, "masterContenido_txtRunAsegurado").send_keys(datos.RUT)
    print(f'Rut: {datos.RUT} ingresado')
    driver.find_element(By.ID, "masterContenido_consultar").click()
