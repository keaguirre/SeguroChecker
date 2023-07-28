#imports
from selenium import webdriver
import pandas as pd
from selenium.webdriver.common.by import By
import time
from datetime import datetime
from selenium.common.exceptions import NoSuchElementException

edge_options = webdriver.EdgeOptions()
edge_options.add_argument("--disable-features=EdgeExperimentsWithFeatures,EdgeExperimentsWithFeaturesCollection")

driver = webdriver.Edge(options=edge_options)

driver.implicitly_wait(5) #Al usar implicitly_wait(), el WebDriver esperará hasta 5 segundos antes de intentar encontrar cualquier elemento en la página.
driver.get('https://covid.aach.cl/') #abre la web.
df = pd.read_excel('bbdd.xlsx') #abre el excel con el modulo openpyxl.

#Ruts de prueba - Envia la data a la caja de texto con .send_keys('value').-------------------

#driver.find_element(By.ID, "masterContenido_txtRunAsegurado").send_keys('198290866') #rut sin poliza
#driver.find_element(By.ID, "masterContenido_txtRunAsegurado").send_keys('197394005') #rut con poliza vencida
#driver.find_element(By.ID, "masterContenido_txtRunAsegurado").send_keys('179281872') #rut con poliza vigente hasta 2024
#driver.find_element(By.ID, "masterContenido_txtRunAsegurado").send_keys('242116038') #rut con poliza vigente hasta septiembre 2023

driver.find_element(By.ID, "masterContenido_consultar").click() #Clickea el boton que ejecuta la consulta

#Revisa cual de las dos estructuras de la web corresponde segun la respuesta por rut.
try:
    resultado_poliza = driver.find_element(By.XPATH, "/html/body/div[2]/div/div/div/main/section/div[2]/div/div/form/div[3]/table/tbody/tr[1]/td/h2").text
except NoSuchElementException:
    resultado_poliza = None
if resultado_poliza is None:
# Si no se encontró la estructura1 (rut sin poliza/vencido), intentara buscar el bloque con el segundo XPath que corresponde a la esctructura de ruts con poliza vigente.
    try:
        resultado_poliza = driver.find_element(By.ID, 'masterContenido_LblTituloGrillaConsulta').text
    except NoSuchElementException:
        resultado_poliza = None
        if resultado_poliza == None:
            print('Error desconocido, contactar al desarrollador.')
            driver.quit()

#titulos de tabla segun estado de la poliza
poliza_vigente = 'Resultado Segun Criterio de Selección'
poliza_encontrada_vencida = 'SEGURO OBLIGATORIO COVID'
poliza_no_encontrada = 'NO SE ENCONTRÓ PÓLIZA PARA RUN CONSULTADO'

if resultado_poliza == poliza_encontrada_vencida:
    print("NO")
    fecha_inicio_poliza = driver.find_element(By.XPATH, '//*[@id="masterContenido_lblFechaIni"]').text[:-9].strip()
    fecha_term_poliza = driver.find_element(By.XPATH, '//*[@id="masterContenido_lblFechaFin"]').text[:-9].strip()
    print('inicio: ',fecha_inicio_poliza)
    print('termino: ',fecha_term_poliza)
    
    def dias_faltantes(fecha_inicio_poliza, fecha_term_poliza):
        # Convertir las cadenas de fecha en objetos datetime
        fecha_inicio_poliza_date = datetime.strptime(fecha_inicio_poliza, '%d-%m-%Y').date()
        fecha_term_poliza_date = datetime.strptime(fecha_term_poliza, '%d-%m-%Y').date()

        # Calcular la diferencia entre las fechas
        fecha_actual = datetime.now().date()
        dias_diferencia = (fecha_term_poliza_date - fecha_actual).days

        if dias_diferencia < 0:
            print(f'La poliza venció hace {dias_diferencia} días.')
        elif dias_diferencia > 0: 
            print(f'Quedan {dias_diferencia} días para terminar.')

    dias_faltantes(fecha_inicio_poliza, fecha_term_poliza)
    try:
        volver = driver.find_element(By.ID, "masterContenido_Volver").click() #Volver
    except NoSuchElementException:
        volver = None
        if volver == None:
            driver.find_element(By.XPATH, '//*[@id="masterContenido_btnVolverGrilla"]').click() #Volver

elif resultado_poliza == poliza_no_encontrada:
    print("NO")
    print("-")

    try:
        volver = driver.find_element(By.ID, "masterContenido_Volver").click() #Volver
    except NoSuchElementException:
        volver = None
        if volver == None:
            driver.find_element(By.XPATH, '//*[@id="masterContenido_btnVolverGrilla"]').click() #Volver

elif resultado_poliza == poliza_vigente:
    print("SI")
    fecha_inicio_poliza = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div/main/section/div[2]/div/div/form/div[3]/table[1]/tbody/tr/td/div[2]/table/tbody/tr[2]/td[7]').text[:-8].strip()
    fecha_termino_poliza = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div/main/section/div[2]/div/div/form/div[3]/table[1]/tbody/tr/td/div[2]/table/tbody/tr[2]/td[8]').text[:-8].strip()
    print('inicio: ',fecha_inicio_poliza)
    print('termino: ',fecha_termino_poliza)

    def dias_faltantes(fecha_inicio_poliza, fecha_term_poliza):
        # Convertir las cadenas de fecha en objetos datetime
        fecha_inicio_poliza_date = datetime.strptime(fecha_inicio_poliza, '%d-%m-%Y').date()
        fecha_term_poliza_date = datetime.strptime(fecha_term_poliza, '%d-%m-%Y').date()

        # Calcular la diferencia entre las fechas
        fecha_actual = datetime.now().date()
        dias_diferencia = (fecha_term_poliza_date - fecha_actual).days

        if dias_diferencia < 0:
            print(f'La poliza venció hace {dias_diferencia} días.')
        elif dias_diferencia > 0: 
            print(f'Quedan {dias_diferencia} días para la finalización de la poliza.')

    dias_faltantes(fecha_inicio_poliza, fecha_termino_poliza)
    try:
        volver = driver.find_element(By.ID, "masterContenido_Volver").click() #Volver
    except NoSuchElementException:
        volver = None
        if volver == None:
            driver.find_element(By.XPATH, '//*[@id="masterContenido_btnVolverGrilla"]').click() #Volver

time.sleep(10) #Espera 5 segs dsp del input