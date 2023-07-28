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

for row, datos in df.iterrows():
    print(f'Inserta rut del indice {row+2}------------------------------------------------------------------------------')
    driver.find_element(By.ID, "masterContenido_txtRunAsegurado").send_keys(datos.RUT) # Inserta el rut de la fila iterada

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
        estado_poliza = 'NO'
        fecha_inicio_poliza = driver.find_element(By.XPATH, '//*[@id="masterContenido_lblFechaIni"]').text[:-9].strip()
        fecha_term_poliza = driver.find_element(By.XPATH, '//*[@id="masterContenido_lblFechaFin"]').text[:-9].strip()
        
        def dias_faltantes(fecha_inicio_poliza, fecha_term_poliza):
            # Convertir las cadenas de fecha en objetos datetime
            fecha_inicio_poliza_date = datetime.strptime(fecha_inicio_poliza, '%d-%m-%Y').date()
            fecha_term_poliza_date = datetime.strptime(fecha_term_poliza, '%d-%m-%Y').date()

            # Calcular la diferencia entre las fechas
            fecha_actual = datetime.now().date()
            dias_diferencia = (fecha_term_poliza_date - fecha_actual).days

            if dias_diferencia < 0:
                comentario = 'La poliza venció hace {dias_diferencia} días.'
                print(comentario)
            elif dias_diferencia > 0: 
                comentario = 'Quedan {dias_diferencia} días para terminar.'
                print(comentario)

        dias_faltantes(fecha_inicio_poliza, fecha_term_poliza)

        print('Estado poliza: ', estado_poliza)
        print('inicio: ',fecha_inicio_poliza)
        print('termino: ',fecha_term_poliza)
        print('comentario: ',comentario)

        df.at[row, 'estado_seguro'] = estado_poliza
        df.at[row, 'fecha_inicio'] = fecha_inicio_poliza
        df.at[row, 'fecha_termino'] = fecha_term_poliza
        df.at[row, 'comentario'] = comentario

        #vuelta
        try:
            volver = driver.find_element(By.ID, "masterContenido_Volver").click() #Volver
        except NoSuchElementException:
            volver = None
            if volver == None:
                driver.find_element(By.XPATH, '//*[@id="masterContenido_btnVolverGrilla"]').click() #Volver

    elif resultado_poliza == poliza_no_encontrada:
        estado_poliza = 'NO'
        fecha_inicio_poliza = '-'
        fecha_term_poliza = '-'
        comentario = 'NO APLICA'

        print('Estado poliza: ', estado_poliza)
        print('inicio: ',fecha_inicio_poliza)
        print('termino: ',fecha_term_poliza)
        print('comentario: ',comentario)

        df.at[row, 'estado_seguro'] = estado_poliza
        df.at[row, 'fecha_inicio'] = fecha_inicio_poliza
        df.at[row, 'fecha_termino'] = fecha_term_poliza
        df.at[row, 'comentario'] = comentario

        #vuelta
        try:
            volver = driver.find_element(By.ID, "masterContenido_Volver").click() #Volver
        except NoSuchElementException:
            volver = None
            if volver == None:
                driver.find_element(By.XPATH, '//*[@id="masterContenido_btnVolverGrilla"]').click() #Volver

    elif resultado_poliza == poliza_vigente:
        estado_poliza = 'SI'
        fecha_inicio_poliza = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div/main/section/div[2]/div/div/form/div[3]/table[1]/tbody/tr/td/div[2]/table/tbody/tr[2]/td[7]').text[:-8].strip()
        fecha_termino_poliza = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div/main/section/div[2]/div/div/form/div[3]/table[1]/tbody/tr/td/div[2]/table/tbody/tr[2]/td[8]').text[:-8].strip()
        
        print('fecha termino extraida: ', fecha_termino_poliza)

        def dias_faltantes(fecha_inicio_poliza, fecha_term_poliza):
            # Convertir las cadenas de fecha en objetos datetime
            fecha_inicio_poliza_date = datetime.strptime(fecha_inicio_poliza, '%d-%m-%Y').date()
            fecha_term_poliza_date = datetime.strptime(fecha_term_poliza, '%d-%m-%Y').date()
            print('fecha termino date: ', fecha_term_poliza_date)
            # Calcular la diferencia entre las fechas
            fecha_actual = datetime.now().date()
            dias_diferencia = (fecha_term_poliza_date - fecha_actual).days

            if dias_diferencia < 0:
                comentario = 'La poliza venció hace '+str(dias_diferencia)+' días.'
                # print(comentario)
            elif dias_diferencia > 0: 
                comentario = 'Quedan '+str(dias_diferencia)+' días para terminar.'
                # print(comentario)


        print('Estado poliza: ', estado_poliza)
        print('inicio: ',fecha_inicio_poliza)
        print('termino: ',fecha_term_poliza)
        dias_faltantes(fecha_inicio_poliza, fecha_termino_poliza)
        print('comentario: ',comentario)

        df.at[row, 'estado_seguro'] = estado_poliza
        df.at[row, 'fecha_inicio'] = fecha_inicio_poliza
        df.at[row, 'fecha_termino'] = fecha_term_poliza
        df.at[row, 'comentario'] = comentario

        #vuelta
        try:
            volver = driver.find_element(By.ID, "masterContenido_Volver").click() #Volver
        except NoSuchElementException:
            volver = None
            if volver == None:
                driver.find_element(By.XPATH, '//*[@id="masterContenido_btnVolverGrilla"]').click() #Volver
    print(f'Cierre de consulta del indice {row+2}------------------------------------------------------------------------------')
    time.sleep(5) #Espera 5 segs entre iteracion
    df.to_excel('bbdd.xlsx', index=False)