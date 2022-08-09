from UnixExcel import *
import msvcrt
from re import L
import shutil
import sys
import time
import os
from datetime import datetime
from webbrowser import Chrome
import pyexcel
import xlwt
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, ElementNotVisibleException, NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from Consultor import *
from Driver import *
from Excel import Excel_1

class Extractor_1(object):


    #Extrae los datos del proceso
    def extraer_datos(datos,driver,i):
        WebDriverWait(driver,5).until(EC.presence_of_element_located((By.XPATH,'//*[@id="mainContent"]/div/div/div/div[2]/div/div/div[2]/div[2]/div[2]/div/div/div/div/div/div/table/tbody/div/tr/th/tr[1]/th')))
        #Se agregan los datos del procesos al diccionario
        datos["cliente"] = i[0]
        datos["radicado"] = i[1]
        datos["fecha"] = (driver.find_element(By.XPATH,'//*[@id="mainContent"]/div/div/div/div[2]/div/div/div[2]/div[2]/div[2]/div/div/div/div/div/div/table/tbody/div/tr/th[1]/tr[1]/td').text)
        datos["despacho"] = (driver.find_element(By.XPATH,'//*[@id="mainContent"]/div/div/div/div[2]/div/div/div[2]/div[2]/div[2]/div/div/div/div/div/div/table/tbody/div/tr/th/tr[2]/td').text)
        datos["ponente"] = (driver.find_element(By.XPATH,'//*[@id="mainContent"]/div/div/div/div[2]/div/div/div[2]/div[2]/div[2]/div/div/div/div/div/div/table/tbody/div/tr/th/tr[3]/td').text)
        datos["tipo_proceso"] = (driver.find_element(By.XPATH,'//*[@id="mainContent"]/div/div/div/div[2]/div/div/div[2]/div[2]/div[2]/div/div/div/div/div/div/table/tbody/div/tr/th/tr[4]/td').text)
        datos["clase_proceso"] = (driver.find_element(By.XPATH,'//*[@id="mainContent"]/div/div/div/div[2]/div/div/div[2]/div[2]/div[2]/div/div/div/div/div/div/table/tbody/div/tr/th/tr[5]/td').text)
        datos["recurso"] = (driver.find_element(By.XPATH,'//*[@id="mainContent"]/div/div/div/div[2]/div/div/div[2]/div[2]/div[2]/div/div/div[1]/div/div/div/table/tbody/div/tr/th[2]/tr[1]/td').text)
        datos["ubicacion"] = (driver.find_element(By.XPATH,'//*[@id="mainContent"]/div/div/div/div[2]/div/div/div[2]/div[2]/div[2]/div/div/div[1]/div/div/div/table/tbody/div/tr/th[2]/tr[2]/td').text)
        datos["contenido"] = (driver.find_element(By.XPATH,'//*[@id="mainContent"]/div/div/div/div[2]/div/div/div[2]/div[2]/div[2]/div/div/div[1]/div/div/div/table/tbody/div/tr/th[2]/tr[4]/th').text)
        
    #Da click en la pestaña "partes del proceso"
    def extraer_partes(datos,driver):
        driver.find_element(By.XPATH,'//*[@id="mainContent"]/div/div/div/div[2]/div/div/div[2]/div[2]/div[1]/div/div/div[2]/div/div[3]').click()
               
    #Extrae las actuaciones del proceso
    def extraer_actuaciones(actuaciones,driver,radicado,inicioBusqueda):
        #Da click en la pestaña "actuaciones"
        driver.find_element(By.XPATH,'//*[@id="mainContent"]/div/div/div/div[2]/div/div/div[2]/div[2]/div[1]/div/div/div[2]/div/div[5]').click()
        WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="mainContent"]/div/div/div/div[2]/div/div/div[2]/div[2]/div[2]/div/div/div[3]/div/div[1]/div[1]')))
        #Se obtiene la tabla en la que se encuentra las actuaciones
        table_actos = driver.find_element(By.XPATH,'//*[@id="mainContent"]/div/div/div/div[2]/div/div/div[2]/div[2]/div[2]/div/div/div[3]/div/div[1]/div[2]/div/table/tbody')
        WebDriverWait(driver,5).until(EC.presence_of_element_located((By.XPATH,'//*[@id="mainContent"]/div/div/div/div[2]/div/div/div[2]/div[2]/div[2]/div/div/div[3]/div/div[1]/div[2]/div/table/tbody/tr[1]/td[1]')))
        #Se da la espera que se cargue la tabla actuaciones
        time.sleep(2)
        
        #Si la busqueda se realizará en los ultimos 4 dias se ejecutará esta instrucción
        if inicioBusqueda == "2": 
            #Se obtiene las utimas 7 filas de la tabla para hacer la verificación
            allrows = table_actos.find_elements(By.TAG_NAME,"tr")[:7]
            #Se recorre cada actuacion de la lista
            for tr in allrows:
                #Lista que guardará cada parte de la actuación por separado
                lista_td = []
                #Se agrega el radicado como primer elemento a la lista
                lista_td.append(radicado)
                #Se obtiene las partes de la actuacion
                allcols = tr.find_elements(By.TAG_NAME,"td")[:-1]
                #Se obtiene la fecha
                fecha_str=allcols[0].text
                #Comparación para verificar si la actuacion es de los ultimos 4 dias y que no esté privada
                if fecha_str != "" and Excel_1.dife_fecha(fecha_str).days <= 4:
                    #Si la actuacion es de los ultimos 4 dias se recorre la lista
                    for j in range(len(allcols)):
                        #Se agrega las partes de la actuacion a la lista
                        lista_td.append(allcols[j].text)
                    #Se agrega la actuación a la lista para agregar al excel
                    actuaciones.append(lista_td)
         #Si la busqueda se realizará en modo historico se ejecutará esta instrucción
        else:
            try:
                #Se encuentra el objeto donde aparecen la cantidad de paginas que hay en las actuaciones
                CuadroPaginas = driver.find_element(By.XPATH,'//*[@id="mainContent"]/div/div/div/div[2]/div/div/div[2]/div[2]/div[2]/div/div/div[2]/div/div[1]/div[3]/div/ul')
                #Se obtiene la cantidad de paginas que hay
                CantidadPaginas = CuadroPaginas.find_elements(By.TAG_NAME,"li")[1:-1]
            except:
                #Si no exiten mas hojas se instancia en 1
                CantidadPaginas = [1]
            #Si la cantidad de paginas es mayor a 1 se ejecuta esta instrucción
            if len(CantidadPaginas) > 1:
                
                #Se ejecutará esta instruccion para cada pagina encontrada
                for pagina in range(len(CantidadPaginas)):
                    pagina +=2
                    try:
                        #Se dará click en la pagina seleccionada
                        driver.find_element(By.XPATH,'//*[@id="mainContent"]/div/div/div/div[2]/div/div/div[2]/div[2]/div[2]/div/div/div[3]/div/div[1]/div[3]/div/ul/li['+str(pagina)+']/button').click()
                    except:
                        pass
                    WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="mainContent"]/div/div/div/div[2]/div/div/div[2]/div[2]/div[2]/div/div/div[3]/div/div[1]/div[1]')))
                    #Se obtiene la tabla en la que se encuentra las actuaciones
                    table_actos = driver.find_element(By.XPATH,'//*[@id="mainContent"]/div/div/div/div[2]/div/div/div[2]/div[2]/div[2]/div/div/div[3]/div/div[1]/div[2]/div/table/tbody')
                    WebDriverWait(driver,5).until(EC.presence_of_element_located((By.XPATH,'//*[@id="mainContent"]/div/div/div/div[2]/div/div/div[2]/div[2]/div[2]/div/div/div[3]/div/div[1]/div[2]/div/table/tbody/tr[1]/td[1]')))
                    #Se da la espera que se cargue la tabla actuaciones
                    time.sleep(2)
                    #Se obtiene todas las filas de la tabla
                    allrows = table_actos.find_elements(By.TAG_NAME,"tr")[:]
                    #Se recorre cada actuacion de la lista
                    for tr in allrows:
                            #Lista que guardará cada parte de la actuación por separado
                            lista_td = []
                            #Se agrega el radicado como primer elemento a la lista
                            lista_td.append(radicado)
                            #Se obtiene las partes de la actuacion
                            allcols = tr.find_elements(By.TAG_NAME,"td")[:-1]
                            #se recorre la lista de cada actuacion
                            for j in range(len(allcols)):
                                #Se agrega las partes de la actuacion a la lista
                                lista_td.append(allcols[j].text)
                            #Se agrega la actuación a la lista para agregar al excel
                            actuaciones.append(lista_td)
                #La pagina al tener conocimiento de automatizaciones se da cuenta de estos y se procede a refrescar la pagina cada vez que se demora mucho tiempo en un proceso
                driver.refresh()
                WebDriverWait(driver,timeout=4).until(EC.presence_of_element_located((By.XPATH,'//*[@id="mainContent"]/div/div/div/div[1]/div/div[2]/span')))
            #Si la cantidad de paginas es igual a 1 se ejecuta esta instrucción
            else:   
                    #Se obtiene todas las filas de la tabla
                    allrows = table_actos.find_elements(By.TAG_NAME,"tr")[:]
                    #Se recorre cada actuacion de la lista
                    for tr in allrows:
                            #Lista que guardará cada parte de la actuación por separado
                            lista_td = []
                            #Se agrega el radicado como primer elemento a la lista
                            lista_td.append(radicado)
                            #Se obtiene las partes de la actuacion
                            allcols = tr.find_elements(By.TAG_NAME,"td")[:-1]
                            #se recorre la lista de cada actuacion
                            for j in range(len(allcols)):
                                #Se agrega las partes de la actuacion a la lista
                                lista_td.append(allcols[j].text)
                            #Se agrega la actuación a la lista para agregar al excel
                            actuaciones.append(lista_td)
                
         
        
