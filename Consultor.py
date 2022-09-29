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
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

class Consultor_1():

    #Ingresar a "Todos los procesos"
    def iniciar_busqueda(driver):
        #Click en boton "Todos los procesos"
        driver.find_element(By.XPATH,'//*[@id="mainContent"]/div/div/div/div[1]/div/form/div[1]/div[1]/div/div/div/div[1]/div/div[2]/div').click()

    #Ingresar el radicado en la caja
    def ingresar_radicado(driver,radicado):
        #Pegar Radicado
        driver.find_element(By.XPATH,'//*[@id="input-72"]').send_keys(Keys.CONTROL + "a")
        driver.find_element(By.XPATH,'//*[@id="input-72"]').send_keys(Keys.DELETE)
        driver.find_element(By.XPATH,'//*[@id="input-72"]').send_keys(radicado)
        #Click en boton 'Consulta'
        driver.find_element(By.XPATH,'//*[@id="mainContent"]/div/div/div/div[1]/div/form/div[2]/button[1]/span').click()


    def click_Proceso(driver,datos):
        #Dar click en boton de "volver" en ventana flotante
        try:
            WebDriverWait(driver,5).until(EC.presence_of_element_located((By.XPATH,'//*[@id="app"]/div[3]/div/div')))
            driver.find_element(By.XPATH,'//*[@id="app"]/div[3]/div/div/div[2]/div/button/span').click()
        except TimeoutException:
            pass
        
        WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="mainContent"]/div/div/div/div[2]/div/div/div[2]')))
        
        #En esta parte de la funcion se ingresará al proceso y, al haber mas de un proceso mostrado en pantalla se eligirá el mas actualizado
        
        #Se obtiene la tabla de los procesos
        table_proceso = driver.find_element(By.XPATH,'//*[@id="mainContent"]/div/div/div/div[2]/div/div/div[2]/div/table/tbody')
        #Se obtiene la cantidad de procesos mostrados
        allrows = table_proceso.find_elements(By.TAG_NAME,"tr")[:]
        #Variable que guarda la fecha mayor que se muestra
        fechaMayor = 0
        #Variable que guarda la posicion de la fecha mayor para poder ingresar
        index = 0
        
        #Se recorre los procesos para encontrar el mas actualizado e ingresar
        for tr in range(len(allrows)):
            tr = tr+1
            #Se guarda la fecha de el proceso  y en caso de estar privado se guarda con el valor de 0
            try:
                fecha = int((driver.find_element(By.XPATH,'//*[@id="mainContent"]/div/div/div/div[2]/div/div/div[2]/div/table/tbody/tr['+str(tr)+']/td[3]/div/button/span').text).replace("-",""))
            except:
                fecha = 0
            #Se guarda en la variable fechaMayor el proceso con la fecha mas actualizada y se asignara el numero de el index para poder ingresar
            if fecha > fechaMayor:
                fechaMayor = fecha
                index = tr
        
        #Se guarda los sujetos procesales que se encuentran al lado derecho de los procesos mostrados
        datos["sujetos_procesales"] = (driver.find_element(By.XPATH,'//*[@id="mainContent"]/div/div/div/div[2]/div/div/div[2]/div/table/tbody/tr['+str(index)+']/td[5]/div').text)
        
        #Se da click en el proceso con la fecha mas actualizada
        driver.find_element(By.XPATH,'//*[@id="mainContent"]/div/div/div/div[2]/div/div/div[2]/div/table/tbody/tr['+str(index)+']/td[2]/button/span').click()

            

