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
        print("el radicado es "+radicado)
        #Pegar Radicado
        driver.find_element(By.XPATH,'//*[@id="input-73"]').send_keys(Keys.CONTROL + "a")
        driver.find_element(By.XPATH,'//*[@id="input-73"]').send_keys(Keys.DELETE)
        driver.find_element(By.XPATH,'//*[@id="input-73"]').send_keys(radicado)
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
        table_proceso = driver.find_element(By.XPATH,'//*[@id="mainContent"]/div/div/div/div[2]/div/div/div[2]/div/table/tbody')
        allrows = table_proceso.find_elements(By.TAG_NAME,"tr")[:]
        fechaMayor = 0
        index = 0
        for tr in range(len(allrows)):
            tr = tr+1
            try:
                fecha = int((driver.find_element(By.XPATH,'//*[@id="mainContent"]/div/div/div/div[2]/div/div/div[2]/div/table/tbody/tr['+str(tr)+']/td[3]/div/button/span').text).replace("-",""))
                print(fecha)
            except:
                fecha = 0
            if fecha > fechaMayor:
                fechaMayor = fecha
                index = tr
                
        datos["sujetos_procesales"] = (driver.find_element(By.XPATH,'//*[@id="mainContent"]/div/div/div/div[2]/div/div/div[2]/div/table/tbody/tr['+str(index)+']/td[5]/div').text)
        
        driver.find_element(By.XPATH,'//*[@id="mainContent"]/div/div/div/div[2]/div/div/div[2]/div/table/tbody/tr['+str(index)+']/td[2]/button/span').click()

            

