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

class Driver_1(object):

    #Declaracion de Variables
    op = webdriver.ChromeOptions()
    #op.add_argument("--disable-gpu")
    op.add_experimental_option('excludeSwitches', ['enable-logging','enable-automation'] )
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=op)

    #Escoger preferencias del WebDriver
    def __init__(self):  
        self.driver.set_window_size(1024, 768)
        self.load_page()
        
    #Cargar la pagina solicitada
    def load_page(self):

        try:
            #Se obtiene la pagina a consultar
            self.driver.get("https://consultaprocesos.ramajudicial.gov.co/Procesos/NumeroRadicacion/")
            os.system ("cls")
            WebDriverWait(self.driver,timeout=4).until(EC.presence_of_element_located((By.XPATH,'//*[@id="mainContent"]/div/div/div/div[1]/div/div[2]/span')))
        except TimeoutException:
            print('line: 38 error: No se cargo la pagina, TimeoutException')
        except:
            print('Error Interno')
    
    #Funcion para parar la ejecución del driver
    def close(self):
        self.driver.quit()