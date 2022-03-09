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

class Driver_1(object):

    op = webdriver.ChromeOptions()
    op.add_experimental_option('excludeSwitches', ['enable-logging'])
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=op)


    def __init__(self):
        #Escoger preferencias del WebDriver
        self.delay = 4
        self.driver.set_window_size(1024, 768)
        self.driver.implicitly_wait(3)
        self.load_page()
        
    #Cargar la pagina solicitada

    def load_page(self):
        wait = WebDriverWait(self.driver, self.delay)
        self.driver.get("https://consultaprocesos.ramajudicial.gov.co/Procesos/NumeroRadicacion/")
        os.system ("cls")
        
        try:
            wait.until(self.driver.find_element(By.XPATH,'//*[@id="mainContent"]/div/div/div/div[1]/div/div[2]/span'))
        except TimeoutException:
            print('line: 38 error: No se cargo la pagina, TimeoutException')
    
    

    #Espera por el elemento
    def WaitForElement(self,path):
        demora = 3
        limit = demora
        inc = 1
        c = 0
        while c < limit:
            try:
                self.driver.find_element(By.XPATH,path)
                return 1
            except:
                time.sleep(inc)
                c+=inc
        return 0