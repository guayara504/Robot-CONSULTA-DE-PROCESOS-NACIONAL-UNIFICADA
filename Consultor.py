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
from Driver import *
from Extractor import *


class Consultor_1(object):
    
    #Ingresar al proceso en la pagina
    def iniciar_busqueda(driver):
        #Click en arrojar todas las actuaciones
        driver.find_element(By.XPATH,'//*[@id="mainContent"]/div/div/div/div[1]/div/form/div[1]/div[1]/div/div/div/div[1]/div/div[2]/div').click()
    def ingresar_radicado(driver):
        #Pegar Radicado
        driver.find_element(By.XPATH,"//input[@maxlength='23']").clear()
        driver.find_element(By.XPATH,"//input[@maxlength='23']").send_keys("76001333300120210015600")
        #Click en boton 'Consulta'
        driver.find_element(By.XPATH,'//*[@id="mainContent"]/div/div/div/div[1]/div/form/div[2]/button[1]/span').click()
        #Ingresar a datos del proceso
        if driver.WaitForElement('//*[@id="mainContent"]/div/div/div/div[2]/div/div/div[2]'):
            driver.find_element(By.XPATH,'//*[@id="mainContent"]/div/div/div/div[2]/div/div/div[2]/div/table/tbody/tr/td[2]').click()



    
