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

class Consultor_1(object):

    #Ingresar a "Todos los procesos"
    def iniciar_busqueda(driver):
        #Click en boton "Todos los procesos"
        driver.find_element(By.XPATH,'//*[@id="mainContent"]/div/div/div/div[1]/div/form/div[1]/div[1]/div/div/div/div[1]/div/div[2]/div').click()

    #Ingresar el radicado en la caja
    def ingresar_radicado(driver):
        #Pegar Radicado
        driver.find_element(By.XPATH,"//input[@maxlength='23']").clear()
        driver.find_element(By.XPATH,"//input[@maxlength='23']").send_keys("76520400300120190015700")
        #Click en boton 'Consulta'
        driver.find_element(By.XPATH,'//*[@id="mainContent"]/div/div/div/div[1]/div/form/div[2]/button[1]/span').click()

    def ventana_emergente(driver):
        WebDriverWait(driver,timeout=4).until(EC.presence_of_element_located((By.XPATH,'//*[@id="app"]/div[3]/div/div/div[2]/div/button/span')))
        ventana = driver.find_element(By.XPATH,'//*[@id="app"]/div[3]/div/div/div[2]/div/button/span').is_displayed()
        return ventana
        
    def click_Proceso(driver):
        #Ingresar a datos del proceso
        WebDriverWait(driver,timeout=4).until(EC.presence_of_element_located((By.XPATH,'//*[@id="mainContent"]/div/div/div/div[2]/div/div/div[2]/div/table/tbody/tr/td')))
        driver.find_element(By.XPATH,'//*[@id="mainContent"]/div/div/div/div[2]/div/div/div[2]/div/table/tbody/tr/td[2]').click()
        driver.find_element(By.XPATH,'//*[@id="mainContent"]/div/div/div/div[2]/div/div/div[2]/div/table/tbody/tr/td[2]').click()

