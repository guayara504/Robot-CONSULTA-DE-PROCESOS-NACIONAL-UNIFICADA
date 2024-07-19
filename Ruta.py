import msvcrt
from re import L
import shutil
import sys
import time
import os
from datetime import datetime
import pyexcel
import xlwt
from Excel import *

class ruta:
    PATH = "C:\\Users\\Study\\Documents\\Universidad\\Grade proyect\\Robot-CONSULTA-DE-PROCESOS-NACIONAL-UNIFICADA-main\\Robot-CONSULTA-DE-PROCESOS-NACIONAL-UNIFICADA-main\\Resultados"
    
    def __init__(self):
        pass
    
    def dife_fecha(self):
        mes= time.strftime("%m")
        if mes == '01':
            mes = 'ENERO'
        if mes == '02':
            mes = 'FEBRERO'
        if mes == '03':
            mes = 'MARZO'
        if mes == '04':
            mes = 'ABRIL'
        if mes == '05':
            mes = 'MAYO'
        if mes == '06':
            mes = 'JUNIO'
        if mes == '07':
            mes = 'JULIO'
        if mes == '08':
            mes = 'AGOSTO'
        if mes == '09':
            mes = 'SEPTIEMBRE'
        if mes == '10':
            mes = 'OCTUBRE'
        if mes == '11':
            mes = 'NOVIEMBRE'
        if mes == '12':
            mes = 'DICIEMBRE'
        return mes
    
    def crear_carpetas(self,dia,mes,ano):
        try:
                os.mkdir(f'{self.PATH}\\{ano}')
        except:
                pass

        mes= self.dife_fecha()
        try:
                os.mkdir(f'{self.PATH}\\{ano}\\{mes}')
        except:
                pass

        try:
                os.mkdir(f'{self.PATH}\\{ano}\\{mes}\\{dia}')
        except:
                pass