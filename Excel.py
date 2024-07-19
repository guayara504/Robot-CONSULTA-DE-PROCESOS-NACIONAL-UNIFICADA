import msvcrt
from re import L
import shutil
import sys
import time
import os
from datetime import datetime
import pyexcel
import xlwt
from Ruta import *

class Excel_1:
    carpetas = ruta()
    dia = time.strftime("%d")
    mes = time.strftime("%m")
    ano = time.strftime("%Y")
    PATH = "C:\\Users\\Study\\Documents\\Universidad\\Grade proyect\\Robot-CONSULTA-DE-PROCESOS-NACIONAL-UNIFICADA-main\\Robot-CONSULTA-DE-PROCESOS-NACIONAL-UNIFICADA-main\\Resultados"

    def __init__(self):
        #Se crea el nombre el nombre de archivo temporal y su extensión
        temporal = "temp-[" + time.strftime("%d-%m-%Y-%H%M%S]") + ".xls"
        self.fnametemp = f'{self.PATH}\\{self.ano}\\{self.carpetas.dife_fecha()}\\{self.dia}\\{temporal}'
        #self.fnametemp = "temp-[" + time.strftime("%d-%m-%Y-%H%M%S]") + ".xls"   
    
    
    #Funcion para crear el archivo de excel
    def crear_xls(self,wb):
        
        #Diccionario en el que se crea la cantidad de hojas que habrá en el archivo y sus columnas que habrá en cada una
        data = {'INPUT': ['CLIENTE','RADICADO'],
                'DATOS DEL PROCESO': ['CLIENTE', 'RADICADO' ,'FECHA RADICACION', 'DESPACHO', 'PONENTE', 'TIPO', 'CLASE', 'RECURSO', 'UBICACION','CONTENIDO','SUJETOS PROCESALES'],
                'ACTUACIONES DEL PROCESO': ['RADICADO', 'FECHA ACTUACION', 'ACTUACION', 'ANOTACION', 'FECHA INICIA TERMINO', 'FECHA FIN TERMINO','FECHA REGISTRO'],
                'ERRORES': ['CC','RADICADO']}
        
        #Se agrega la hoja
        for key, nomHoja in enumerate(data):
            ws = wb.add_sheet(nomHoja)
            #Se agrega las columnas a cada hoja
            for clave, valor in enumerate(data[nomHoja]):
                ws.write(0, clave, valor)
        #Se guarda el archivo
        wb.save(self.fnametemp)
        
    #Funcion para escribir en el excel creado anteriormente
    def escribir_xls(self,datos=None, actuaciones=None,i=None,errores=None):
        #Si el proceso estuvo exito se ejecuta esta instruccion
        if errores == None:
            #Se extrae los datos y se guarda como lista
            datos = list(datos.values())
            #Se obtiene el archivo de excel para escribir
            wb = pyexcel.get_book(file_name=self.fnametemp)
            #Se agregan los datos iniciales de el proceso
            wb.sheet_by_name('INPUT').row += i
            #Se agregan todos los datos de el proceso
            wb.sheet_by_name('DATOS DEL PROCESO').row += datos
            #Se agregan las actuaciones de el proceso
            for dats in actuaciones:
                wb.sheet_by_name('ACTUACIONES DEL PROCESO').row += dats
        #Si el proceso obtuvo error se ejecuta esta instruccion
        else:
            #Se extrae los datos y se guarda como lista
            errores = list(errores.values())
            #Se obtiene el archivo de excel para escribir
            wb = pyexcel.get_book(file_name=self.fnametemp)
            #Se agrega el proceso a la hoja errores
            wb.sheet_by_name('ERRORES').row += errores
            
        #Se guarda el archivo   
        wb.save_as(self.fnametemp)
    
    #Funcion en la que se guarda el archivo ya finalizado
    def terminar(self,excelFile):
        # cambiar el nombre del temp.xls y eliminarlo
        fnamefinal = excelFile.split(".")[0] + "-[" + time.strftime("%d-%m-%Y-%H%M%S]") + ".xls"
        fname = f'{self.PATH}\\{self.ano}\\{self.carpetas.dife_fecha()}\\{self.dia}\\{fnamefinal}'
        shutil.move((self.fnametemp), fname)
        print("\nCreado el archivo: " + fname)
    
    #Funcion en la que se verifica si la actuacion es menor a 4
    def dife_fecha(fecha):
        hoy = datetime.now()
        ano,mes,dia = fecha.split("-")
        fecha_str = dia + ' ' + mes + ' ' + ano
        dt_obj = datetime.strptime(fecha_str, '%d %m %Y')
        return (hoy - dt_obj)
    
    
    
    
    
    
    
    
    