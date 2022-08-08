import msvcrt
from re import L
import shutil
import sys
import time
import os
from datetime import datetime
import pyexcel
import xlwt

class Excel_1:

    def __init__(self):
        self.fnametemp = "temp-[" + time.strftime("%d-%m-%Y-%H%M%S]") + ".xls"   
    
    

    def crear_xls(self,wb):
        data = {'INPUT': ['CLIENTE','RADICADO'],
                'DATOS DEL PROCESO': ['CLIENTE', 'RADICADO' ,'FECHA RADICACION', 'DESPACHO', 'PONENTE', 'TIPO', 'CLASE', 'RECURSO', 'UBICACION','CONTENIDO','SUJETOS PROCESALES'],
                'ACTUACIONES DEL PROCESO': ['RADICADO', 'FECHA ACTUACION', 'ACTUACION', 'ANOTACION', 'FECHA INICIA TERMINO', 'FECHA FIN TERMINO','FECHA REGISTRO'],
                'ERRORES': ['CC','RADICADO']}
        
        for key, nomHoja in enumerate(data):
            ws = wb.add_sheet(nomHoja)
            for clave, valor in enumerate(data[nomHoja]):
                ws.write(0, clave, valor)
        wb.save(self.fnametemp)

    def escribir_xls(self,datos=None, actuaciones=None,i=None,errores=None):
        if errores == None:
            
            datos = list(datos.values())
            wb = pyexcel.get_book(file_name=self.fnametemp)
            wb.sheet_by_name('INPUT').row += i
            wb.sheet_by_name('DATOS DEL PROCESO').row += datos
            for dats in actuaciones:
                wb.sheet_by_name('ACTUACIONES DEL PROCESO').row += dats
        else:
            errores = list(errores.values())
            wb = pyexcel.get_book(file_name=self.fnametemp)
            wb.sheet_by_name('ERRORES').row += errores
            
        wb.save_as(self.fnametemp)
    
    def terminar(self,excelFile):
        # cambiar el nombre del temp.xls y eliminarlo
        fname = excelFile.split(".")[0] + "-[" + time.strftime("%d-%m-%Y-%H%M%S]") + ".xls"   
        shutil.move((self.fnametemp), fname)
        print("\nCreado el archivo: " + fname)
    
    def dife_fecha(fecha):
        hoy = datetime.now()
        ano,mes,dia = fecha.split("-")
        fecha_str = dia + ' ' + mes + ' ' + ano
        dt_obj = datetime.strptime(fecha_str, '%d %m %Y')
        return (hoy - dt_obj)
    
    
    
    
    
    
    
    
    