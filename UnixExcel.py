from Consultor import *
from Driver import *
from Extractor import *
from Excel import Excel_1
import sys
import os
import msvcrt

#Clase principal   
if __name__ == "__main__":
    os.system ("cls")
    #Escoger los archivos con los que se trabajara
    inputFile=input("\033[1;34m\n\nIngrese archivo(s) de Excel separado por comas: ")
    listFile=inputFile.split(",")
    os.system ("cls")
    inicioBusqueda=input("\033[1;34m\n\n1.Historico\n2.Recientes\nIngrese el tipo de busqueda: ")
    os.system ("cls")
    #Se llaman las clases a trabajar
    browser =Driver_1()
    consulta =Consultor_1
    extractor = Extractor_1
    excel = Excel_1()
    
    if (inicioBusqueda == "1") or (inicioBusqueda == "2"):
        
          #Hacer esto con todos los archivos ingresados
        for file in listFile:
                  excelFile= file+".xlsx"
        
                  print(f'\033[1;32m\nEjecutando: {file}\n')
        
                  # extraer los datos del archivo de entrada
                  my_array = pyexcel.get_array(file_name=excelFile, start_row=1)
                  #Creamos la hoja de excel y la mandamos como parametro
                  wout = xlwt.Workbook()
                  excel.crear_xls(wb=wout)
                  conteoReLoad = 0
                  numerojuzgado = 0
                  for i in my_array:
                    if i[0] == "": break
                    numerojuzgado += 1
                    
                    try:
                          #Lista donde guardara las actuaciones del proceso
                            actuaciones = []
                            #Lista donde guardara el proceso que arrojo error
                            errores = {"cliente":"","radicado":""}
                            #Diccionario para guardar los datos del proceso
                            datos = {"cliente":"","radicado":"","fecha": "","despacho":"","ponente":"","tipo_proceso":"","clase_proceso":"","recurso":"","ubicacion":"","contenido":"","sujetos_procesales":""}
                            #Damos click en boton "TODOS LOS PROCESOS"
                            consulta.iniciar_busqueda((browser.driver))
                            #Se ingresa el radicado a consultar
                            consulta.ingresar_radicado((browser.driver),i[1])
                            #Se selecciona el proceso a consultar
                            consulta.click_Proceso((browser.driver),datos)
                            #Se extrae los datos del proceso
                            extractor.extraer_datos(datos,(browser.driver),i)
                            #Se extrae las partes interesadas en el proceso
                            extractor.extraer_partes(datos,(browser.driver))
                            #Se extrae las actuaciones del proceso
                            extractor.extraer_actuaciones(actuaciones,(browser.driver),i[1],inicioBusqueda)
                            #Se imprime en pantalla los datos
                            print("\033[1;36mRadicado: ",f"\033[1;34m{i[1]} #{numerojuzgado}")
                            #Se ingresa la informacion al excel
                            excel.escribir_xls(datos=datos,actuaciones=actuaciones,i=i)
                    except:
                          conteoReLoad += 1
                          print("\033[1;33m*"*40)
                          print(f"\033[1;31mERROR: {i[1]} #{numerojuzgado}")
                          print("\033[1;33m*"*40)
                          errores["cliente"] = i[0]
                          errores["radicado"] = i[1]
                          if conteoReLoad == 1:
                              conteoReLoad = 0
                              resp = input("\033[1;36m¿Volver a recargar?\n1.si\n2.no\nrespuesta \2: ")
                              os.system ("cls")
                              while resp == "1":
                                  print(f'\033[1;32m\nEjecutando: {file}\n')
                                  browser.driver.refresh()
                                  resp = input("\033[1;36m¿Volver a recargar?\n1.si\n2.no\nrespuesta \2: ")
                                  os.system ("cls")
                                  print(f'\033[1;32m\nEjecutando: {file}\n')
                              WebDriverWait(browser.driver,timeout=4).until(EC.presence_of_element_located((By.XPATH,'//*[@id="mainContent"]/div/div/div/div[1]/div/div[2]/span')))
                          excel.escribir_xls(errores=errores)
                        
                  excel.terminar(file)
                  print("\n","\033[1;36m*"*20,"\n","\033[1;36mFinalizó ",file,"\n","\033[1;36m*"*20,"\n")
        #Se Termina el programa 
        print("\n","\033[1;36m*"*20,"\n","\033[1;36mFinalizó la ejecución","\n","\033[1;36m*"*20,"\n")
        browser.close()
    else:
            print("\033[1;36mUsar\n1.Historico = Hacer historico de los procesos\n2.Recientes = Para obtener los ultimos movimientos publicados")
            msvcrt.getch()
            sys.exit(1)
    print("\033[1;36m\n----------------------\n")  
    print("\033[1;36mPresiona una tecla para cerrar")
    msvcrt.getch()
    sys.exit(1)                    

    
        
        
        

        
