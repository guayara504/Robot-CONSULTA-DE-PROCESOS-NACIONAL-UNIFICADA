from Consultor import *
from Driver import *
from Extractor import *
from Ruta import *
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
    #La busqueda se hará en tipo historico o los ultimos 4 dias
    inicioBusqueda=input("\033[1;34m\n\n1.Historico\n2.Recientes\nIngrese el tipo de busqueda: ")
    os.system ("cls")
    #Se llaman las clases a trabajar
    browser =Driver_1()
    consulta =Consultor_1
    extractor = Extractor_1
    excel = Excel_1()
    carpetas = ruta()
    
   
    carpetas.crear_carpetas(dia=excel.dia, mes=excel.mes, ano=excel.ano)

    if (inicioBusqueda == "1") or (inicioBusqueda == "2"):
        
        #Hacer busqueda con todos los archivos ingresados
        for file in listFile:
                  #se crea la variable con el nombre del archivo y su extension
                  #excelFile= f'.\\Clientes\\{file}.xls'
                  PATH = "C:\\Users\\Study\\Documents\\Universidad\\Grade proyect\\Robot-CONSULTA-DE-PROCESOS-NACIONAL-UNIFICADA-main\\Robot-CONSULTA-DE-PROCESOS-NACIONAL-UNIFICADA-main"
                  excelFile = os.path.join(PATH,"Clientes", f"{file}.xls")                    
                  print(f'\033[1;32m\nEjecutando: {file}\n')
        
                  # extraer los datos del archivo de entrada
                  my_array = pyexcel.get_array(file_name=excelFile, start_row=1)
                  #Creamos la hoja de excel y la mandamos como parametro
                  wout = xlwt.Workbook()
                  #Creamos el archivo de excel
                  excel.crear_xls(wb=wout)
                  #Variable para controlar los errores que habran antes de parar la ejecucion del programa
                  conteoReLoad = 0
                  #Variable que enumera por qué numero de juzgado se encuentra actualmente
                  numerojuzgado = 0
                  
                  #Bucle que recorre cada linea de la hoja de excel i[0] = Cliente , i[1] = Radicado
                  for i in my_array:
                    #Condicion que verifica si ya no hay mas radicados que consultar y acaba la ejecución
                    if i[0] == "": break
                    #Se suma la variable que verifica en que numero de juzgado se encuentra actualmente  
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
                          #Si se genera una excepcion se suma la variable para parar la ejecución
                          conteoReLoad += 1
                          print("\033[1;33m*"*40)
                          print(f"\033[1;31mERROR: {i[1]} #{numerojuzgado}")
                          print("\033[1;33m*"*40)
                          #Se agrega al excel el proceso en el que sucedió la excepción
                          errores["cliente"] = i[0]
                          errores["radicado"] = i[1]
                          
                          #Al llegar a la cantidad de excepciones aceptadas se genera un proceso para recargar la pagina hasta que esta funcione
                          if conteoReLoad == 1:
                              #Se vuelve a instanciar la variable a 0
                              conteoReLoad = 0
                              resp = input("\033[1;36m¿Volver a recargar?\n1.si\n2.no\nrespuesta \2: ")
                              os.system ("cls")
                              #Mientras el usuario pida recargar la pagina se ejecutará dicha instruccion
                              while resp == "1":
                                  print(f'\033[1;32m\nEjecutando: {file}\n')
                                  #Se recargará la pagina
                                  browser.driver.refresh()
                                  resp = input("\033[1;36m¿Volver a recargar?\n1.si\n2.no\nrespuesta \2: ")
                                  os.system ("cls")
                                  print(f'\033[1;32m\nEjecutando: {file}\n')
                              WebDriverWait(browser.driver,timeout=4).until(EC.presence_of_element_located((By.XPATH,'//*[@id="mainContent"]/div/div/div/div[1]/div/div[2]/span')))
                          #Se escribia en la hoja de excel los errores encontrados
                          excel.escribir_xls(errores=errores)
                  #Se creara el excel final con todos los procesos encontrados     
                  excel.terminar(file)
                  print("\n","\033[1;36m*"*20,"\n","\033[1;36mFinalizó ",file,"\n","\033[1;36m*"*20,"\n")
        #Se Termina el programa 
        print("\n","\033[1;36m*"*20,"\n","\033[1;36mFinalizó la ejecución","\n","\033[1;36m*"*20,"\n")
        #Se cerrará el driver y finalizará la ejecución
        browser.close()
    else:
            print("\033[1;36mUsar\n1.Historico = Hacer historico de los procesos\n2.Recientes = Para obtener los ultimos movimientos publicados")
            msvcrt.getch()
            sys.exit(1)
    print("\033[1;36m\n----------------------\n")  
    print("\033[1;36mPresiona una tecla para cerrar")
    msvcrt.getch()
    sys.exit(1)                    

    
        
        
        

        
