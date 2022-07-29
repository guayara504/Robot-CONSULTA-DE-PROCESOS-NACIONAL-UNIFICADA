from Consultor import *
from Driver import *
from Extractor import *
from Excel import Excel_1
import sys
import msvcrt

#Clase principal   
if __name__ == "__main__":
    #Escoger los archivos con los que se trabajara
    inputFile=input("\n\nIngrese archivo(s) de Excel separado por comas: ")
    listFile=inputFile.split(",")
    inicioBusqueda=input("\n\n1.Historico\n2.Recientes\nDonde comenzará la busqueda en la pagina: ")
    
    #Se llaman las clases a trabajar
    browser =Driver_1()
    consulta =Consultor_1
    extractor = Extractor_1
    excel = Excel_1()
    
    if (inicioBusqueda == "1") or (inicioBusqueda == "2"):
        
          #Hacer esto con todos los archivos ingresados
        for file in listFile:
                  excelFile= file+".xlsx"
        
                  print('Ejecutando ...'+file)
        
                  # extraer los datos del archivo de entrada
                  my_array = pyexcel.get_array(file_name=excelFile, start_row=1)
                  #Creamos la hoja de excel y la mandamos como parametro
                  wout = xlwt.Workbook()
                  excel.crear_xls(wb=wout)
        
                  for i in my_array:
                      if i[0] == "": break
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
                          for dato in datos:
                            print(dato+":",datos[dato])
                          #for act in actuaciones:
                           # print(act)
                          #Se ingresa la informacion al excel
                          excel.escribir_xls(datos=datos,actuaciones=actuaciones,i=i)
                      except:
                          print("\n","*"*20,"\n","-ERROR-"*10,"\n","*"*20,"\n")
                          errores["cliente"] = i[0]
                          errores["radicado"] = i[1]
                          excel.escribir_xls(errores=errores)
                  excel.terminar(file)
                  print("\n","*"*20,"\n","Finalizo ",file,"\n","*"*20,"\n")
        #Se Termina el programa 
        print("\n","*"*20,"\n","Se Finalizo la ejecucion","\n","*"*20,"\n")
        browser.close()
    else:
            print("Usar\n1.Historico = Hacer historico de los procesos\n2.Recientes = Para obtener los ultimos movimientos publicados")
            msvcrt.getch()
            sys.exit(1)
    print("\n----------------------\n")  
    print("Presiona una tecla para cerrar")
    msvcrt.getch()
    sys.exit(1)                    

    
        
        
        

        
