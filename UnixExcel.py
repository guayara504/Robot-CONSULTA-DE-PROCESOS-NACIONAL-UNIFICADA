from Consultor import *
from Driver import *
from Extractor import *
from Excel import Excel_1

#Clase principal   
if __name__ == "__main__":
    #Escoger los archivos con los que se trabajara
    inputFile=input("\n\nIngrese archivo(s) de Excel separado por comas: ")
    listFile=inputFile.split(",")
    
    #Se llaman las clases a trabajar
    browser =Driver_1()
    consulta =Consultor_1
    extractor = Extractor_1
    excel = Excel_1()
    
    demora = 3
    def WaitForElement(driver,path):
        limit = demora
        inc = 1
        c = 0
        while c < limit:
            try:
                driver.find_element(By.XPATH,path)
                return 1
            except:
                time.sleep(inc)
                c+=inc
        return 0
    
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
                  #Lista donde guardara las actuaciones del proceso
                  actuaciones = []
                  #Diccionario para guardar los datos del proceso
                  datos = {"cliente":"","radicado":"","fecha": "","despacho":"","ponente":"","tipo_proceso":"","clase_proceso":"","recurso":"","ubicacion":"","contenido":"","demandante":"","demandado":""}
                  #Damos click en boton "TODOS LOS PROCESOS"
                  consulta.iniciar_busqueda((browser.driver))
                  #Se ingresa el radicado a consultar
                  consulta.ingresar_radicado((browser.driver),i[1])
                  #Se selecciona el proceso a consultar
                  consulta.click_Proceso((browser.driver),WaitForElement)
                  #Se extrae los datos del proceso
                  extractor.extraer_datos(datos,(browser.driver),i)
                  #Se extrae las partes interesadas en el proceso
                  extractor.extraer_partes(datos,(browser.driver),WaitForElement)
                  #Se extrae las actuaciones del proceso
                  extractor.extraer_actuaciones(actuaciones,(browser.driver),i[1],WaitForElement)
                  #Se imprime en pantalla los datos
                  for dato in datos:
                    print(dato+":",datos[dato])
                  #for act in actuaciones:
                   # print(act)
                  #Se ingresa la informacion al excel
                  excel.escribir_xls(datos, actuaciones,i)
                  

    
        
        
        

        
