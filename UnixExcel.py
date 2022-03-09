from Driver import *
actuaciones = []

#Clase principal   
if __name__ == "__main__":
    datos = {"fecha": "","despacho":"","ponente":"","tipo_proceso":"","clase_proceso":"","demandante":"","demandado":""}
    driver = Driver_1()
    for dato in datos:
        print(dato+":",datos[dato])
