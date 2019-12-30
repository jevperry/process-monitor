import win32com.client
import wmi
import sys
import MySQLdb
import datetime
import socket
import threading
import thread
import pythoncom
import time

class monitor():
    con = None
    c = wmi.WMI ()
    global ip
    global chromeCont
    chromeCont = 0
    int(chromeCont)
    ip = socket.gethostbyname(socket.gethostname())

    def monitoreo(strExcluidos,tipo):
        #Aqui se define el Hilo a monitorear los procesos Creados y Borrados, se usa el mismo hilo con diferente "tipo"
        #siguiente linea se utiliza para comunicacion desde el hilo hacia afuera
        pythoncom.CoInitializeEx(10)
        #se define nuevo como una bandera para controlar el flujo del hilo intentando que su iteracion consuma menos carga
        nuevo = 0
        #tipo 1 para Creacion/tipo 2 para Borrado
        if(tipo==1):
            process_watcher = wmi.WMI().Win32_Process.watch_for("creation")
        else:
            process_watcher = wmi.WMI().Win32_Process.watch_for("deletion")
        while True:
            #Ocurre un nuevo proceso
            nuevo = process_watcher()
            if(nuevo != 0):
                #print(nuevo.Caption) #impresion para revisar/se captura la fecha/hora del nuevo proceso (borrado o creado)
                now = datetime.datetime.now()
                #se redefine el contador de chrome como global para comunicarlo fuera del hilo
                global chromeCont
                if(nuevo.Caption in strExcluidos):
                        #print("Proceso Excluido de Insercion") #para revisar/si se encuentra el proceso en los excluidos lo ignora
                        continue
                if((nuevo.Caption == "chrome.exe") & (tipo == 1)):
                        chromeCont = chromeCont+1
                        #print(chromeCont) #para revisar
                        if(chromeCont == 1):
                                pass
                        else:
                                continue
                elif((tipo == 2) & (nuevo.Caption == "chrome.exe")):
                        #print("chromeCont : "+str(chromeCont))
                        chromeCont = chromeCont-1
                        #print(chromeCont)
                        if(chromeCont == 0):
                                pass
                        else:
                                continue
                #ya que el proceso no se encontraba ni en excluidos y paso el filtro de chrome, se crea la conexion con la base de datos y ejecutan los cambios
                con2 = MySQLdb.connect(host="152.74.180.10",user="procesos2",passwd="proc",db="labomatic")
                cursor = con2.cursor()
                
                if(tipo == 1):
                        insert = "insert into procesos(nombre_proceso,pid,fecha_hora_inicio,ip,ruta)values('"+nuevo.Caption+"',"+str(nuevo.ProcessId)+",'"+now.strftime('%Y-%m-%d %H:%M:%S')+"','"+str(ip)+"','"+str(nuevo.ExecutablePath)+"');commit;"
                        #print(insert)
                        cursor.execute(insert)
                elif(tipo == 2):
                        update = "update procesos set fecha_hora_fin ='"+now.strftime('%Y-%m-%d %H:%M:%S')+"' where ip='"+ip+"' and fecha_hora_fin is null and nombre_proceso='"+nuevo.Caption+"';commit;"
                        #print(update)
                        cursor.execute(update)
                #ya terminada la ejecucion en la base de datos se desconecta y limpian variables
                cursor.close()
                con2.close()
                con2 = None
                nuevo = 0
            
            

            
    #Conexion a la base de datos Procesos e inicializacion del cursor a usar para la misma
    con = MySQLdb.connect(host="152.74.180.10",user="procesos2",passwd="proc",db="labomatic")
    cursor = con.cursor()
    #Obtenemos la fecha/hora de ejecucion para ingresar en la base de datos
    now = datetime.datetime.now()
    #Actualizacion de la base de datos para cerrar aquellos procesos sin fecha/hora fin previos a la nueva sesion
    update_sin_cerrar = "UPDATE procesos SET fecha_hora_fin='"+now.strftime('%Y-%m-%d %H:%M:%S')+"',comentario='mal cerrado' where fecha_hora_fin is null and ip='"+ip+"';commit;"
    #print(update_sin_cerrar) #solo para revisar
    cursor.execute(update_sin_cerrar)
       
    #Se almacenan localmente los Procesos Excluidos de ser manipulados
    #Llenado de Array con los procesos excluidos
    cursor.close()
    cursor = con.cursor()
    cursor.execute("select nombre from procesos_excluyentes;")
    row = cursor.fetchall()
    q = 1
    Excluidos = []
    for q in range(0,len(row)-1):
        Excluidos.append(str(row[q]))
        #print(Excluidos[q])
    strExcluidos = ''.join(Excluidos)
    con.close()
    cursor.close()
    con = None

    #creacion de los hilos de monitoreo para creacion y borrado de procesos
    d = threading.Thread(target=monitoreo, args=(strExcluidos,1))
    e = threading.Thread(target=monitoreo, args=(strExcluidos,2))
    d.start()
    e.start()
     
if __name__ == "__main__":
    monitor()
