import csv
import pandas as pd
import pyodbc
import pymssql
from meza import io
import os



def createCSV():
    try:
        print(" + Read & Create CSV Files")
        if os.path.exists("Empleado.csv"):
            os.remove("Empleado.csv")

        if os.path.exists("Evento.csv"):
            os.remove("Evento.csv")

        if os.path.exists("EmpleadoHorario.csv"):
            os.remove("EmpleadoHorario.csv")

        for i in [1,2,3]:
            #EMPLEADO
            if i == 1:            
                # connEm = pyodbc.connect(r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};" + r"Dbq=R:\datos.mdb;")
                connEm = pyodbc.connect(r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};" + r"Dbq=\\192.168.0.250\TimeWork\datos.mdb;")
                cursEm = connEm.cursor()
                SQL = 'SELECT ID, Numero, Nombre, Apellidos, Clave FROM Empleado;' # insert your query here
                cursEm.execute(SQL)
                rows = cursEm.fetchall()
                cursEm.close()
                connEm.close()
                csv_writerEm = csv.writer(open('Empleado.csv', 'a'), lineterminator='\n')
                for row in rows:
                    csv_writerEm.writerow(row)
                print("     - ✔️File Empleado.csv created")

            #EVENTO
            if i == 2:
                connEvento = pyodbc.connect(r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};" + r"Dbq=\\192.168.0.250\TimeWork\datos.mdb;")
                cursEvento = connEvento.cursor()
                SQL = 'SELECT ID, IDEmpleado, Entrada, Salida, TeMinimo, IDSucursalFuente, IDRowEnSucursalFuente FROM Evento;' # insert your query here
                cursEvento.execute(SQL)
                rows = cursEvento.fetchall()
                cursEvento.close()
                connEvento.close()
                csv_writerEvento = csv.writer(open('Evento.csv', 'a'), lineterminator='\n')
                for row in rows:
                    csv_writerEvento.writerow(row)
                print("     - ✔️File Evento.csv created")

            #EMPLEADOHORARIO
            if i == 3:
                connEmpleadoH = pyodbc.connect(r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};" + r"Dbq=\\192.168.0.250\TimeWork\datos.mdb;")
                cursEmpleadoH = connEmpleadoH.cursor()
                SQLE = 'SELECT * FROM EmpleadoHorario;' # insert your query here
                cursEmpleadoH.execute(SQLE)
                rows = cursEmpleadoH.fetchall()
                cursEmpleadoH.close()
                connEmpleadoH.close()
                csv_writerEmpleadoHorario = csv.writer(open('EmpleadoHorario.csv', 'a'), lineterminator='\n')
                for row in rows:
                    csv_writerEmpleadoHorario.writerow(row)
                print("     - ✔️File EmpleadoHorario.csv created")
    except:
        print(" ⚠️ Something is wrong, TimeChecker is Open ⚠️")

def deleteTables():
    print(" + Delete data from Intranet Tables")
    for i in [1,2,3]:
        #Evento
        if i == 1:
            connectionEv = pymssql.connect(server='192.168.0.206', user='sa', password='Sql@dmin1', database='INTRANET')
            deleteEvento = "DELETE FROM A_Evento"
            cursorEvento = connectionEv.cursor()
            cursorEvento.execute(deleteEvento)
            connectionEv.commit()
            print("     - ✔️Data from Evento removed")
        #Empleado
        if i == 2:
            connectionEm = pymssql.connect(server='192.168.0.206', user='sa', password='Sql@dmin1', database='INTRANET')
            deleteEmpleado = "DELETE FROM A_Empleado"
            cursorEmpleado = connectionEm.cursor()
            cursorEmpleado.execute(deleteEmpleado)
            connectionEm.commit()            
            print("     - ✔️Data from Empleado removed")
        #EmpleadoHorario
        if i == 3:
            connectionEH = pymssql.connect(server='192.168.0.206', user='sa', password='Sql@dmin1', database='INTRANET')    
            deleteEH = "DELETE FROM H_EmpleadoHorario"        
            cursorEH = connectionEH.cursor()
            cursorEH.execute(deleteEH)
            connectionEH.commit()
            print("     - ✔️Data from EmpleadoHorario removed")

def writeDatabase():
    print(" + Save Data in Intranet DataBase")
    for i in [1,2,3]:
        #Empleado
        if i ==1:                        
            connectionEm = pymssql.connect(server='192.168.0.206', user='sa', password='Sql@dmin1', database='INTRANET')    

            with open('Empleado.csv', 'r') as fileEv:
                readerEm = csv.reader(fileEv)
                columnsEm = next(readerEm)    
                # queryEm = 'INSERT INTO H_Empleado values (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)'
                queryEm = 'INSERT INTO A_Empleado values (%s, %s, %s, %s, %s)'
                cursorEm = connectionEm.cursor()

                for data in readerEm:
                    cursorEm.execute(queryEm, tuple(data))
                    connectionEm.commit()
            print("     - ✔️Data saved in Empleado ")            
        #Evento
        if i == 2:            
            connectionEventos = pymssql.connect(server='192.168.0.206', user='sa', password='Sql@dmin1', database='INTRANET')    

            with open('Evento.csv', 'r') as fileEv:
                readerEv = csv.reader(fileEv)
                columnsEv = next(readerEv)    
                queryEv = 'INSERT INTO A_Evento values (%s,%s,%s,%s,%s,%s,%s)'
                cursorEv = connectionEventos.cursor()

                for data in readerEv:
                    cursorEv.execute(queryEv, tuple(data))
                    connectionEventos.commit()
            print("     - ✔️Data saved in Evento")
        #EmpladoHorario
        if i == 3:
            connectionEH = pymssql.connect(server='192.168.0.206', user='sa', password='Sql@dmin1', database='INTRANET')    

            with open('EmpleadoHorario.csv', 'r') as fileEH:
                readerEH = csv.reader(fileEH)
                columnsEH = next(readerEH)
                queryEH = 'INSERT INTO H_EmpleadoHorario values (%s,%s,%s,%s)'
                # query = query.format(','.join(columns), ','.join(',' * len(columns)))
                cursorEH = connectionEH.cursor()

                for data in readerEH:
                    cursorEH.execute(queryEH, tuple(data))
                    connectionEH.commit()
            print("     - ✔️Data saved in EmpleadoHorario")
    

def main():
    createCSV()
    deleteTables()
    writeDatabase()
    

if __name__ == "__main__":
    print("**** Start ****")
    main()
    print("**** End ****")


