import csv
import pandas as pd
import pyodbc
import pymssql
from meza import io
import os



def createCSV():

    if os.path.exists("Empleado.csv"):
        os.remove("Empleado.csv")

    if os.path.exists("Evento.csv"):
        os.remove("Evento.csv")

    if os.path.exists("EmpleadoHorario.csv"):
        os.remove("EmpleadoHorario.csv")

    for i in [1,2,3]:
        #EMPLEADO
        if i == 1:
            print("No funciona lectura de empleado")
            # connEmpleado = pyodbc.connect(r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};" + r"Dbq=R:\datos.mdb;")
            # connEmpleado.setdecoding(pyodbc.SQL_CHAR, encoding='utf8')
            # connEmpleado.setdecoding(pyodbc.SQL_WCHAR, encoding='utf8')
            # connEmpleado.setencoding(encoding='utf8')

            # cursEmopleado = connEmpleado.cursor()
            # SQL = 'SELECT * FROM Empleado;' # insert your query here
            # cursEmopleado.execute(SQL)
            # rows = cursEmopleado.fetchall()
            # cursEmopleado.close()
            # connEmpleado.close()
            # csv_writerEmpleado = csv.writer(open('Empleado.csv', 'a'), lineterminator='\n')
            # for row in rows:
            #     csv_writerEmpleado.writerow(row)
                # csv_writerEmpleado.writerow(repr(row))

        #EVENTO
        if i == 2:
            connEvento = pyodbc.connect(r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};" + r"Dbq=R:\datos.mdb;")
            cursEvento = connEvento.cursor()
            SQL = 'SELECT * FROM Evento;' # insert your query here
            cursEvento.execute(SQL)
            rows = cursEvento.fetchall()
            cursEvento.close()
            connEvento.close()
            csv_writerEvento = csv.writer(open('Evento.csv', 'a'), lineterminator='\n')
            for row in rows:
                csv_writerEvento.writerow(row)

        #EMPLEADOHORARIO
        if i == 3:
            connEmpleadoH = pyodbc.connect(r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};" + r"Dbq=R:\datos.mdb;")
            cursEmpleadoH = connEmpleadoH.cursor()
            SQLE = 'SELECT * FROM EmpleadoHorario;' # insert your query here
            cursEmpleadoH.execute(SQLE)
            rows = cursEmpleadoH.fetchall()
            cursEmpleadoH.close()
            connEmpleadoH.close()
            csv_writerEmpleadoHorario = csv.writer(open('EmpleadoHorario.csv', 'a'), lineterminator='\n')
            for row in rows:
                csv_writerEmpleadoHorario.writerow(row)

def deleteTables():

    for i in [1,2,3]:
        if i == 1:
            connectionEv = pymssql.connect(server='192.168.0.206', user='sa', password='Sql@dmin1', database='INTRANET')
            deleteEvento = "DELETE FROM H_Evento"
            cursorEvento = connectionEv.cursor()
            cursorEvento.execute(deleteEvento)
            connectionEv.commit()
            print(" - Datos de Evento eliminado")
        if i == 2:
            connectionEm = pymssql.connect(server='192.168.0.206', user='sa', password='Sql@dmin1', database='INTRANET')
            deleteEmpleado = "DELETE FROM H_Empleado"
            cursorEmpleado = connectionEm.cursor()
            cursorEmpleado.execute(deleteEmpleado)
            connectionEm.commit()
            
            print(" - Datos de Empleado eliminado")
        if i == 3:
            connectionEH = pymssql.connect(server='192.168.0.206', user='sa', password='Sql@dmin1', database='INTRANET')    
            deleteEH = "DELETE FROM H_EmpleadoHorario"        
            cursorEH = connectionEH.cursor()
            cursorEH.execute(deleteEH)
            connectionEH.commit()
            print(" - Datos de EmpleadoHorario eliminado")

def writeDatabase():
    for i in [1,2,3]:
        if i == 1:            
            print(" - Informacion guardada en la tabla Empleado")

        if i == 2:
            # connectionEventos = pymssql.connect(server='192.168.0.206', user='sa', password='Sql@dmin1', database='INTRANET')    

            # with open('Evento.csv', 'r') as fileEv:
            #     readerEv = csv.reader(fileEv)
            #     columnsEv = next(readerEv)    
            #     queryEv = 'INSERT INTO H_Evento values (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)'
            #     cursorEv = connectionEventos.cursor()

            #     for data in readerEv:
            #         cursorEv.execute(queryEv, tuple(columnsEv))
            #         connectionEventos.commit()
            print(" - Informacion guardada en la tabla Evento")

        if i == 3:
            connectionEH = pymssql.connect(server='192.168.0.206', user='sa', password='Sql@dmin1', database='INTRANET')    

            with open('EmpleadoHorario.csv', 'r') as fileEH:
                readerEH = csv.reader(fileEH)
                columnsEH = next(readerEH)
                queryEH = 'INSERT INTO H_EmpleadoHorario values (%s,%s,%s,%s)'
                # query = query.format(','.join(columns), ','.join(',' * len(columns)))
                cursorEH = connectionEH.cursor()

                for data in readerEH:
                    cursorEH.execute(queryEH, tuple(columnsEH))
                    connectionEH.commit()
            print(" - Informacion guardada en la tabla EmpleadoHorario")
    

def main():
    createCSV()
    deleteTables()
    writeDatabase()
    

if __name__ == "__main__":
    print("***Iniciando Rutina para creacion de reporte de TimeChecker")
    main()
    print("***Rutina Finalizada")


