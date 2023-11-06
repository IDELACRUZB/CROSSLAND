import pandas as pd
import pymysql
import json
from datetime import datetime
import os
import re

class ReportService:
    def loadData(self,filePath, tableName : str, dbName : str, primary_key, properties : {}, renameColumns : {}, converters : []):
        filePath = filePath

        # Nombre de la tabla
        dbTable = tableName

        percentage = {}
        for convert in converters:
            percentage[convert] = self.convertToPercentage
        
        # Lee el archivo EXCEL con pandas
        #df = pd.read_excel( filePath,  sheet_name=0, engine='openpyxl', skiprows=0, dtype=properties, converters=percentage)
        df = pd.read_csv(filePath, sep=',', dtype=str, header=0)

        #df =df.loc[df.iloc[:, 0] != "~Total"]

        # Renombrar la columna        
        df.columns = df.columns.str.strip()
        df.rename(columns=renameColumns, inplace=True)

        # Formatear las columnas
        df = df.fillna(value='')
        df.columns = df.columns.str.lower()
        df.columns = df.columns.str.replace(r'^\d+\.-\s*', '', regex=True)
        df.columns = df.columns.str.replace('\n', '')
        df.columns = df.columns.str.strip()
        df.columns = df.columns.str.replace(' ', '_')
        df.columns = df.columns.str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('utf-8')
        df.columns = df.columns.str.lstrip('_')

        def limpiar_encabezados(encabezado):
            return re.sub(r'\W+', '_', encabezado)
        df.columns = df.columns.map(limpiar_encabezados)

        """
        print(df.columns)
        for i in df.columns:
            print(i)
        exit() #"""
        # Configurar la conexión a la base de datos
        properties = self.getProperties()
        conn = pymysql.connect(
            host=properties['DB_HOST'],
            database= dbName,
            user=properties['DB_USER'],
            password=properties['DB_PASSWORD'],
            port=3306
        )

        try:
            # Crear un cursor y comenzar una transacción
            cur = conn.cursor()
            cur.execute("START TRANSACTION;")
            #cur.execute(f"TRUNCATE TABLE {dbTable};")

            sqlHeading = "`"+"`,`".join(df.columns)+"`"
            columns2 = [col for col in df.columns if col != primary_key]
            sqlHeading2 = ', '.join(f"`{item}` = VALUES(`{item}`)" for item in columns2)
            
            chunks = [df[i:i + 200] for i in range(0, df.shape[0], 200)]

            if primary_key != None:
                for chunk in chunks:
                    values = [tuple(row) for _, row in chunk.iterrows()]
                    try:
                        # Ejecutar el comando INSERT INTO para cada grupo de 50 filas
                        consulta = f"""INSERT INTO {dbTable} ({sqlHeading}) VALUES ({', '.join(['%s'] * len(df.columns))})
                                    ON DUPLICATE KEY UPDATE {sqlHeading2};"""
                        cur.executemany(consulta, values)
                        conn.commit()
                        #print(f"Se insertaron con éxito {len(chunk)} filas")
                    except Exception as e:
                        print("Ocurrió un error:", e)
            else:
                for chunk in chunks:
                    values = [tuple(row) for _, row in chunk.iterrows()]
                    try:
                        # Ejecutar el comando INSERT INTO para cada grupo de 50 filas
                        consulta = f"""INSERT INTO {dbTable} ({sqlHeading}) VALUES ({', '.join(['%s'] * len(df.columns))});"""
                        cur.executemany(consulta, values)
                        conn.commit()
                        #print(f"Se insertaron con éxito {len(chunk)} filas")
                    except Exception as e:
                        print("Ocurrió un error:", e)


                            
            print('Se ejecuto correctamente la consulta: ' + dbName + " / " + tableName)

        except Exception as e:
            # Revertir la transacción si hay un error            
            conn.rollback()    
            print("Hubo un error al importar la informacion: " + str(e) )
            return 400 
            
        finally:
            # Cerrar la conexión a la base de datos
            cur.close()
            conn.close()

        return 200
    
    def getProperties(self):
        config_data = None
        with open('./importador/config.json') as config_file:
            config_data = json.load(config_file)

        return config_data
    
    def convertToPercentage(self,x):
        return "{:.2f}%".format(x * 100)
