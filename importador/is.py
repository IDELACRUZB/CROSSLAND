import os
from util.email import Email
import json
import glob
from datetime import datetime
from service.ReportService import ReportService

mReport = ReportService()
properties = mReport.getProperties()

# Obtenemos el nombre del archivo
# filePath = r'/home/renovaciones/3ERIZA_RENOVACIONES_ALL.csv'
ROOT_PATH = properties['LOAD_PATH']

#ROOT_PATH = "C:\\Users\\Usuario\\Documents\\terceriza\\Robot\\CROSSLAND_local\\carga\\"
directoryPath = ROOT_PATH  + 'crossland' + "\\" + "contacto"
currentDate = datetime.now().strftime("%Y%m%d")
print(directoryPath)
currentDate = '20231104'

# Patron de coincidencia para los nombres de archivo
patron = f"*_{currentDate}_*"
archivos_coincidentes = glob.glob(os.path.join(directoryPath, patron))
print(archivos_coincidentes)

    
    