from web_scraper import descargaReportes
from isdb import TablaValidacion2
import time
import datetime
import subprocess

# Paso 1: Descarga de Reportes
#Rango de fechas para descarga de Reportes
D0 =  datetime.date.today()
D_1 =  D0 + datetime.timedelta(days=-1)
inicio = str(D_1) #'2023-11-04'
fin = str(D_1) #None#'2023-08-08'

username = 'zoila.cortez@3eriza.pe'
password = 'Crosland2023+'

tablaValidacion = TablaValidacion2()
tablaValidacion.crearBD()
tablaValidacion.crearTabla()
tablaValidacion.truncateTable()

descarga = descargaReportes()
def logueo():
    descarga.login()
    descarga.iniciarSesion(username, password)
    inicioSesion = descarga.validaInicioSesion()

    while not inicioSesion:
        descarga.reiniciar()
        descarga.login()
        descarga.iniciarSesion(username, password)
        inicioSesion = descarga.validaInicioSesion()
    else:
        print('Inicio de Sesion Exitosa')
        pass
        
logueo()

fecD0 = False
contador_descargas = 1
campana = "crossland"

# ===== I. Reporte Contacto =====
def crossland_contacto():
    try:
        descarga.reporte_contacto()
        nombreAsignado = 'crossland_contacto_'
        nombre = descarga.nombreReporte(nombreAsignado, inicio, fin, fecD0)
        destino = descarga.directoryPath + r'/carga\crossland\contacto'
        descarga.renombrarReubicar(nombre, destino)

        datos=[(contador_descargas, campana, nombreAsignado, 1)]
        tablaValidacion.agregarVariosDatos(datos)
    except Exception as e:
        datos=[(contador_descargas, campana, nombreAsignado, 0)]
        tablaValidacion.agregarVariosDatos(datos)
        pass

crossland_contacto()
ultimoRegistro = tablaValidacion.leerDatos()
descargo = ultimoRegistro[0][3]

while descargo == 0:
    tablaValidacion.deleteTable(contador_descargas)

    descarga.reiniciar()
    logueo()

    crossland_contacto()
    ultimoRegistro = tablaValidacion.leerDatos()
    descargo = ultimoRegistro[0][3]
else:
    contador_descargas += 1
    pass

# ===== II. Reporte Mensajes =====

def crossland_mensajes():
    try:
        descarga.reporte_mensajes(inicio, fin)
        nombreAsignado = 'crossland_mensajes_'
        nombre = descarga.nombreReporte(nombreAsignado, inicio, fin, fecD0)
        destino = descarga.directoryPath + r'/carga\crossland\mensajes'
        descarga.renombrarReubicar(nombre, destino)

        datos=[(contador_descargas, campana, nombreAsignado, 1)]
        tablaValidacion.agregarVariosDatos(datos)
    except Exception as e:
        datos=[(contador_descargas, campana, nombreAsignado, 0)]
        tablaValidacion.agregarVariosDatos(datos)
        pass

crossland_mensajes()
ultimoRegistro = tablaValidacion.leerDatos()
descargo = ultimoRegistro[0][3]

while descargo == 0:
    tablaValidacion.deleteTable(contador_descargas)

    descarga.reiniciar()
    logueo()

    crossland_mensajes()
    ultimoRegistro = tablaValidacion.leerDatos()
    descargo = ultimoRegistro[0][3]
else:
    contador_descargas += 1
    pass

descarga.cerrarSesion()
descarga.gameOver()

# Paso 2: Carga la base de datos al servidor
subprocess.call(['python', './importador/controller.py'])