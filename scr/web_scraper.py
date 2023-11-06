import time
import requests
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException, UnexpectedAlertPresentException
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.chrome.service import Service
from twocaptcha import TwoCaptcha, NetworkException
from PIL import Image
from io import BytesIO
import pyautogui
import os
import glob
import shutil
import datetime
import random
import subprocess
import string
import re
#para enviar email
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

class descargaReportes():
    def __init__(self):
        self.directoryPath = os.getcwd()
        self.defaultPathDownloads = self.directoryPath + r'\temp'
        self.options = webdriver.ChromeOptions()
        self.options.add_experimental_option("prefs", {
            "download.default_directory": self.defaultPathDownloads,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        })

        # Para Ignorar los errores de certificado SSL (La conexion no es privada)
        self.options.add_argument("--ignore-certificate-errors")
        #self.pathDriver = "driver/chromedriver.exe"
        self.url = "https://multiagente.flexinumber.com/space/58868/settings/data_export"
        #self.service = Service(self.pathDriver)
        self.driver = webdriver.Chrome(options=self.options)
        self.driver.maximize_window()

    def reiniciar(self):
        self.__init__()
    
    def login(self):
        self.driver.get(self.url)
        time.sleep(1)

    def iniciarSesion(self, username, password):
        iniciarSesion = self.driver.find_element(By.CSS_SELECTOR, '[id="input-30"]')
        iniciarSesion.send_keys(username)
        time.sleep(1)

        ingresarPass = self.driver.find_element(By.CSS_SELECTOR, '[id="input-35"]')
        ingresarPass.send_keys(password)
        time.sleep(1)
        
        sigIn = self.driver.find_element(By.CSS_SELECTOR, '[type="submit"]')
        sigIn.click()
        time.sleep(5)
    
    def validaInicioSesion(self):
        wait = WebDriverWait(self.driver, 60)
        ajustes = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '[data-pw="ajustes"]')))
        if ajustes:
            return True
        else:
            return False
    
    def cerrarSesion(self):
        logoImagen = self.driver.find_element(By.CSS_SELECTOR, '[alt="User Avatar"]')
        logoImagen.click()
        time.sleep(1)

        cerrarSesion = self.driver.find_element(By.XPATH, "//div[@class='v-list-item__title caption' and text()='Cerrar sesión']")
        cerrarSesion.click()
        time.sleep(1)

    def notificaciones(self):
        script = """
                var elemento = document.querySelector('.v-badge__badge.orange');
                var texto = elemento.textContent.trim();
                return texto;
                """
        total_notificaciones = self.driver.execute_script(script)
        return total_notificaciones

    def cantidadCSV(self):
        ruta_carpeta = self.defaultPathDownloads
        extension = '*.csv'
        patron_busqueda = os.path.join(ruta_carpeta, extension)
        archivos = glob.glob(patron_busqueda)
        cantidad_archivos = len(archivos)
        return cantidad_archivos

    def reporte_contacto(self):
        ajustes = self.driver.find_element(By.CSS_SELECTOR, '[data-pw="ajustes"]')
        ajustes.click()
        time.sleep(1)

        exportacion_de_datos = self.driver.find_element(By.XPATH, "//div[@class='v-list-item__title ps-3' and contains(text(), 'Exportación de datos')]")
        exportacion_de_datos.click()
        time.sleep(1)

        lista_reportes = self.driver.find_element(By.CSS_SELECTOR, '[class="v-select__selections"]')
        lista_reportes.click()
        time.sleep(1)

        contactos = self.driver.find_element(By.XPATH, "//div[@class='v-list-item__title' and text()='Contactos']")
        contactos.click()
        time.sleep(1)

        cantidad_notificacion_inicial = self.notificaciones()

        boton_exportar = self.driver.find_element(By.CSS_SELECTOR, '[data-pw="btn-exp-data"]')
        boton_exportar.click()
        time.sleep(1)

        cantidad_notificacion_final = cantidad_notificacion_inicial
        while cantidad_notificacion_final == cantidad_notificacion_inicial:
            time.sleep(1)
            cantidad_notificacion_final = self.notificaciones()
        else:
            pass          

        boton_actualizar = self.driver.find_element(By.XPATH, "//div[contains(text(), 'Actualizar')]")
        boton_actualizar.click()
        time.sleep(3)

        cantidad_csv_inicial = self.cantidadCSV()

        descargar_archivo = self.driver.find_element(By.XPATH, "(//div[@class='v-data-table__wrapper']//tbody/tr[1]//a)[1]")
        descargar_archivo.click()

        #Valida que la descarga concluya
        cantidad_csv_final = cantidad_csv_inicial
        while cantidad_csv_final == cantidad_csv_inicial:
            time.sleep(1)
            cantidad_csv_final = self.cantidadCSV()
        else:
            pass
        time.sleep(3)
    
    def reporte_mensajes(self, finicio, ffin):
        ajustes = self.driver.find_element(By.CSS_SELECTOR, '[data-pw="ajustes"]')
        ajustes.click()
        time.sleep(1)

        exportacion_de_datos = self.driver.find_element(By.XPATH, "//div[@class='v-list-item__title ps-3' and contains(text(), 'Exportación de datos')]")
        exportacion_de_datos.click()
        time.sleep(1)

        lista_reportes = self.driver.find_element(By.CSS_SELECTOR, '[class="v-select__selections"]')
        lista_reportes.click()
        time.sleep(1)

        mensajes = self.driver.find_element(By.XPATH, "//div[@class='v-list-item__title' and text()='Mensajes']")
        mensajes.click()
        time.sleep(1)

        wait = WebDriverWait(self.driver, 30)
        intervalo_fechas = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '[class="v-text-field__slot"]')))
        intervalo_fechas.click()
        time.sleep(1)

        calendario = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '[class="v-menu__content theme--light menuable__content__active"]')))
        mes_anterior = self.driver.find_element(By.XPATH, '//button[@aria-label="Mes anterior"]')
        proximo_mes = self.driver.find_element(By.XPATH, '//button[@aria-label="Próximo mes"]')

        meses ={'enero': '01',
                'febrero': '02',
                'marzo': '03',
                'abril': '04',
                'mayo': '05',
                'junio': '06',
                'julio': '07',
                'agosto': '08',
                'septiembre': '09',
                'octubre': '10',
                'noviembre': '11',
                'diciembre': '12'
                }
        
        def fecha_default():
            anio_mes = self.driver.find_element(By.XPATH, "//div[@class='accent--text']/button")
            anio_mes_txt = anio_mes.text

            anio_default = re.findall(r'\d+', anio_mes_txt)
            anio_numero_default = anio_default[0]

            palabras = anio_mes_txt.split()
            mes_default = palabras[0]
            mes_numero_default = meses[mes_default]

            periodo = int(anio_numero_default + mes_numero_default)

            return periodo
        
        # === fecha inicio ====
        fecha_inicio = datetime.datetime.strptime(finicio, '%Y-%m-%d')
        periodo_inicio = int(datetime.datetime.strftime(fecha_inicio, '%Y%m'))
        dia_inicio = fecha_inicio.day

        periodo_sistema = fecha_default()

        if periodo_inicio == periodo_sistema:
            pass
        elif periodo_inicio < periodo_sistema:
            while periodo_inicio < periodo_sistema:
                mes_anterior.click()
                time.sleep(1)
                periodo_sistema = fecha_default()
        else:
            while periodo_inicio > periodo_sistema:
                proximo_mes.click()
                time.sleep(1)
                periodo_sistema = fecha_default()
        
        fecha_inicial_seleccionar = self.driver.find_element(By.XPATH,f'//td/button[@type="button"]/div[text()="{dia_inicio}"]')
        fecha_inicial_seleccionar.click()
        time.sleep(1)

        # === fecha fin ===
        fecha_final = datetime.datetime.strptime(ffin, '%Y-%m-%d')
        periodo_final = int(datetime.datetime.strftime(fecha_final, '%Y%m'))
        dia_final = fecha_final.day

        periodo_sistema = fecha_default()

        if periodo_final == periodo_sistema:
            pass
        elif periodo_final < periodo_sistema:
            while periodo_inicio < periodo_sistema:
                mes_anterior.click()
                time.sleep(1)
                periodo_sistema = fecha_default()
        else:
            while periodo_final > periodo_sistema:
                proximo_mes.click()
                time.sleep(1)
                periodo_sistema = fecha_default()
        
        fecha_final_seleccionar = self.driver.find_element(By.XPATH,f'//td/button[@type="button"]/div[text()="{dia_final}"]')
        fecha_final_seleccionar.click()
        time.sleep(1)

        cerrar_calendario = self.driver.find_element(By.XPATH, '//button[@data-pw="btn-close"]')
        cerrar_calendario.click()
        time.sleep(1)

        cantidad_notificacion_inicial = self.notificaciones()

        boton_exportar = self.driver.find_element(By.CSS_SELECTOR, '[data-pw="btn-exp-data"]')
        boton_exportar.click()
        time.sleep(1)

        cantidad_notificacion_final = cantidad_notificacion_inicial
        while cantidad_notificacion_final == cantidad_notificacion_inicial:
            time.sleep(1)
            cantidad_notificacion_final = self.notificaciones()
        else:
            pass          

        boton_actualizar = self.driver.find_element(By.XPATH, "//div[contains(text(), 'Actualizar')]")
        boton_actualizar.click()
        time.sleep(3)

        cantidad_csv_inicial = self.cantidadCSV()

        descargar_archivo = self.driver.find_element(By.XPATH, "(//div[@class='v-data-table__wrapper']//tbody/tr[1]//a)[1]")
        descargar_archivo.click()

        #Valida que la descarga concluya
        cantidad_csv_final = cantidad_csv_inicial
        while cantidad_csv_final == cantidad_csv_inicial:
            time.sleep(1)
            cantidad_csv_final = self.cantidadCSV()
        else:
            pass
        time.sleep(3)

    # Funcion que reubicará las descargas en sus respectivas carpetas
    def renombrarReubicar(self, nuevoNombre, carpetaDestino):
        # ruta_descargas = r"C:\Users\Usuario\Documents\terceriza\Robot\descargasPython\descargaRobotin"
        ruta_descargas = self.directoryPath + r'/temp'
        archivos_descargados = sorted(
            glob.glob(os.path.join(ruta_descargas, '*')), key=os.path.getmtime, reverse=True
        )
        # Comprobar si hay archivos descargados
        if len(archivos_descargados) > 0:
            ultimo_archivo = archivos_descargados[0]
            # Cambiar el nombre del archivo --1er argumento de la funcion
            nuevo_nombre = f'{nuevoNombre}.csv'
            carpeta_destino = carpetaDestino
            # Comprobar si la carpeta de destino existe, si no, crearla
            if not os.path.exists(carpeta_destino):
                os.makedirs(carpeta_destino)
            # Ruta completa del archivo de destino
            ruta_destino = os.path.join(carpeta_destino, nuevo_nombre)
            # Mover el archivo a la carpeta de destino con el nuevo nombre
            shutil.move(ultimo_archivo, ruta_destino)

    # Funcion que crea el nombre del reporte
    def nombreReporte(self, name, finicio, ffin, fechaD0 = True):
        if fechaD0:
            fechaHora = datetime.datetime.now()
            fecha = fechaHora.strftime("%Y%m%d_%H%M%S")
            aleatorio = str(random.randint(100, 999))
            nameFile = name + fecha + '_' + aleatorio
        else:
            if ffin == None:
                ffin = finicio
            else:
                pass
            h = datetime.datetime.now()
            hora = h.strftime('%H%M%S')
            fechan = datetime.datetime.strptime(ffin, '%Y-%m-%d')
            fechan = fechan + datetime.timedelta(days=1)
            fecha = fechan.strftime("%Y%m%d_")
            aleatorio = str(random.randint(100, 999))
            nameFile = name + fecha + hora + '_' + aleatorio
        
        return nameFile
        
    def gameOver(self):
        self.driver.quit()


