from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains  # Para realizar scroll hasta un elemento
import os
from random import uniform
from bs4 import BeautifulSoup
import re
from difflib import SequenceMatcher
from time import sleep


class DownloadManager():
    """
    Clase para automatizar la descarga de archivos Excel desde el portal de liquidaciones.
    
    La clase utiliza Selenium para interactuar con la página y BeautifulSoup para analizar el HTML.
    Se implementan métodos para navegar por el portal, identificar elementos relevantes (meses, 
    revisiones y archivos) y realizar la descarga del archivo.
    """

    def __init__(self, downloads_path):
        """
        Inicializa el driver de Chrome y configura las opciones de descarga.
        
        Args:
            downloads_path (str): Ruta del directorio donde se guardarán los archivos descargados.
        """
        self.url = 'https://www.coes.org.pe/Portal/mercadomayorista/liquidaciones'
        self.downloads_path = downloads_path
        # Genera un tiempo de espera aleatorio entre 2 y 5 segundos para simular la interacción humana
        self.time_wait = lambda: round(uniform(2, 5), 3)

        # Configuración del navegador
        options = Options()
        # Ejecución en modo incógnito (opcional)
        # options.add_argument("--incognito")
        # Ejecución en modo headless (opcional)
        # options.add_argument('--headless')
        # Deshabilitar aceleración por GPU (opcional)
        # options.add_argument('--disable-gpu')
        
        # Configuración de la ruta de descarga y desactivación del prompt de descarga
        options.add_experimental_option("prefs", {
            "download.default_directory": self.downloads_path,
            "download.prompt_for_download": False,
        })

        # Inicializa el driver y abre la URL del portal
        self.driver = webdriver.Chrome(options=options)
        self.driver.maximize_window()
        self.driver.get(self.url)

    # --------------------- Métodos Generales --------------------- #
    def _scroll_to_element(self, element):
        """
        Realiza scroll hasta el elemento especificado para que sea visible en la pantalla.
        
        Args:
            element (WebElement): Elemento al que se hará scroll.
        """
        actions = ActionChains(self.driver)
        actions.move_to_element(element).perform()

    def _optic_click(self, xpath):
        """
        Realiza clic sobre un elemento identificado por un XPath, esperando que sea clickeable.
        
        Se realizan hasta 5 intentos. Si el elemento no se vuelve clickeable en 10 segundos,
        se muestra un mensaje de error.
        
        Args:
            xpath (str): XPath del elemento a clicar.
        """
        for _ in range(5):
            try:
                # Espera hasta que el elemento sea clickeable (máximo 10 segundos)
                element = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, xpath))
                )
                self._scroll_to_element(element)
                element.click()
                print("Se hizo clic en el elemento:", xpath)
                break
            except TimeoutException:
                print("Timeout: No se pudo clicar el elemento en el tiempo esperado.")
            except Exception as ex:
                print("Se produjo la excepción:", ex)

    def _count_xlsx_files(self):
        """
        Cuenta el número de archivos con extensión .xlsx en el directorio de descargas.
        
        Returns:
            int: Cantidad de archivos .xlsx encontrados.
        """
        archivos = os.listdir(self.downloads_path)
        xlsx_files = [archivo for archivo in archivos if archivo.endswith('.xlsx')]
        return len(xlsx_files)

    def _get_html(self):
        """
        Recupera el HTML actual de la página.
        
        Returns:
            str: Contenido HTML de la página.
        """
        # Se puede usar self.driver.page_source, pero se opta por ejecutar un script para obtener el body.
        body = self.driver.execute_script("return document.body")
        return body.get_attribute('innerHTML')

    # --------------------- Métodos para Navegación y Selección --------------------- #
    def _reference_month(self, mes):
        """
        Retorna el nombre del mes en mayúsculas a partir del número de mes (en formato '01', '02', etc.).
        
        Args:
            mes (str): Número del mes (ejemplo: '01' para enero).
            
        Returns:
            str: Nombre del mes en mayúsculas.
        """
        meses = {
            '01': 'ENERO',
            '02': 'FEBRERO',
            '03': 'MARZO',
            '04': 'ABRIL',
            '05': 'MAYO',
            '06': 'JUNIO',
            '07': 'JULIO',
            '08': 'AGOSTO',
            '09': 'SEPTIEMBRE',
            '10': 'OCTUBRE',
            '11': 'NOVIEMBRE',
            '12': 'DICIEMBRE'
        }
        return meses[mes]

    def _identify_monthname_button(self, mes):
        """
        Identifica el botón correspondiente al mes requerido en la interfaz.
        
        Se compara cada elemento de la lista de meses con la referencia obtenida y se retorna 
        aquel que presente mayor similitud.
        
        Args:
            mes (str): Número del mes (ejemplo: '01').
            
        Returns:
            str: Nombre del mes tal como aparece en la interfaz.
        """
        sleep(self.time_wait())
        html = self._get_html()
        soup = BeautifulSoup(html, 'html.parser')
        
        # Ubica el contenedor con la lista de meses
        container = soup.find('div', id='browserDocument')
        uls = container.find_all('ul')
        month_elements = uls[1].find_all('li')
        
        month_names = []
        similarity_scores = []
        reference = self._reference_month(mes)
        
        for element in month_elements:
            month_str = element.get_text().strip()
            month_names.append(month_str)
            # Se extraen solo las letras para la comparación
            letters_only = ''.join(re.findall('[a-zA-Z]', month_str))
            sim = SequenceMatcher(a=reference, b=letters_only.upper()).ratio()
            similarity_scores.append(round(sim, 3))
        
        # Se selecciona el mes con mayor similitud
        index_max = similarity_scores.index(max(similarity_scores))
        return month_names[index_max]

    def _identify_filenametodownload_button(self):
        """
        Busca y retorna el nombre del archivo "ResumenCuadros" dentro de la tabla de documentos.
        
        Se recorre cada fila de la tabla y se compara la combinación de las dos primeras palabras
        del nombre del archivo (después de reemplazar '_' y '-' por espacios) con la referencia 
        "RESUMENCUADROS". Si la similitud es mayor al 90%, se asume que es el archivo deseado.
        
        Returns:
            str: Nombre del archivo encontrado o 'no_file' si no se identifica.
        """
        sleep(self.time_wait())
        html = self._get_html()
        soup = BeautifulSoup(html, 'html.parser')
        
        reference = 'RESUMENCUADROS'
        table_body = soup.find('table', id='tbDocumentLibrary').find('tbody')
        rows = table_body.find_all('tr')
        
        for row in rows:
            file_name = row.find_all('td')[2].get_text()
            # Reemplaza '_' y '-' por espacios para separar palabras
            pattern = re.compile('[_\\-]')
            modified_name = pattern.sub(' ', file_name)
            words = modified_name.split()
            try:
                # Combina las dos primeras palabras para la comparación
                combined = words[0] + words[1]
                similarity = SequenceMatcher(a=reference, b=combined.upper()).ratio()
                if similarity > 0.90:
                    return file_name
            except IndexError:
                continue  # Si no hay suficientes palabras, se omite la fila
        
        return 'no_file'

    def _click_element_to_download(self, año, mes):
        """
        Navega por la estructura de carpetas del portal para identificar y descargar el archivo 
        correspondiente a la última revisión disponible. Si no se encuentra en las carpetas de revisión, 
        se intenta en la carpeta mensual.
        
        Args:
            año (str/int): Año correspondiente.
            mes (str): Mes requerido.
        """
        sleep(self.time_wait())
        html = self._get_html()
        soup = BeautifulSoup(html, 'html.parser')
        
        # Construir la ruta base utilizada para generar los XPath de navegación
        base_xpath = (
            'Mercado Mayorista/Liquidaciones del MME/01 Mercado de Corto Plazo/'
            'Liquidaciones VTEA/{año}/{mes}/'
        ).format(año=año, mes=mes)
        
        # Se extraen los nombres de las carpetas desde el contenedor principal
        container = soup.find('div', id="browserDocument")
        uls = container.find_all('ul')
        folder_elements = uls[1].find_all('li')
        
        revision_folders = []
        revision_versions = []
        monthly_folder = None
        references = ['REVISION', 'MENSUAL']

        # Clasifica las carpetas según si corresponden a revisión o mensual
        for folder in folder_elements:
            folder_name = folder.get_text().strip()
            # Extrae solo las letras para comparar
            comp_name = ''.join(re.findall(r'[a-zA-Z]', folder_name)).upper()
            sim_revision = SequenceMatcher(a=references[0], b=comp_name).ratio()
            if sim_revision > 0.90:
                revision_folders.append(folder_name)
                # Extrae el número de versión (omitiendo ceros iniciales)
                version_str = ''.join(re.findall(r'(?<!\b0)0*(\d+)', folder_name))
                revision_versions.append(int(version_str))
            else:
                sim_monthly = SequenceMatcher(a=references[1], b=comp_name).ratio()
                if sim_monthly > 0.90:
                    monthly_folder = folder_name

        # Se intenta acceder a la carpeta de revisión con la versión más alta
        while revision_versions:
            max_index = revision_versions.index(max(revision_versions))
            selected_revision = revision_folders[max_index]

            # Se ajusta el nombre de la revisión para el caso específico de agosto 2018
            if año=='2018' and mes == '08_Agosto 2018':
                selected_revision = 'Revisión 01';

            # Construye el XPath para la carpeta de revisión
            xpath_revision = f'//a[@id="{base_xpath}{selected_revision}/"]'
            self._optic_click(xpath_revision)
            
            try:
                file_name = self._identify_filenametodownload_button()
            except Exception:
                file_name = 'no_file'
            
            if file_name == 'no_file':
                # Si no se encontró el archivo, se elimina esta revisión y se vuelve a la carpeta anterior
                xpath_backtracking = f"//a[text()='{mes}']"
                revision_versions.pop(max_index)
                revision_folders.pop(max_index)
                self._optic_click(xpath_backtracking)
            else:
                # Se construye el XPath para el botón de descarga y se hace clic
                xpath_download = f'//*[@id="{base_xpath}{selected_revision}/{file_name}"]'
                self._optic_click(xpath_download)
                break

        # Si no se encontró en ninguna carpeta de revisión, se busca en la carpeta mensual
        if not revision_versions:
            xpath_monthly = f'//a[@id="{base_xpath}{monthly_folder}/"]'
            self._optic_click(xpath_monthly)
            file_name = self._identify_filenametodownload_button()
            xpath_download = f'//*[@id="{base_xpath}{monthly_folder}/{file_name}"]'
            self._optic_click(xpath_download)

    # --------------------- Método Principal --------------------- #
    def download_excel_file(self, año, mes):
        """
        Método principal para la descarga del archivo Excel.
        
        Realiza la siguiente secuencia:
          1. Navega a la sección "Mercado de Corto Plazo".
          2. Ingresa a "Liquidaciones VTEA".
          3. Selecciona el año y mes requeridos.
          4. Realiza la navegación dinámica para identificar y descargar el archivo.
          5. Espera hasta que se detecte el nuevo archivo .xlsx en el directorio de descargas.
          6. Cierra el navegador.
        
        Args:
            año (str/int): Año de la descarga.
            mes (str): Mes de la descarga (en formato numérico, e.g., '01' para enero).
        """
        # Clic en "Mercado de Corto Plazo"
        xpath_corto_plazo = (
            '//*[@id="Mercado Mayorista/Liquidaciones del MME/01 Mercado de Corto Plazo/"]'
        )
        self._optic_click(xpath_corto_plazo)

        # Clic en "Liquidaciones VTEA"
        xpath_liquidaciones = (
            '//*[@id="Mercado Mayorista/Liquidaciones del MME/01 Mercado de Corto Plazo/'
            'Liquidaciones VTEA/"]'
        )
        self._optic_click(xpath_liquidaciones)

        # Selecciona el año requerido
        xpath_año = (
            '//a[@id="Mercado Mayorista/Liquidaciones del MME/01 Mercado de Corto Plazo/'
            'Liquidaciones VTEA/{año}/"]'
        ).format(año=año)
        self._optic_click(xpath_año)

        # Selecciona el mes requerido mediante identificación por similitud
        mes_identificado = self._identify_monthname_button(mes)
        xpath_mes = (
            '//a[@id="Mercado Mayorista/Liquidaciones del MME/01 Mercado de Corto Plazo/'
            'Liquidaciones VTEA/{año}/{mes}/"]'
        ).format(año=año, mes=mes_identificado)
        self._optic_click(xpath_mes)

        # Navegación dinámica para la descarga final del archivo
        self._click_element_to_download(año, mes_identificado)

        # Espera hasta que se detecte que se descargó un nuevo archivo .xlsx
        num_files_before = self._count_xlsx_files()
        WebDriverWait(self.driver, 60).until(lambda d: self._count_xlsx_files() > num_files_before)

        # Cierra el navegador una vez completada la descarga
        self.driver.close()
