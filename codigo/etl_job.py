from download import DownloadManager
from extract import DataExtractor
from transform import DataTransformer
from load import DataLoader
from get_pathfile import identify_last_xlsx_file
import os
from dotenv import load_dotenv
import time

# Cargar variables de entorno desde el archivo .env
ENV_FILE_PATH = 'E:/BI_Comercial/ETL_Python/R006_ETL_ValorizacionEnergia/globals.env'
load_dotenv(ENV_FILE_PATH)

def run_job():
    
    # Obtener fecha a recuoerar 
    start_year = 2018
    start_month = 1
    end_year = 2024
    end_month = 12

    year = start_year
    month = start_month

    while (year < end_year) or (year == end_year and month <= end_month):
        print(f"{year}-{month:02d}")
        # Esperar 3 segundos
        time.sleep(3)
        downloads_path_ve = 'E:\BI_Comercial\ETL_Python\R006_ETL_ValorizacionEnergia\Archivos\Descargas'
        downloads_path = r"{}".format(downloads_path_ve)
        objdownload = DownloadManager(downloads_path)
        objextract = DataExtractor()
        objload = DataLoader()
        objdownload.download_excel_file(f"{year}", f"{month:02d}")

        # Descarga 
        # objdownload.download_excel_file(anio, mes)
        # extraccion de los datos
        filepath = identify_last_xlsx_file(downloads_path)
        data = objextract.extract_data_from_sheet(filepath)
        # transformacion de los datos
        data = DataTransformer.transform_data(data, f"{year}", f"{month:02d}")
        # LLevar a una tabla de base de datos   
        objload.load_data_to_landing(data)

        if month == 12:
            month = 1
            year += 1
        else:
            month += 1
    

def run_etl_job():
    """
    Ejecuta el proceso ETL completo para la valorización de energía.
    
    Este proceso consta de las siguientes etapas:
      1. Descarga del archivo Excel del portal.
      2. Extracción de datos de la hoja de Excel descargada.
      3. Transformación de los datos extraídos (limpieza, renombrado de columnas y adición de campos).
      4. Carga de los datos transformados en la base de datos.
    """
    # Obtener la ruta de descargas definida en las variables de entorno
    downloads_path = os.getenv('DOWNLOADS_PATH')

    # Inicializar los gestores para cada etapa del proceso ETL
    download_manager = DownloadManager(downloads_path)
    db_job = DataLoader()
    data_extractor = DataExtractor()

    # Obtener el periodo a procesar (año y mes) a partir de la base de datos
    year, month = db_job.get_date_to_retrieve()

    # Descargar el archivo Excel correspondiente al periodo indicado
    download_manager.download_excel_file(year, month)

    # Identificar el archivo .xlsx descargado (el más reciente en el directorio de descargas)
    file_path = identify_last_xlsx_file(downloads_path)

    # Extraer los datos de la hoja de Excel
    extracted_data = data_extractor.extract_data_from_sheet(file_path)

    # Transformar los datos extraídos (limpieza y adición de columnas 'Periodo' y 'FechaCreacion')
    transformed_data = DataTransformer.transform_data(extracted_data, year, month)

    # Cargar los datos transformados en la tabla de destino en la base de datos
    db_job.load_data_to_landing(transformed_data)


if __name__ == '__main__':
    run_job()