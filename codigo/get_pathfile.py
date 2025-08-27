import glob
import os

def identify_last_xlsx_file(downloads_path):
    """
    Identifica el archivo .xlsx más reciente en el directorio indicado,
    basándose en la fecha de creación del archivo.

    Args:
        downloads_path (str): Ruta del directorio donde buscar los archivos .xlsx.

    Returns:
        str: Ruta completa del archivo .xlsx con la fecha de creación más reciente.
             Si no se encuentra ningún archivo, se producirá un ValueError.
    """
    # Construir el patrón de búsqueda usando os.path.join para mayor robustez
    file_pattern = os.path.join(downloads_path, '*.xlsx')
    
    # Obtener la lista de archivos que coinciden con el patrón
    xlsx_files = glob.glob(file_pattern)
    
    # Seleccionar el archivo con la fecha de creación más reciente
    # os.path.getctime obtiene la fecha de creación de cada archivo
    latest_file = max(xlsx_files, key=os.path.getctime)
    
    return latest_file
