import pandas as pd

class DataExtractor:
    """
    Clase para extraer y delimitar datos de la hoja "CUADRO 4" de un archivo Excel.
    
    Esta clase proporciona métodos para identificar la fila de cabecera, delimitar el
    área de datos y aplicar filtros sobre los datos extraídos.
    """
    
    def __init__(self):
        # Constructor vacío ya que no se requiere inicialización adicional
        pass
    
    def _get_start_row(self, data_from_sheet):
        """
        Determina el índice de la fila que contiene los nombres de columnas.
        
        Se asume que la fila que tiene más de 9 valores no nulos es la que contiene
        los nombres de las columnas.

        Args:
            data_from_sheet (DataFrame): DataFrame obtenido de la hoja de Excel.

        Returns:
            int: Índice de la fila de cabecera. Si no se encuentra, retorna 0.
        """
        for i in range(data_from_sheet.shape[0]):
            row = data_from_sheet.iloc[i]
            # Se asume que la fila de cabecera contiene al menos 10 valores no nulos.
            if row.count() > 9:
                return i 
        return 0
    
    def _delimit_data(self, data_from_sheet):
        """
        Delimita el área de datos a partir de la hoja de Excel.
        
        Una vez identificada la fila de cabecera, se extraen las filas a partir de la
        siguiente (start_row + 1) y se seleccionan las columnas de la 2 a la 11 (índices 1 a 10).

        Args:
            data_from_sheet (DataFrame): DataFrame completo obtenido de la hoja de Excel.

        Returns:
            DataFrame: Subconjunto de datos delimitados según filas y columnas de interés.
        """
        start_row = self._get_start_row(data_from_sheet)
        # Se extraen las filas posteriores a la cabecera y las columnas de interés
        return data_from_sheet.iloc[start_row + 1:, 1:11]
    
    def extract_data_from_sheet(self, path):
        """
        Extrae y procesa los datos de la hoja "CUADRO 4" de un archivo Excel.
        
        Lee el archivo Excel sin cabecera definida, delimita el área de datos de interés
        y filtra las filas que tengan menos de 5 valores no nulos.

        Args:
            path (str): Ruta del archivo Excel.

        Returns:
            DataFrame: Datos procesados y filtrados.
        """
        # Leer el archivo Excel sin cabecera (header=None) para la hoja "CUADRO 4"
        data = pd.read_excel(path, sheet_name="CUADRO 4", header=None, dtype=str)
        # Delimitar el DataFrame para extraer la sección de interés
        data = self._delimit_data(data)
        # Eliminar filas con menos de 5 valores no nulos y reiniciar el índice
        data = data.dropna(thresh=5, ignore_index=True)
        return data
