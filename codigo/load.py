from db import DatabaseManager
from dateutil.relativedelta import relativedelta
import os
from dotenv import load_dotenv

# Cargar variables de entorno desde el archivo .env
env_file_path = 'E:/BI_Comercial/ETL_Python/R006_ETL_ValorizacionEnergia/globals.env'
load_dotenv(env_file_path)


class DataLoader:
    """
    Clase para gestionar operaciones ETL en la base de datos de valorización de energía.
    
    Esta clase se conecta a la base de datos utilizando credenciales definidas en las variables
    de entorno y proporciona métodos para obtener el periodo a procesar y para cargar datos en
    la tabla de destino.
    """
    
    def __init__(self):
        """
        Inicializa la conexión a la base de datos y define nombres de tablas y esquemas.
        """
        # Obtener credenciales y configuración de conexión desde las variables de entorno
        driver = os.getenv('SQL_DRIVE')
        server = os.getenv('SERVER_NAME')
        database = os.getenv('DB')
        admin = os.getenv('ADMIN')
        password = os.getenv('PSWD')
        
        # Construir la cadena de conexión para SQL Server usando pyodbc
        connection_string = f'mssql+pyodbc://{admin}:{password}@{server}/{database}?driver={driver}'
        
        # Definir nombres de tablas y esquemas para las distintas etapas del proceso ETL
        self.staging_table = '[plata].[ValorizacionEnergia]'
        self.landing_table = 'ValorizacionEnergia'
        self.dim_month_table = '[oro].[Periodo]'
        self.generals_table = '[plata].[Generales]'
        
        # Crear una instancia de la clase de acceso a la base de datos
        self.db = DatabaseManager(connection_string)
    
    def get_date_to_retrieve(self):
        """
        Obtiene la fecha a partir de la cual se deben recuperar nuevos datos.
        
        Ejecuta una consulta para obtener la fecha máxima existente en la tabla de staging, uniendo
        con la dimensión de meses. Si no se encuentra ninguna fecha (valor None), se obtiene el 
        periodo inicial definido en la tabla de variables generales. Si se encuentra una fecha,
        se incrementa en un mes para procesar el siguiente periodo.
        
        Returns:
            tuple: Una tupla (año, mes) representando el periodo a procesar.
        """
        # Consulta para obtener la fecha máxima registrada en la tabla de staging
        sql_query = f'''
            SELECT 
                MAX(Fecha)
            FROM (
                SELECT DISTINCT tm.Fecha
                FROM {self.staging_table} AS pc
                INNER JOIN {self.dim_month_table} AS tm
                    ON pc.Periodo = tm.Periodo
            ) AS sq
        '''
        date = self.db.get_single_value(sql_query)
        
        if date is None:
            # Si no existe fecha en la tabla de staging, obtener el periodo inicial desde la tabla de variables generales
            sql_query = f'''
                SELECT tm.Fecha
                FROM {self.dim_month_table} AS tm
                INNER JOIN (
                    SELECT Valor
                    FROM {self.generals_table}
                    WHERE IdVariablesGlobales = 'VAR01'
                ) AS pin
                    ON tm.Periodo = pin.Valor
            '''
            date = self.db.get_single_value(sql_query)
        else:
            # Si se encontró una fecha, incrementar en un mes para procesar el siguiente periodo
            date = date + relativedelta(months=1)
        
        # Convertir la fecha a cadena y separar el año y el mes
        date_parts = str(date).split('-')
        return (date_parts[0], date_parts[1])
    
    def load_data_to_landing(self, data):
        """
        Carga los datos en la tabla de landing de la base de datos.
        
        Args:
            data (DataFrame): DataFrame con los datos a almacenar en la base de datos.
        """
        self.db.store_data_pandas(data, self.landing_table)
