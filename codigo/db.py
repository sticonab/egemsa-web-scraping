from sqlalchemy import create_engine

class DatabaseManager:
    """
    Clase para gestionar operaciones de base de datos utilizando SQLAlchemy.
    
    Permite ejecutar consultas SQL y almacenar datos desde un DataFrame de Pandas.
    """

    def __init__(self, connection_string: str):
        """
        Inicializa la conexión a la base de datos usando la cadena de conexión.

        Args:
            connection_string (str): Cadena de conexión para la base de datos.
        """
        self.engine = create_engine(connection_string)

    def get_single_value(self, sql_query: str):
        """
        Ejecuta una consulta SQL y retorna el primer valor (la primera columna de la primera fila).

        Args:
            sql_query (str): Consulta SQL a ejecutar.

        Returns:
            El primer valor obtenido de la consulta o None si no se encontraron resultados.
        """
        # Se utiliza una conexión cruda para obtener un cursor y ejecutar la consulta
        with self.engine.raw_connection().cursor() as cursor:
            cursor.execute(sql_query)
            result = cursor.fetchone()
            # Retorna el primer valor de la fila o None si no se encontró ningún resultado
            return result[0] if result else None

    def get_first_row(self, sql_query: str):
        """
        Ejecuta una consulta SQL y retorna la primera fila completa del conjunto de resultados.

        Args:
            sql_query (str): Consulta SQL a ejecutar.

        Returns:
            tuple: La primera fila del conjunto de resultados o None si no hay resultados.
        """
        with self.engine.raw_connection().cursor() as cursor:
            cursor.execute(sql_query)
            results = cursor.fetchall()
            # Retorna la primera fila completa o None si la lista de resultados está vacía
            return results[0] if results else None

    def store_data_pandas(self, data, table_name: str):
        """
        Almacena un DataFrame de Pandas en la tabla especificada de la base de datos.

        Los datos se insertan en el esquema 'bronce'. Si la tabla ya existe, se añaden nuevos registros.

        Args:
            data (DataFrame): DataFrame con los datos a almacenar.
            table_name (str): Nombre de la tabla donde se almacenarán los datos.
        """
        data.to_sql(
            name=table_name,
            con=self.engine,
            index=False,
            schema='bronce',
            if_exists='append'
        )
