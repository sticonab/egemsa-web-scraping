from datetime import datetime

class DataTransformer:
    """
    Clase para transformar datos extraídos de un workbook.

    Proporciona métodos para limpiar los datos, asignar nombres de columnas y agregar
    columnas adicionales relacionadas con el periodo y la fecha de creación.
    """

    @classmethod
    def transform_data(cls, data, year, mes):
        """
        Transforma el DataFrame aplicando las siguientes operaciones:
          - Elimina espacios en blanco de cada valor de tipo cadena.
          - Renombra las columnas con nombres predefinidos.
          - Agrega una columna 'Periodo' con el formato 'año-mes'.
          - Agrega una columna 'FechaCreacion' con la fecha y hora actual.

        Args:
            data (DataFrame): DataFrame extraído del workbook.
            año (str/int): Año correspondiente al periodo.
            mes (str/int): Mes correspondiente al periodo.

        Returns:
            DataFrame: DataFrame transformado con columnas renombradas y columnas adicionales.
        """
        # Aplicar strip a todos los valores que sean cadenas para eliminar espacios en blanco
        data = data.applymap(lambda x: x.strip() if isinstance(x, str) else x)

        # Renombrar las columnas del DataFrame
        data.columns = [
            'Empresa',
            'BarraTransferencia',
            'TipoUsuario',
            'TipoContrato',
            'EntregaRetiro',
            'ClienteCentralGeneracion',
            'EnergiaMWh',
            'ValorizacionSoles',
            'RentaCongestionLicitacion',
            'RentaCongestionBilateral'
        ]

        # Agregar la columna 'Periodo' combinando el año y mes proporcionados
        data['Periodo'] = f'{year}-{mes}'

        # Agregar la columna 'FechaCreacion' con la fecha y hora actual
        data['FechaCreacion'] = datetime.now().strftime('%Y-%m-%dT%H:%M:%S')

        return data
