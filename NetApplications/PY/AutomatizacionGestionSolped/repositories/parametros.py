from Config.Database import Database
from Config.settings import DB_CONFIG
import logging

logger = logging.getLogger(__name__)
schemadb = DB_CONFIG.get("schema")


class ParametrosRepository:
    
    def __init__(self, schema: str):
        self.schema = schema
        
    def cargar_parametros(self) -> dict:
        """
        Carga parámetros desde la tabla de configuración en la base de datos.
        """
        query = f"""
            SELECT nombre, valor
            FROM {self.schema}.parametros
        """
        with Database.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query)

            config = {}
            for nombre, valor in cursor.fetchall():
                config[nombre] = valor

            cursor.close()

        return config
