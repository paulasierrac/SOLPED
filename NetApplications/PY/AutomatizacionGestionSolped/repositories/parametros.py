from config.database import Database
import logging

logger = logging.getLogger(__name__)

class ParametrosRepository:

    @staticmethod
    def cargar_parametros() -> dict:
        conn = Database.get_connection()
        cursor = conn.cursor()

        query = """
            SELECT Nombre, Valor
            FROM GestionSolped.parametros
        """
        cursor.execute(query)

        config = {}
        for nombre, valor in cursor.fetchall():
            config[nombre] = valor

        cursor.close()
        conn.close()

        return config