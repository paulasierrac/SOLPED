from config.database import Database
from config.settings import DB_CONFIG

schemadb = DB_CONFIG.get("schema")

class Querys:

    def __init__(self, schema):
        self.schema = schema or schemadb

    def fetch_all(self, tabla):
        query = f"SELECT * FROM {self.schema}.{tabla}"

        with Database.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query)
            
            columnas = [col[0] for col in cursor.description]
            rows = cursor.fetchall()
            cursor.close()
            diccionario = [dict(zip(columnas, fila)) for fila in rows]

            return rows, diccionario