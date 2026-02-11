from Config.database import Database

class ParametrosRepository:

    def __init__(self, schema: str):
        self.schema = schema 
        
    def cargar_parametros(self) -> dict:
        conn = Database.get_connection()
        cursor = conn.cursor()

        query = f"""
            SELECT Nombre, Valor
            FROM {self.schema}.Parametros
        """
        cursor.execute(query)

        config = {}
        for nombre, valor in cursor.fetchall():
            config[nombre] = valor

        cursor.close()
        conn.close()

        return config