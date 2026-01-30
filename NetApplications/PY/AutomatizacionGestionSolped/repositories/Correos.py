from Config.Database import Database
from Config.settings import DB_CONFIG

schemadb = DB_CONFIG.get("schema")

class CorreosRepo:

    def __init__(self, schema: str):
        self.schema = schema or schemadb
        
    def ObtenerParametrosCorreo(self, cod_email: int):
        query = f""""
            SELECT * FROM {self.schema}.ParametrosCorreo WHERE CodEmailParamter = ?
        """

        with Database.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, cod_email)
            fila = cursor.fetchone()

            if not fila:
                raise ValueError(f"No existe configuracion de correo para el codigo: {cod_email}")
            
            return {
                "to": fila.TOEmailParameter,
                "cc": fila.CCEmailParameter,
                "bcc": fila.BCCEmailParameter,
                "asunto": fila.AsuntoEmailParameter,
                "body": fila.BodyEmailParameter,
                "is_html": bool(fila.IsHTMLEmailParameter)
            }