from datetime import datetime
from config.database import Database
from config.settings import DB_CONFIG

schemadb = DB_CONFIG["schema"]

class TicketInsumoRepo:

    def __init__(self, schema: str):
        self.schema = schema or schemadb

    def obtener_por_codigo(self, codigo):
        query = f"""
            SELECT *
            FROM {self.schema}.TicketInsumo
            WHERE Codigo = ?
        """
        with Database.get_connection(dictionary=True) as conn:
            cursor = conn.cursor()
            cursor.execute(query, (codigo,))
            cursor.close()
            return cursor.fetchone()
        
    def crear(self, codigo: str, maquina: str):
        query = f"""
            INSERT INTO {self.schema}.TicketInsumo
            (Codigo, fechainsercion, estado, numeroreintentos, maquina)
            VALUES (?, ?, ?, ?, ?)
        """
        with Database.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(
                query,
                (codigo, datetime.now(), "PENDIENTE", 0, maquina)
            )

    def actualizar_estado(
        self,
        codigo,
        estado,
        observaciones=None,
        incrementar_reintento=False,
        finalizar=False
    ):
        query = f"""
            UPDATE {self.schema}.TicketInsumo
            SET estado = %s,
                observaciones = %s,
                numeroreintentos = numeroreintentos + %s,
                fechamodificacion = %s,
                fechafin = %s
            WHERE Codigo = %s
        """

        fechafin = datetime.now() if finalizar else None
        incremento = 1 if incrementar_reintento else 0

        with Database.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(
                query,
                (
                    estado,
                    observaciones,
                    incremento,
                    datetime.now(),
                    fechafin,
                    codigo
                )
            )