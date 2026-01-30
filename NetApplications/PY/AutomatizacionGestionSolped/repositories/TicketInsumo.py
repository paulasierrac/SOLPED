from Config.Database import Database
from datetime import datetime
from Config.settings import DB_CONFIG

schemadb = DB_CONFIG.get("schema")

class TicketInsumoRepository:
    
    def __init__(self, conn, schema):
        self.conn = conn or Database.get_connection()
        self.schema = schema or schemadb
        
    def obtener_por_codigo(self, codigo: str):
        
        query = f"""
            SELECT *
            FROM {self.schema}.TicketsInsumo
            WHERE codigo = ?
        """
        with self.conn(dictionary = True) as cursor:
            cursor.execute(query, (codigo,))
            cursor.close()
            return cursor.fetchone()
        
    def crear(self):
        query = f"""
            INSERT INTO {self.schema}.TicketsInsumo 
            (codigo, fechainsercion, estado, numeroreintentos, maquina)
            VALUES (?, ?, ?, ?, ?)
        """
        
        with self.conn.cursor() as cursor:
            cursor.execute(query, (self.codigo, datetime.now(), "PENDIENTE", 0, self.maquina))
            self.conn.commit()
            cursor.close()    
    
    
    def actualizar_estado(
        self, 
        codigo: str, 
        estado: str, 
        observaciones: str = None, 
        incrementar_reintento=False,
        finalizar=False
    ):
        query = f"""
            UPDATE {self.schema}.TicketsInsumo
            SET estado = ?, observaciones = ?, numeroreintentos = ?, fechamodificacion = ?, fechafin = ?
            WHERE codigo = ?
        """
        
        fechafin = datetime.now() if finalizar else None
        incremento = 1 if incrementar_reintento else 0
        
        with self.conn.cursor() as cursor:
            cursor.execute(query,
                            (estado, 
                             observaciones, 
                             incremento, 
                             datetime.now(), 
                             fechafin, 
                             codigo))
            self.conn.commit()
            cursor.close()