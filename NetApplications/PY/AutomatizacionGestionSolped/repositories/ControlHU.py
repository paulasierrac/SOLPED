from Config.Database import Database
from Config.settings import DB_CONFIG

schemadb = DB_CONFIG.get("schema")

class ControlHURepository:

    def __init__(self, schema: str):
        self.schema = schema or schemadb
        
    @staticmethod
    def upsert_control_hu(self, hu_id: int, nombre_hu: str, estado: int, activa: int, maquina: str):
        """
        Inserta o actualiza la HU
        """
        query = f"""
        MERGE {self.schema}.ControlHU as target
        USING (SELECT ? AS HU) AS source
        ON target.hu = source.HU
        WHEN MATCHED THEN
            UPDATE SET
                Estado = ?,
                Activa = ?,
                Maquina = ?,
                FechaActualizacion = SYSDATETIME()
        WHEN NOT MATCHED THEN
            INSERT (HU, NombreHU, Estado, Activa, Maquina)
            VALUES (?, ?, ?, ?, ?);
        """

        with Database.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(
                query,
                (
                    hu_id,
                    estado,
                    activa,
                    maquina,
                    hu_id,
                    nombre_hu,
                    estado,
                    activa,
                    maquina
                )
            )
            conn.commit()