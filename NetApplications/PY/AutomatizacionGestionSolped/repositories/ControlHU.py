from Config.database import Database
from Config.settings import DB_CONFIG

schemadb = DB_CONFIG["schema"]

class ControlHURepo:
    
    def __init__(self, schema: str):
        self.schema = schema or schemadb
        
    def actualizar_estado_hu(
        self, 
        hu: int,
        iNombreHU: str, 
        estado: int,
        activa: int,
        maquina: str
    ) -> bool:
        query = f"""
        MERGE {self.schema}.ControlHU AS target
        USING (SELECT ? AS HU) AS source
        ON target.HU = source.HU
        WHEN MATCHED THEN
            UPDATE SET
                NombreHU = ?,
                Estado = ?,
                Activa = ?,
                Maquina = ?
        WHEN NOT MATCHED THEN
            INSERT (HU, NombreHU, Estado, Activa, Maquina) 
            VALUES (?, ?, ?, ?, ?)
        """
        
        with Database.get_connection() as conn:
            try:
                cursor = conn.cursor()
                cursor.execute(
                    query,
                    (
                        hu,
                        iNombreHU,
                        estado,
                        activa,
                        maquina,
                        hu,
                        iNombreHU,
                        estado,
                        activa,
                        maquina
                    )
                )
                conn.commit()
                return True
            except Exception as e:
                print(f"Error al actualizar el estado de HU: {e}")
                return False