from Config.database import Database
from Config.settings import DB_CONFIG

schemadb = DB_CONFIG["schema"]

class ControlHURepo:

    """
    Repositorio encargado de la gestión y persistencia de los estados de las 
    Historias de Usuario (HU) en la base de datos de control.

    Esta clase permite centralizar el seguimiento de la ejecución del bot,
    asegurando que cada HU tenga un registro de su estado actual y la máquina 
    donde se está procesando.

    Attributes:
        schema (str): Esquema de la base de datos donde reside la tabla ControlHU.
    """
    
    def __init__(self, schema: str):
        self.schema = schema or schemadb

    def ActualizarEstadoHU(
        self, 
        iHuId: int,
        iNombreHU: str, 
        estado: int,
        activa: int,
        maquina: str
    ) -> bool:
        """
        Inicializa el repositorio con un esquema específico.

        Args:
            schema (str): Nombre del esquema (ej. 'GestionSolped'). 
                          Si es None, se utiliza el valor por defecto schemadb.
        """
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
            VALUES (?, ?, ?, ?, ?);
        """
        
        with Database.get_connection() as conn:
            try:
                cursor = conn.cursor()
                cursor.execute(
                    query,
                    (
                        iHuId,
                        iNombreHU,
                        estado,
                        activa,
                        maquina,
                        iHuId,
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