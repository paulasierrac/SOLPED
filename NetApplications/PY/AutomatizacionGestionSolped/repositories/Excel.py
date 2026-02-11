from Config.database import Database
from Config.settings import DB_CONFIG

schemadb = DB_CONFIG.get("schema")


class ExcelRepo:

    def __init__(self, schema: str):
        self.schema = schema or schemadb

    # -----------------------------
    # CREAR COLUMNAS DINÁMICAS
    # -----------------------------
    @staticmethod
    def _construir_columnas(columnas: list[str]) -> str:
        return ",\n".join(f"{col} VARCHAR(MAX) NULL" for col in columnas)

    # -----------------------------
    # CREAR TABLA TEMPORAL
    # -----------------------------
    def crear_tabla_temp(self, tabla: str, columnas: list[str]) -> bool:

        tabla_temp = f"{tabla}_temp"
        if not columnas:
            raise ValueError("La lista de columnas está vacía")

        columnas_sql = ",\n".join(f"[{col}] NVARCHAR(MAX)" for col in columnas)

        query = f"""
        IF OBJECT_ID('{self.schema}.{tabla_temp}', 'U') IS NOT NULL
            DROP TABLE {self.schema}.{tabla_temp};

        CREATE TABLE {self.schema}.{tabla_temp} (
            {columnas_sql}
        );
        """

        try:
            with Database.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(query)
                conn.commit()
                cursor.close()
                print("Tabla temporal creada con exito")
            return True
        except Exception as e:
            print(f"Error creando tabla temporal {tabla_temp}: {e}")
            return False

    # -----------------------------
    # CREAR TABLA FINAL
    # -----------------------------
    def crear_tabla_final(self, tabla: str, columnas: list[str]) -> bool:

        columnas_sql = ExcelRepo._construir_columnas(columnas)

        query = f"""
        IF OBJECT_ID('{self.schema}.{tabla}', 'U') IS NOT NULL
            DROP TABLE {self.schema}.{tabla};

        CREATE TABLE {self.schema}.{tabla} (
            {columnas_sql},
            EstadoRegistro VARCHAR(50) NOT NULL DEFAULT 'Pendiente'
        );
        """

        try:
            with Database.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(query)
                conn.commit()
                print("Tabla final creada con exito")
                cursor.close()
            return True
        except Exception as e:
            print(f"Error creando tabla final {tabla}: {e}")
            return False

    # -----------------------------
    # BULK A TEMP + TRANSFERENCIA
    # -----------------------------
    def ejecutar_bulk_dinamico(self, ruta_txt: str, tabla: str, columnas: list[str]):

        tabla_temp = f"{tabla}_temp"
        columnas_sql = ", ".join(columnas)

        bulk_query = f"""
        BULK INSERT {self.schema}.{tabla_temp}
        FROM '{ruta_txt}'
        WITH (
            FIRSTROW = 2,
            FIELDTERMINATOR = ';',
            ROWTERMINATOR = '0x0a',
            CODEPAGE = '65001',
            TABLOCK
        );
        """

        insert_query = f"""
        INSERT INTO {self.schema}.{tabla} ({columnas_sql}, EstadoRegistro)
        SELECT {columnas_sql}, 'Pendiente'
        FROM {self.schema}.{tabla_temp};
        """

        drop_query = f"""
        DROP TABLE {self.schema}.{tabla_temp};
        """

        try:
            with Database.get_connection() as conn:
                conn.autocommit = True
                cursor = conn.cursor()

                exrepo = ExcelRepo("GestionSolped")

                if not exrepo.crear_tabla_temp(tabla, columnas):
                    return

                exrepo.crear_tabla_final(tabla, columnas)

                cursor.execute(bulk_query)
                cursor.execute(insert_query)
                cursor.execute(drop_query)

                cursor.close()

            print(f"Carga BULK ejecutada correctamente en {tabla}")

        except Exception as e:
            print(f"Error durante BULK dinámico en {tabla}: {e}")

    # -----------------------------
    # OBTENER SOLO PENDIENTES
    # -----------------------------
    def obtener_valores(self, tabla: str):

        query = f"""
        SELECT*
        FROM {self.schema}.{tabla}
        WHERE EstadoRegistro = 'Pendiente'
        """

        with Database.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query)

            columnas = [col[0] for col in cursor.description]
            rows = cursor.fetchall()
            cursor.close()

            return [dict(zip(columnas, fila)) for fila in rows]
