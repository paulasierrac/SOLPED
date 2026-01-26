import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
from typing import List, Optional
from pathlib import Path
from Config.settings import CONFIG_EMAIL

# # Configuración por defecto (puede ser sobrescrita al instanciar)
# CONFIG_EMAIL = {
#     "smtp_server": "smtp.office365.com",
#     "smtp_port": 587,
#     "email": "steven.navarro@netapplications.com.co",
#     "password": "Hop57050",  # IMPORTANTE: Cambiar por variable de entorno en producción
# }


class EmailSender:
    def __init__(
        self,
        smtp_server: str = None,
        smtp_port: int = None,
        email: str = None,
        password: str = None,
    ):
        """
        Inicializa el cliente de envío de correos

        Args:
            smtp_server: Servidor SMTP (por defecto usa CONFIG_EMAIL)
            smtp_port: Puerto SMTP (por defecto usa CONFIG_EMAIL)
            email: Tu dirección de correo (por defecto usa CONFIG_EMAIL)
            password: Tu contraseña o app password (por defecto usa CONFIG_EMAIL)
        """
        self.smtp_server = smtp_server or CONFIG_EMAIL["smtp_server"]
        self.smtp_port = smtp_port or CONFIG_EMAIL["smtp_port"]
        self.email = email or CONFIG_EMAIL["email"]
        self.password = password or CONFIG_EMAIL["password"]

    def leer_excel(self, archivo_excel: str) -> pd.DataFrame:
        """
        Lee el archivo Excel con la estructura de correos

        Args:
            archivo_excel: Ruta al archivo Excel

        Returns:
            DataFrame con los datos de los correos
        """
        try:
            # Intentar leer con diferentes engines para mejor compatibilidad
            df = pd.read_excel(archivo_excel, engine="openpyxl")

            # Limpiar espacios en blanco de las columnas
            df.columns = df.columns.str.strip()

            return df
        except Exception as e:
            print(f"Error al leer el archivo Excel: {e}")
            return None

    def enviar_correo(
        self,
        destinatario: str,
        asunto: str,
        cuerpo: str,
        cc: Optional[List[str]] = None,
        bcc: Optional[List[str]] = None,
        adjuntos: Optional[List[str]] = None,
    ) -> bool:
        """
        Envía un correo electrónico

        Args:
            destinatario: Email del destinatario
            asunto: Asunto del correo
            cuerpo: Cuerpo del mensaje (puede ser HTML)
            cc: Lista de correos en copia
            bcc: Lista de correos en copia oculta
            adjuntos: Lista de rutas de archivos a adjuntar

        Returns:
            True si se envió exitosamente, False en caso contrario
        """
        try:
            # Crear mensaje
            mensaje = MIMEMultipart()
            mensaje["From"] = self.email
            mensaje["To"] = destinatario
            mensaje["Subject"] = asunto

            if cc:
                mensaje["Cc"] = ", ".join(cc)
            if bcc:
                mensaje["Bcc"] = ", ".join(bcc)

            # Agregar cuerpo del mensaje
            mensaje.attach(MIMEText(cuerpo, "html"))

            # Agregar adjuntos si existen
            if adjuntos:
                for archivo in adjuntos:
                    if os.path.exists(archivo):
                        self._adjuntar_archivo(mensaje, archivo)
                    else:
                        print(f"Advertencia: El archivo {archivo} no existe")

            # Conectar y enviar
            with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
                server.starttls()
                server.login(self.email, self.password)

                # Construir lista de destinatarios
                destinatarios = [destinatario]
                if cc:
                    destinatarios.extend(cc)
                if bcc:
                    destinatarios.extend(bcc)

                server.sendmail(self.email, destinatarios, mensaje.as_string())

            print(f"✓ Correo enviado exitosamente a {destinatario}")
            return True

        except Exception as e:
            print(f"✗ Error al enviar correo a {destinatario}: {e}")
            return False

    def enviar_correo_personalizado(
        self,
        destinatario: str,
        asunto: str,
        cuerpo: str,
        adjuntos: Optional[List[str]] = None,
        cc: Optional[List[str]] = None,
        bcc: Optional[List[str]] = None,
    ) -> bool:
        """
        Envía un correo con estructura personalizada. Es un alias para enviar_correo.
        """
        print(f"📧 Iniciando envío personalizado para: {destinatario}")
        return self.enviar_correo(
            destinatario=destinatario,
            asunto=asunto,
            cuerpo=cuerpo,
            cc=cc,
            bcc=bcc,
            adjuntos=adjuntos,
        )

    def _adjuntar_archivo(self, mensaje: MIMEMultipart, ruta_archivo: str):
        """
        Adjunta un archivo al mensaje

        Args:
            mensaje: Objeto MIMEMultipart
            ruta_archivo: Ruta del archivo a adjuntar
        """
        nombre_archivo = Path(ruta_archivo).name

        with open(ruta_archivo, "rb") as archivo:
            parte = MIMEBase("application", "octet-stream")
            parte.set_payload(archivo.read())

        encoders.encode_base64(parte)
        parte.add_header(
            "Content-Disposition", f"attachment; filename={nombre_archivo}"
        )
        mensaje.attach(parte)

    def procesar_excel_y_enviar(
        self,
        archivo_excel: str,
        codigo_correo: Optional[int] = None,
        columna_codigo: str = "codemailparameter",
        columna_destinatario: str = "toemailparameter",
        columna_asunto: str = "asuntoemailparameter",
        columna_cuerpo: str = "bodyemailparameter",
        columna_cc: Optional[str] = "ccemailparameter",
        columna_bcc: Optional[str] = "bccemailparameter",
        columna_adjuntos: Optional[str] = None,
        adjuntos_dinamicos: Optional[List[str]] = None,
        separador_adjuntos: str = ";",
    ) -> dict:
        """
        Procesa el Excel y envía todos los correos

        Args:
            archivo_excel: Ruta al archivo Excel
            codigo_correo: Código específico para filtrar (ej: 1, 2, 3). Si es None, envía todos
            columna_codigo: Nombre de la columna con el código de correo
            columna_destinatario: Nombre de la columna con el email destino
            columna_asunto: Nombre de la columna con el asunto
            columna_cuerpo: Nombre de la columna con el cuerpo del mensaje
            columna_cc: Nombre de la columna con correos CC
            columna_bcc: Nombre de la columna con correos BCC
            columna_adjuntos: Nombre de la columna con rutas de archivos (separados por ;)
            adjuntos_dinamicos: Lista de adjuntos a incluir en TODOS los correos (sobrescribe Excel)
            separador_adjuntos: Caracter para separar múltiples adjuntos

        Returns:
            Diccionario con estadísticas de envío
        """
        # DEBUG: Imprimir ruta del archivo
        # print(f"🔍 DEBUG: Leyendo archivo: {archivo_excel}")
        # print(f"🔍 DEBUG: Archivo existe: {os.path.exists(archivo_excel)}")

        df = self.leer_excel(archivo_excel)

        if df is None:
            return {"exitosos": 0, "fallidos": 0, "total": 0}

        # DEBUG: Mostrar columnas encontradas
        # print(f"🔍 DEBUG: Columnas en el DataFrame: {df.columns.tolist()}")
        # print(f"🔍 DEBUG: Buscando columna: '{columna_destinatario}'")
        # print(f"🔍 DEBUG: ¿Columna existe?: {columna_destinatario in df.columns}")

        # Filtrar por código si se proporciona
        if codigo_correo is not None:
            df = df[df[columna_codigo] == codigo_correo]
            if df.empty:
                print(f"⚠️  No se encontraron correos con código {codigo_correo}")
                return {"exitosos": 0, "fallidos": 0, "total": 0}
            print(
                f"📧 Enviando correos con código {codigo_correo} ({len(df)} correo(s))\n"
            )

        exitosos = 0
        fallidos = 0

        for idx, fila in df.iterrows():
            # Extraer datos de la fila - USAR ACCESO DIRECTO EN LUGAR DE .get()
            try:
                destinatario_raw = fila[columna_destinatario]
            except KeyError:
                print(f"⚠️  ERROR: Columna '{columna_destinatario}' no encontrada")
                print(f"   Columnas disponibles: {fila.index.tolist()}")
                fallidos += 1
                continue

            # DEBUG TEMPORAL
            # print(
            #    f"🔍 DEBUG Fila {idx + 2}: destinatario_raw = {repr(destinatario_raw)}, tipo = {type(destinatario_raw).__name__}"
            # )

            # Convertir a string solo si no es NaN/None
            if pd.isna(destinatario_raw):
                destinatario = ""
            else:
                destinatario = str(destinatario_raw).strip()

            # print(f"🔍 DEBUG Fila {idx + 2}: destinatario procesado = '{destinatario}'")

            asunto = str(fila.get(columna_asunto, "Sin asunto"))
            cuerpo = str(fila.get(columna_cuerpo, ""))

            # Validar que el destinatario no esté vacío
            if (
                not destinatario
                or destinatario == "nan"
                or destinatario == ""
                or pd.isna(destinatario_raw)
            ):
                # print(f"⚠️  Fila {idx + 2}: Destinatario vacío, omitiendo...")
                fallidos += 1
                continue

            # Procesar CC
            cc = None
            if columna_cc and columna_cc in df.columns:
                cc_raw = fila.get(columna_cc)
                if not pd.isna(cc_raw):
                    cc_str = str(cc_raw).strip()
                    if cc_str and cc_str != "nan" and cc_str != "":
                        cc = [
                            email.strip()
                            for email in cc_str.split(",")
                            if email.strip()
                        ]

            # Procesar BCC
            bcc = None
            if columna_bcc and columna_bcc in df.columns:
                bcc_raw = fila.get(columna_bcc)
                if not pd.isna(bcc_raw):
                    bcc_str = str(bcc_raw).strip()
                    if bcc_str and bcc_str != "nan" and bcc_str != "":
                        bcc = [
                            email.strip()
                            for email in bcc_str.split(",")
                            if email.strip()
                        ]

            # Procesar adjuntos
            adjuntos = None

            # Si hay adjuntos dinámicos, usar esos (tienen prioridad)
            if adjuntos_dinamicos:
                adjuntos = adjuntos_dinamicos
            # Si no, buscar en el Excel
            elif columna_adjuntos and columna_adjuntos in df.columns:
                adj_raw = fila.get(columna_adjuntos)
                if not pd.isna(adj_raw):
                    adj_str = str(adj_raw).strip()
                    if adj_str and adj_str != "nan" and adj_str != "":
                        adjuntos = [
                            adj.strip()
                            for adj in adj_str.split(separador_adjuntos)
                            if adj.strip()
                        ]

            # Enviar correo
            if self.enviar_correo(destinatario, asunto, cuerpo, cc, bcc, adjuntos):
                exitosos += 1
            else:
                fallidos += 1

        total = exitosos + fallidos
        print(f"\n{'='*50}")
        print(f"Resumen de envío:")
        print(f"Total de correos: {total}")
        print(f"Exitosos: {exitosos}")
        print(f"Fallidos: {fallidos}")
        print(f"{'='*50}")

        return {"exitosos": exitosos, "fallidos": fallidos, "total": total}


# Ejemplo de uso
if __name__ == "__main__":

    # 1. Instanciar la clase EmailSender (usa la configuración por defecto)
    sender = EmailSender()

    # Opción 2: Sobrescribir configuración si es necesario (COMENTADO)
    # sender = EmailSender(
    #     smtp_server='smtp.gmail.com',
    #     smtp_port=587,
    #     email='otro@gmail.com',
    #     password='otra_contraseña'
    # )

    # --- PRUEBA 1: Usando el NUEVO método personalizado (enviar_correo_personalizado) ---
    print("\n--- PRUEBA 1: Envío Personalizado (Método Nuevo) ---")
    exito_personalizado = sender.enviar_correo_personalizado(
        destinatario="destinatario_personalizado@ejemplo.com",
        asunto="Correo de Prueba vía Método Personalizado",
        cuerpo="<p>Mensaje HTML enviado directamente con el nuevo método.</p>",
        adjuntos=["archivo1.pdf"],  # Asegúrate de que este archivo exista en la ruta
        cc=["info@netapplications.com.co"],
    )

    if exito_personalizado:
        print("Envío personalizado exitoso (Método Nuevo).")
    else:
        print("Envío personalizado fallido (Método Nuevo).")

    # --- PRUEBA 2: Usando el método enviar_correo ORIGINAL (Envío Individual) ---
    # Nota: Esta prueba es redundante si se usa la Prueba 1, pero se incluye para probar la función original.
    print("\n--- PRUEBA 2: Envío Individual (Método Original) ---")
    sender.enviar_correo(
        destinatario="otro_destinatario@example.com",
        asunto="Prueba de correo Original",
        cuerpo="<h1>Hola</h1><p>Este es un correo de prueba usando el método 'enviar_correo'.</p>",
        adjuntos=["archivo1.pdf", "documento.xlsx"],  # Opcional
    )

    # --- PRUEBA 3: Usando el método de Procesamiento Masivo (procesar_excel_y_enviar) ---
    print("\n--- PRUEBA 3: Procesamiento Masivo (Excel) ---")
    resultados = sender.procesar_excel_y_enviar(
        archivo_excel="correos.xlsx",  # Asegúrate de que este archivo exista
        codigo_correo=1,
        adjuntos_dinamicos=["reporte.pdf", "log.txt"],
    )
    print(f"Resumen del procesamiento por Excel: {resultados}")
