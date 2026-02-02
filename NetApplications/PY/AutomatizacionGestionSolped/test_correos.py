import pandas as pd

# Ruta de tu archivo Excel
archivo = r"C:\Users\CGRPA009\Documents\SOLPED-main\SOLPED\NetApplications\PY\AutomatizacionGestionSolped\Insumo\EnvioCorreos.xlsx"

print("=" * 70)
print("ANÁLISIS DEL ARCHIVO EnvioCorreos.xlsx")
print("=" * 70)

# Leer el archivo
df = pd.read_excel(archivo, engine="openpyxl")
df.columns = df.columns.str.strip()

print("\n✓ Archivo leído correctamente")
print(f"  Total de filas: {len(df)}")
print(f"  Total de columnas: {len(df.columns)}")

print("\n📋 COLUMNAS ENCONTRADAS:")
columnas_esperadas = [
    "codemailparameter",
    "actividad",
    "toemailparameter",
    "ccemailparameter",
    "bccemailparameter",
    "asuntoemailparameter",
    "bodyemailparameter",
    "observacion",
]

for col in df.columns:
    estado = "✓" if col in columnas_esperadas else "⚠️"
    print(f"  {estado} {col}")

# Verificar si faltan columnas esperadas
columnas_faltantes = set(columnas_esperadas) - set(df.columns)
if columnas_faltantes:
    print(f"\n⚠️  COLUMNAS FALTANTES: {columnas_faltantes}")

print("\n🔍 ANÁLISIS POR CÓDIGO:")

# Códigos que quieres revisar
codigos = [1, 2, 3, 5, 28, 29]

for codigo in codigos:
    print(f"\n{'='*70}")
    print(f"CÓDIGO {codigo}:")
    df_codigo = df[df["codemailparameter"] == codigo]

    if len(df_codigo) == 0:
        print(f"  ⚠️  No se encontraron filas con código {codigo}")
        continue

    print(f"  ✓ Encontradas {len(df_codigo)} fila(s)")

    for idx, row in df_codigo.iterrows():
        print(f"\n  📧 Fila {idx + 2}:")

        # Verificar destinatario
        destinatario = row.get("toemailparameter")
        print(f"     toemailparameter: ", end="")
        if pd.isna(destinatario):
            print("❌ VACÍO (NaN)")
        elif str(destinatario).strip() == "":
            print("❌ VACÍO (string vacío)")
        else:
            print(f"✓ '{destinatario}'")

        # Verificar asunto
        asunto = row.get("asuntoemailparameter")
        print(f"     asuntoemailparameter: ", end="")
        if pd.isna(asunto):
            print("⚠️  VACÍO")
        else:
            print(f"✓ '{asunto}'")

        # Verificar cuerpo
        bodyemail = row.get("bodyemailparameter")
        print(f"     bodyemailparameter: ", end="")
        if pd.isna(bodyemail):
            print("⚠️  VACÍO")
        else:
            cuerpo_preview = str(bodyemail)[:50]
            print(f"✓ '{cuerpo_preview}...'")

        # Verificar CC (puede estar vacío)
        cc = row.get("ccemailparameter")
        print(f"     ccemailparameter: ", end="")
        if pd.isna(cc) or str(cc).strip() == "":
            print("(vacío - OK)")
        else:
            print(f"'{cc}'")

        # Verificar BCC (puede estar vacío)
        bcc = row.get("bccemailparameter")
        print(f"     bccemailparameter: ", end="")
        if pd.isna(bcc) or str(bcc).strip() == "":
            print("(vacío - OK)")
        else:
            print(f"'{bcc}'")

print("\n" + "=" * 70)
print("RESUMEN:")
print("=" * 70)

# Contar filas con destinatario vacío
filas_sin_destino = df[
    df["toemailparameter"].isna()
    | (df["toemailparameter"].astype(str).str.strip() == "")
]
print(f"⚠️  Filas SIN destinatario: {len(filas_sin_destino)}")
if len(filas_sin_destino) > 0:
    print(f"   Filas afectadas: {[i+2 for i in filas_sin_destino.index.tolist()]}")

# Contar filas con destinatario válido
filas_con_destino = df[
    ~df["toemailparameter"].isna()
    & (df["toemailparameter"].astype(str).str.strip() != "")
]
print(f"✓ Filas CON destinatario: {len(filas_con_destino)}")
if len(filas_con_destino) > 0:
    print(
        f"   Códigos con destinatario: {filas_con_destino['codemailparameter'].unique().tolist()}"
    )

print("\n" + "=" * 70)


# # Enviar correo de inicio (código 1)
# # EnviarNotificacionCorreo(codigo_correo=1, task_name=task_name)
# archivo_descargado = rf"{RUTAS['PathReportes']}/Reporte_1300139268_10.txt"
# # Enviar correo de inicio (código 2 adjunto)
# EnviarNotificacionCorreo(
#     codigo_correo=54, task_name=task_name, adjuntos=[archivo_descargado]
# )

# exito_personalizado = EnviarCorreoPersonalizado(
#     destinatario="soporte_critico@netapplications.com.co",
#     asunto="Alerta Crítica: El servicio X ha fallado",
#     cuerpo=(
#         "<h1>Error Inesperado</h1>"
#         "<p>El proceso de sincronización ha fallado en la etapa de validación de datos.</p>"
#         "<p><strong>Revisar logs en:</strong> \\\\servidor\\logs\\errores.txt</p>"
#     ),
#     task_name=task_name,
#     adjuntos=["C:/Archivos/log_error_20251204.txt"],
#     cc=["steven.navarro@netapplications.com.co"],
# )

# if exito_personalizado:
#     print(f"Notificación enviada exitosamente exito_personalizado.")
# else:
#     print(f"Fallo al enviar la notificación exito_personalizado.")

# NUMERO_SOLPED = "8000012345"
# DESTINOS = ["usuario.revision@empresa.com", "supervisor@empresa.com"]
# RAZONES_VALIDACION = (
#     "1. El centro de costo asignado no es válido para el tipo de material.\n"
#     "2. La cantidad solicitada supera el límite sin aprobación especial."
# )

# # Llamada a la función
# exito_notificacion = NotificarRevisionManualSolped(
#     destinatarios=DESTINOS,
#     numero_solped=NUMERO_SOLPED,
#     validaciones=RAZONES_VALIDACION,
# )

# exito_notificacion = NotificarRevisionManualSolped(
#     destinatarios=["usuario.revision@empresa.com", "supervisor@empresa.com"],
#     numero_solped="8000012345",
#     validaciones=(
#         "1. El centro de costo asignado no es válido para el tipo de material.\n"
#         "2. La cantidad solicitada supera el límite sin aprobación especial."
#     ),
# )

# if exito_notificacion:
#     print(f"Notificación enviada exitosamente para SOLPED {NUMERO_SOLPED}.")
# else:
#     print(f"Fallo al enviar la notificación para SOLPED {NUMERO_SOLPED}.")
