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
