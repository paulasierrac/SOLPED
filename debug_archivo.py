import pandas as pd
import os

# Ruta del archivo
archivo = r"C:\Users\CGRPA009\Documents\SOLPED-main\SOLPED\NetApplications\PY\AutomatizacionGestionSolped\Insumo\EnvioCorreos.xlsx"

print("=" * 80)
print("VERIFICACIÓN PROFUNDA DEL ARCHIVO")
print("=" * 80)

# Verificar que el archivo existe
print(f"\n1. ¿El archivo existe? {os.path.exists(archivo)}")
print(f"   Ruta completa: {archivo}")
print(f"   Tamaño: {os.path.getsize(archivo)} bytes")
print(f"   Última modificación: {os.path.getmtime(archivo)}")

# Leer el archivo
df = pd.read_excel(archivo, engine="openpyxl")
df.columns = df.columns.str.strip()

print(f"\n2. Archivo leído: {len(df)} filas, {len(df.columns)} columnas")

# Verificar código 1
print("\n3. ANÁLISIS DETALLADO CÓDIGO 1:")
df_codigo1 = df[df["codemailparameter"] == 1]

if len(df_codigo1) > 0:
    for idx, row in df_codigo1.iterrows():
        print(f"\n   Fila índice: {idx} (Excel fila {idx + 2})")

        # Obtener el valor exactamente como lo hace EmailSender
        destinatario_raw = row.get("toemailparameter")

        print(f"   - Valor raw: {repr(destinatario_raw)}")
        print(f"   - Tipo: {type(destinatario_raw)}")
        print(f"   - pd.isna(): {pd.isna(destinatario_raw)}")

        if pd.isna(destinatario_raw):
            destinatario = ""
            print(f"   - Es NaN, se convierte a: '{destinatario}'")
        else:
            destinatario = str(destinatario_raw).strip()
            print(f"   - Después de str().strip(): '{destinatario}'")

        # Verificar la condición que usa EmailSender
        condicion = (
            not destinatario
            or destinatario == "nan"
            or destinatario == ""
            or pd.isna(destinatario_raw)
        )
        print(f"\n   ¿Pasa validación como vacío? {condicion}")
        print(f"   - not destinatario: {not destinatario}")
        print(f"   - destinatario == 'nan': {destinatario == 'nan'}")
        print(f"   - destinatario == '': {destinatario == ''}")
        print(f"   - pd.isna(destinatario_raw): {pd.isna(destinatario_raw)}")

# Verificar código 2
print("\n" + "=" * 80)
print("4. ANÁLISIS DETALLADO CÓDIGO 2:")
df_codigo2 = df[df["codemailparameter"] == 2]

if len(df_codigo2) > 0:
    for idx, row in df_codigo2.iterrows():
        print(f"\n   Fila índice: {idx} (Excel fila {idx + 2})")

        destinatario_raw = row.get("toemailparameter")

        print(f"   - Valor raw: {repr(destinatario_raw)}")
        print(f"   - Tipo: {type(destinatario_raw)}")
        print(f"   - pd.isna(): {pd.isna(destinatario_raw)}")

        if pd.isna(destinatario_raw):
            destinatario = ""
            print(f"   - Es NaN, se convierte a: '{destinatario}'")
        else:
            destinatario = str(destinatario_raw).strip()
            print(f"   - Después de str().strip(): '{destinatario}'")

        condicion = (
            not destinatario
            or destinatario == "nan"
            or destinatario == ""
            or pd.isna(destinatario_raw)
        )
        print(f"\n   ¿Pasa validación como vacío? {condicion}")
        print(f"   - not destinatario: {not destinatario}")
        print(f"   - destinatario == 'nan': {destinatario == 'nan'}")
        print(f"   - destinatario == '': {destinatario == ''}")
        print(f"   - pd.isna(destinatario_raw): {pd.isna(destinatario_raw)}")

print("\n" + "=" * 80)
print("5. VERIFICACIÓN DE COLUMNAS:")
print(f"   Columnas del DataFrame: {df.columns.tolist()}")
print(f"   ¿'toemailparameter' existe? {'toemailparameter' in df.columns}")

print("\n" + "=" * 80)
print("6. PRIMERAS 5 FILAS COMPLETAS:")
print(df[["codemailparameter", "toemailparameter", "asuntoemailparameter"]].head())

print("\n" + "=" * 80)
