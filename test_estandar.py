import ast
import os

def obtener_funciones(ruta_archivo):
    try:
        with open(ruta_archivo, "r", encoding="utf-8") as f:
            tree = ast.parse(f.read())
        
        # Extrae solo las definiciones de funciones (def)
        funciones = [node.name for node in ast.walk(tree) if isinstance(node, ast.FunctionDef)]
        return funciones
    except Exception as e:
        return [f"Error al leer archivo: {str(e)}"]

def analizar_carpetas(carpetas):
    # Obtener la ruta base del script actual
    base_path = os.path.dirname(os.path.abspath(__file__))
    
    for carpeta in carpetas:
        ruta_carpeta = os.path.join(base_path, carpeta)
        if not os.path.exists(ruta_carpeta):
            print(f"Carpeta no encontrada: {ruta_carpeta}")
            continue
            
        print(f"\n{'='*20} Analizando carpeta: {carpeta} {'='*20}")
        
        # Listar archivos en la carpeta
        for archivo in os.listdir(ruta_carpeta):
            if archivo.endswith(".py"):
                ruta_completa = os.path.join(ruta_carpeta, archivo)
                funciones = obtener_funciones(ruta_completa)
                
                print(f"\nArchivo: {archivo}")
                if funciones:
                    for func in funciones:
                        print(f"  > {func}")
                else:
                    print("  (Sin funciones definidas)")

if __name__ == "__main__":
    carpetas_objetivo = ["Funciones", "HU"]
    
    # Redirigir stdout a un archivo
    import sys
    original_stdout = sys.stdout
    with open("function_list.txt", "w", encoding="utf-8") as f:
        sys.stdout = f
        analizar_carpetas(carpetas_objetivo)
        sys.stdout = original_stdout
    
    print("Análisis completado. Resultados guardados en function_list.txt")