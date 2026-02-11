import os
import re

# Mapping of Old Name -> New Name
RENAME_MAP = {
    "EnviarCorro": "EnviarCorreo",
    "ColsultarSolped": "ConsultarSolped",
    "EnviarCorro_personalizado": "EnviarCorreoPersonalizado",
    "SetGuiComboBoxkey": "EstablecerLlaveGuiComboBox",
    "ObtenerTextoCampoGuitextfield": "ObtenerTextoCampoGui",
    "set_GuiTextField_text": "EstablecerTextoCampoGui",
    "obtener_valor": "ObtenerValor",
    "buscar_objeto_por_id_parcial": "BuscarObjetoPorIdParcial",
    "obtener_importe_por_denominacion": "ObtenerImportePorDenominacion",
    "ObtenerColumnasdf": "ObtenerColumnasDf",
    "get_importesCondiciones": "ObtenerImportesCondiciones",
    "_get_shell": "ObtenerShell",
    "get_line": "ObtenerLinea",
    "get_all_text": "ObtenerTodoElTexto",
    "set_editable_line": "EstablecerLineaEditable",
    "replace_in_text": "ReemplazarEnTexto",
    "buscar_tabla": "BuscarTabla",
    "buscar_boton": "BuscarBoton",
    "buscar_combobox": "BuscarCombobox",
    "buscar_ctextfield": "BuscarCampoTextoC",
    "buscar_textfield": "BuscarCampoTexto",
    "buscar_tab": "BuscarPestana",
    "buscar_recursivo": "BuscarRecursivo",
    "obtener_secreto_keyvault": "ObtenerSecretoKeyvault",
    "determinar_estado_reporte": "DeterminarEstadoReporte",
    "generar_reporte_resumen": "GenerarReporteResumen",
    "imprimir_resumen_reporte": "ImprimirResumenReporte",
    "validar_estructura_fila": "ValidarEstructuraFila",
    "limpiar_datos_fila": "LimpiarDatosFila",
    "exportar_a_csv": "ExportarACsv",
    "buscar_columna": "BuscarColumna",
    "extraerDatosReporte": "ExtraerDatosReporte",
    "AppendHipervinculoObservaciones": "AgregarHipervinculoObservaciones",
    "obtenerFilaExpSolped": "ObtenerFilaExpSolped",
    "abrir_sap_logon": "AbrirSapLogon",
    "ConectarSAP": "ConectarSap",
    "validarLoginDiag": "ValidarLoginDiag",
    "estado_ok": "EstadoOk",
    "generar_envios_y_reporte": "GenerarEnviosYReporte",
    "log": "RegistrarLog",
    "extract_text_from_pdf": "ExtraerTextoDePdf",
    "parse_oc": "ProcesarOc",
    "parse_empresa": "ProcesarEmpresa",
    "parse_proveedor": "ProcesarProveedor",
    "parse_proveedor_sr": "ProcesarProveedorSr",
    "limpiar_nombre": "LimpiarNombre",
    "safe_name": "NombreSeguro",
    "organizar_pdf": "OrganizarPdf",
    "EnviarCorro_simulado": "EnviarCorreoSimulado",
    "parse_correos": "ProcesarCorreos",
    "obtener_tipo_proveedor": "ObtenerTipoProveedor"
}

def replace_in_file(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        original_content = content
        
        # Sort keys by length (descending)
        sorted_keys = sorted(RENAME_MAP.keys(), key=len, reverse=True)
        
        for old_name in sorted_keys:
            new_name = RENAME_MAP[old_name]
            pattern = r'\b' + re.escape(old_name) + r'\b'
            content = re.sub(pattern, new_name, content)
            
        if content != original_content:
            print(f"Modifying: {file_path}")
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(content)
            return True
        return False
    except Exception as e:
        print(f"Error processing {file_path}: {e}")
        return False

def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    # Directories to scan. Root is '.', plus subdirs
    directories = [base_dir, os.path.join(base_dir, "Funciones"), os.path.join(base_dir, "HU")]
    
    print(f"Scanning directories: {directories}")
    
    count = 0
    for d in directories:
        if not os.path.exists(d):
            print(f"Skipping missing directory: {d}")
            continue
            
        for filename in os.listdir(d):
            if filename.endswith(".py") and filename != "refactor_names.py":
                file_path = os.path.join(d, filename)
                if replace_in_file(file_path):
                    count += 1
                    
    print(f"Refactoring completed. Modified {count} files.")

if __name__ == "__main__":
    main()
