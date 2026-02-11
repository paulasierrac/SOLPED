 
"""
Script para obtener secretos de Azure Key Vault usando autenticación de usuario
Requiere: pip install azure-keyvault-secrets azure-identity
"""
from azure.keyvault.secrets import SecretClient # pyright: ignore[reportMissingImports]
from azure.identity import InteractiveBrowserCredential, DefaultAzureCredential # type: ignore
from azure.identity import AzureCliCredential # pyright: ignore[reportMissingImports]
import sys
def obtener_secreto_keyvault(vault_url= "https://appmonitoreo.vault.azure.net/", secret_name="SAP-Usuario", use_browser_auth=True):


    # vault_url = https://appmonitoreo.vault.azure.net/
    # secret_name = nombre de la variable 
    # use_browser_auth=True  // usuario con acceso  
    """
    Obtiene un secreto de Azure Key Vault usando autenticación de usuario
    Args:
        vault_url (str): URL del Key Vault (ej: https://mi-vault.vault.azure.net/)
        secret_name (str): Nombre del secreto a obtener
        use_browser_auth (bool): Si True usa autenticación por navegador
    Returns:
        str: Valor del secreto
    """
    try:
        # Autenticación del usuario
        credential = AzureCliCredential()
        """
        if use_browser_auth:
            # Autenticación interactiva por navegador
            print("Se abrirá el navegador para autenticación...")
            credential = InteractiveBrowserCredential()
        else:
            # Intenta usar credenciales del entorno (Azure CLI, variables de entorno, etc.)
            credential = DefaultAzureCredential()
        """            
        # Crear cliente del Key Vault


        client = SecretClient(vault_url=vault_url, credential=credential)
        # Obtener el secreto
        print(f"Obteniendo secreto '{secret_name}' del vault...")
        secret = client.get_secret(secret_name)
        print(f"✓ Secreto obtenido exitosamente")
        print(f"  - Nombre: {secret.name}")
        print(f"  - Versión: {secret.properties.version}")
        print(f"  - Creado: {secret.properties.created_on}")
        return secret.value
    except Exception as e:
        print(f"✗ Error al obtener el secreto: {str(e)}")
        sys.exit(1)





"""
def listar_secretos(vault_url):
    
    Lista todos los secretos disponibles en el Key Vault
    
    try:
        credential = InteractiveBrowserCredential()
        client = SecretClient(vault_url=vault_url, credential=credential)
        print("\nSecretos disponibles en el vault:")
        print("-" * 50)
        secret_properties = client.list_properties_of_secrets()
        for secret in secret_properties:
            print(f"  - {secret.name}")
    except Exception as e:
        print(f"✗ Error al listar secretos: {str(e)}")
 
"""

    # conexion = "ERP-CORPORATIVO-PRODUCCION"
    # mandante = "410"
    # usuario = "CGRPA065"

    # password = "sT1f%4L*" 
    # idioma = "ES"

    # print(usuario)
    # print(password)

    # abrirSap = AbrirSAPLogon()
    # if abrirSap:
    #     print(" SAP Logon 750 se encuentra abierto")
    # else:
    #     print(" SAP Logon 750 abierto ")

    # #session = ConectarSAP(conexion, mandante, usuario, password, idioma)