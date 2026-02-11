from Repositories.parametros import ParametrosRepository
from Config.settings import DB_CONFIG

schema = DB_CONFIG.get("schema")

_CONFIG_CACHE = None


def init_config():
    """Carga configuración desde BD una sola vez"""
    global _CONFIG_CACHE

    if _CONFIG_CACHE is not None:
        return

    parametros = ParametrosRepository(schema)

    _CONFIG_CACHE = parametros.cargar_parametros()


def in_config(nombre, default=None):
    if _CONFIG_CACHE is None:
        raise RuntimeError(
            "Configuración no inicializada. "
            "Ejecute HU00_DespliegueAmbiente antes de usar in_config()"
        )

    return _CONFIG_CACHE.get(nombre, default)
