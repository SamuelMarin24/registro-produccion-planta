"""
Configuración de la aplicación.

Hay dos niveles:
  1. Rutas compartidas (maestras, base de OPs, Excel consolidado, carpetas de
     intercambio): se leen de variables de entorno o del archivo .env, para no
     dejar rutas de red dentro del código.
  2. Configuración local de cada PC (qué máquina tiene asignada): se guarda en
     config.json, que se genera solo la primera vez que se abre la app.
"""

import json
import os
from pathlib import Path

try:
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).parent / ".env")
except ImportError:
    pass  # Sin python-dotenv se usan las variables de entorno del sistema

BASE_DIR = Path(__file__).parent
CONFIG_FILE = BASE_DIR / "config.json"

# Carpeta compartida donde viven los datos (normalmente una unidad de red)
CARPETA_DATOS = Path(os.getenv("CARPETA_DATOS", BASE_DIR / "datos"))

# Carpetas del flujo archivo → consolidador
CARPETA_PENDIENTES = Path(os.getenv("CARPETA_PENDIENTES", CARPETA_DATOS / "PENDIENTES"))
CARPETA_PROCESADOS = Path(os.getenv("CARPETA_PROCESADOS", CARPETA_DATOS / "PROCESADOS"))

# Respaldo en el PC de la máquina cuando la red no está disponible
CARPETA_RESPALDO_LOCAL = BASE_DIR / "_pendientes_local"

# Parámetros de operación
INTERVALO_RECARGA_MIN = int(os.getenv("INTERVALO_RECARGA_MIN", "10"))
CICLO_SEGUNDOS = int(os.getenv("CICLO_SEGUNDOS", "30"))

# Identidad visual (se puede personalizar sin tocar el resto del código)
NOMBRE_APP = os.getenv("NOMBRE_APP", "Registro de Producción")
TEXTO_LOGO = os.getenv("TEXTO_LOGO", "PLANTA")


class Config:
    """Configuración local de este PC (qué máquina tiene asignada)."""

    def __init__(self):
        self._data = {}
        self._load()

    def _load(self):
        if CONFIG_FILE.exists():
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    self._data = json.load(f)
            except (json.JSONDecodeError, OSError):
                self._data = {}

    def is_configured(self):
        return bool(self._data.get("maquina"))

    def get_machine_config(self):
        return self._data

    def save_machine_config(self, cod_maquina: str):
        self._data["maquina"] = cod_maquina
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(self._data, f, ensure_ascii=False, indent=2)

    # ── Rutas compartidas ─────────────────────────────────────────────────────

    def get_ruta_maestras(self):
        """Maestras de máquinas, áreas, eventos y operarios."""
        return os.getenv("ARCHIVO_MAESTRAS", str(CARPETA_DATOS / "MAESTRAS.xlsx"))

    def get_ruta_op(self):
        """Base de órdenes de producción (para autocompletar al digitar la OP)."""
        return os.getenv("ARCHIVO_OP", str(CARPETA_DATOS / "BD_ORDENES_PRODUCCION.xlsx"))

    def get_ruta_excel(self):
        """Excel consolidado que escribe únicamente el consolidador."""
        return os.getenv("ARCHIVO_CONSOLIDADO", str(CARPETA_DATOS / "BD_PRODUCCION.xlsx"))

    # ── Carpetas del flujo de archivos ────────────────────────────────────────

    def get_carpeta_pendientes(self):
        """Bandeja de entrada: cada máquina deja aquí su archivo JSON."""
        return str(CARPETA_PENDIENTES)

    def get_carpeta_procesados(self):
        """Respaldo de los archivos ya volcados al Excel, en subcarpetas mensuales."""
        return str(CARPETA_PROCESADOS)

    def get_carpeta_respaldo_local(self):
        """Respaldo en el PC de la máquina si la red no está disponible al guardar."""
        return str(CARPETA_RESPALDO_LOCAL)
