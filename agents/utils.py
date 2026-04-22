"""
Utilidades compartidas entre todos los agentes FERRAFORME.
"""
import re
from pathlib import Path

# Carpeta temporal para imágenes extraídas del Excel
IMAGENES_TEMP_PATH = Path(__file__).parent.parent / "imagenes_temp"

# Ruta del cache local de Odoo
CACHE_PATH = Path(__file__).parent.parent / "odoo_cache.pkl"
CACHE_MAX_HORAS = 24


def _safe_float(val, default: float = 0.0) -> float:
    """Convierte val a float de forma segura; devuelve default si no es numérico."""
    if val is None:
        return default
    try:
        return float(val)
    except (ValueError, TypeError):
        return default


def _sanitizar_nombre_archivo(nombre: str) -> str:
    """Convierte nombre de producto a nombre de archivo seguro."""
    nombre = nombre.strip().lower()
    nombre = re.sub(r"[^\w\s-]", "", nombre, flags=re.UNICODE)
    nombre = re.sub(r"[\s]+", "_", nombre)
    return nombre[:100] or "producto"


def _phash_imagen(image_data: bytes):
    """Genera perceptual hash de una imagen. Devuelve objeto imagehash o None."""
    try:
        import imagehash
        import io
        from PIL import Image
        img = Image.open(io.BytesIO(image_data)).convert("RGB")
        return imagehash.phash(img)
    except Exception:
        return None


def _similitud_nombres(a: str, b: str) -> float:
    from difflib import SequenceMatcher
    return SequenceMatcher(None, a.lower().strip(), b.lower().strip()).ratio()
