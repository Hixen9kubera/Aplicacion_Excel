"""
Agente Odoo — Fase D de la arquitectura multi-agente FERRAFORME.

Responsabilidades:
- Cache local y Supabase de productos Odoo
- Carga/búsqueda de productos en Odoo via XML-RPC
- Validación de SKUs contra Odoo
- Creación de productos y subida de imágenes
- Búsqueda de producto padre por nombre/imagen

Solo contiene funciones puras (sin dependencia de st.session_state).
Las funciones con st (sincronizar_contadores, generar_sku, crear_clasificacion_en_odoo)
permanecen en app_v3.py hasta la Fase E del grafo LangGraph.
"""
from __future__ import annotations

import base64
import concurrent.futures
import os
import pickle
import re
import xmlrpc.client
from datetime import datetime, timedelta
from pathlib import Path

from langchain_anthropic import ChatAnthropic
from langchain_core.messages import HumanMessage

from agents.utils import (
    CACHE_PATH, CACHE_MAX_HORAS, IMAGENES_TEMP_PATH,
    _safe_float, _phash_imagen, _similitud_nombres,
)


# ── Cache local ───────────────────────────────────────────────────────────────

def guardar_cache_odoo(skus: list[str], productos: list[dict], phashes: dict) -> None:
    """Guarda SKUs, productos (sin imagen) y phashes en cache local."""
    phashes_hex   = {k: str(v) for k, v in phashes.items()}
    prods_sin_img = [{k: v for k, v in p.items() if k != "image_128"} for p in productos]
    cache = {
        "timestamp": datetime.now().isoformat(),
        "skus":      skus,
        "productos": prods_sin_img,
        "phashes":   phashes_hex,
    }
    with open(CACHE_PATH, "wb") as f:
        pickle.dump(cache, f)


def cargar_cache_odoo() -> dict | None:
    """
    Carga el cache local si existe y no tiene más de CACHE_MAX_HORAS horas.
    Devuelve dict con skus, productos, phashes o None.
    """
    if not CACHE_PATH.exists():
        return None
    try:
        with open(CACHE_PATH, "rb") as f:
            cache = pickle.load(f)
        ts = datetime.fromisoformat(cache["timestamp"])
        if datetime.now() - ts > timedelta(hours=CACHE_MAX_HORAS):
            return None
        import imagehash
        phashes = {}
        for k, v in cache.get("phashes", {}).items():
            try:
                phashes[k] = imagehash.hex_to_hash(v)
            except Exception:
                pass
        return {
            "timestamp": ts,
            "skus":      cache["skus"],
            "productos": cache["productos"],
            "phashes":   phashes,
        }
    except Exception:
        return None


def info_cache_odoo() -> str | None:
    """Devuelve string con info del cache (fecha, cantidad) o None si no existe."""
    if not CACHE_PATH.exists():
        return None
    try:
        with open(CACHE_PATH, "rb") as f:
            cache = pickle.load(f)
        ts  = datetime.fromisoformat(cache["timestamp"])
        n   = len(cache.get("skus", []))
        nph = len(cache.get("phashes", {}))
        return f"{ts.strftime('%d/%m/%Y %H:%M')} · {n} SKUs · {nph} imágenes"
    except Exception:
        return None


# ── Supabase — cache persistente ─────────────────────────────────────────────

def _supabase_client():
    """Devuelve cliente Supabase o None si no está configurado."""
    url = os.environ.get("SUPABASE_URL", "")
    key = os.environ.get("SUPABASE_KEY", "")
    if not url or not key:
        return None
    try:
        from supabase import create_client
        return create_client(url, key)
    except Exception:
        return None


def _cargar_desde_supabase() -> dict | None:
    """
    Carga productos desde Supabase (tabla odoo_productos).
    Devuelve dict con skus, productos, phashes, imagen_urls o None.
    """
    sb = _supabase_client()
    if not sb:
        return None
    try:
        import imagehash
        resp = sb.table("odoo_productos").select(
            "sku,nombre,phash_hex,imagen_url,actualizado_en"
        ).execute()
        rows = resp.data or []
        if not rows:
            return None
        skus      = [r["sku"] for r in rows]
        productos = [{"default_code": r["sku"], "name": r.get("nombre", "")} for r in rows]
        phashes, imagen_urls = {}, {}
        for r in rows:
            sku = r["sku"]
            if r.get("phash_hex"):
                try:
                    phashes[sku] = imagehash.hex_to_hash(r["phash_hex"])
                except Exception:
                    pass
            if r.get("imagen_url"):
                imagen_urls[sku] = r["imagen_url"]
        ts_str = max((r.get("actualizado_en", "") for r in rows), default="")
        try:
            ts = datetime.fromisoformat(ts_str.replace("Z", "+00:00"))
        except Exception:
            ts = datetime.now()
        return {"timestamp": ts, "skus": skus, "productos": productos,
                "phashes": phashes, "imagen_urls": imagen_urls}
    except Exception:
        return None


def _sincronizar_a_supabase(prods_odoo: list[dict]) -> dict:
    """
    Sube productos de ODOO a Supabase (phash + imagen al bucket).
    Devuelve dict con phashes e imagen_urls calculados.
    """
    sb = _supabase_client()
    phashes, imagen_urls = {}, {}
    if not sb:
        return {"phashes": phashes, "imagen_urls": imagen_urls}

    bucket = sb.storage.from_("odoo-imagenes")
    rows_to_upsert = []
    prods_con_img  = []

    for p in prods_odoo:
        sku    = p.get("default_code", "")
        nombre = p.get("name", "")
        if not sku:
            continue
        phash_hex = None
        img_b64   = p.get("image_128")
        img_bytes = None
        if img_b64:
            try:
                img_bytes = base64.b64decode(img_b64)
                ph = _phash_imagen(img_bytes)
                if ph is not None:
                    phash_hex    = str(ph)
                    phashes[sku] = ph
            except Exception:
                pass
        rows_to_upsert.append({"sku": sku, "nombre": nombre,
                                "phash_hex": phash_hex, "imagen_url": None})
        if img_bytes:
            prods_con_img.append((sku, img_bytes))

    def _upload_img(args):
        sku, img_bytes = args
        try:
            bucket.upload(f"{sku}.png", img_bytes,
                          {"content-type": "image/png", "x-upsert": "true"})
            return sku, bucket.get_public_url(f"{sku}.png")
        except Exception:
            return sku, None

    with concurrent.futures.ThreadPoolExecutor(max_workers=20) as ex:
        for sku, url in ex.map(_upload_img, prods_con_img):
            if url:
                imagen_urls[sku] = url

    for row in rows_to_upsert:
        row["imagen_url"] = imagen_urls.get(row["sku"])

    try:
        for i in range(0, len(rows_to_upsert), 100):
            sb.table("odoo_productos").upsert(rows_to_upsert[i:i+100]).execute()
    except Exception:
        pass

    return {"phashes": phashes, "imagen_urls": imagen_urls}


# ── Búsqueda de similares ─────────────────────────────────────────────────────

def buscar_similares_odoo(
    imagen_data: bytes | None,
    nombre_producto: str,
    productos_odoo: list[dict],
    phashes_odoo: dict,
    umbral_phash: int = 12,
    umbral_nombre: float = 0.70,
) -> list[dict]:
    """
    Busca productos similares en ODOO por imagen (phash) y nombre (fuzzy).
    Devuelve lista ordenada por relevancia.
    """
    phash_nuevo = _phash_imagen(imagen_data) if imagen_data else None
    similares   = []

    for prod in productos_odoo:
        sku        = prod.get("default_code", "")
        nombre_odo = prod.get("name", "")
        if not sku:
            continue
        sim_nombre = _similitud_nombres(nombre_producto, nombre_odo)
        sim_img, por_imagen = None, False
        if phash_nuevo is not None and sku in phashes_odoo:
            dist       = phash_nuevo - phashes_odoo[sku]
            sim_img    = round(1.0 - dist / 64.0, 3)
            por_imagen = dist <= umbral_phash
        por_nombre = sim_nombre >= umbral_nombre
        if por_imagen or por_nombre:
            similares.append({
                "sku":              sku,
                "nombre":           nombre_odo,
                "similitud_imagen": sim_img,
                "similitud_nombre": round(sim_nombre, 3),
                "por_imagen":       por_imagen,
                "por_nombre":       por_nombre,
                "producto_odoo":    prod,
            })

    similares.sort(key=lambda x: (x["por_imagen"], x["similitud_nombre"]), reverse=True)
    return similares[:5]


def _buscar_padre_en_odoo_por_nombre(
    nombre_base: str,
    prods_odoo: list[dict],
    phashes_odoo: dict,
    imagen_data: bytes | None = None,
    umbral: float = 0.75,
) -> dict | None:
    """
    Busca en prods_odoo el mejor candidato a template padre por nombre similar.
    Pondera opcionalmente con similitud de imagen (phash).
    Devuelve dict {prod, score, sim_nombre} o None.
    """
    if not nombre_base or not prods_odoo:
        return None
    phash_nuevo = _phash_imagen(imagen_data) if imagen_data else None
    mejor, mejor_score = None, 0.0
    for prod in prods_odoo:
        nombre_odo = prod.get("name", "")
        sim_nombre = _similitud_nombres(nombre_base, nombre_odo)
        sim_img    = 0.0
        sku        = prod.get("default_code", "")
        if phash_nuevo is not None and sku in phashes_odoo:
            dist    = phash_nuevo - phashes_odoo[sku]
            sim_img = max(0.0, 1.0 - dist / 64.0)
        score = sim_nombre * 0.7 + sim_img * 0.3
        if score > mejor_score and sim_nombre >= umbral:
            mejor_score = score
            mejor = {"prod": prod, "score": round(score, 3), "sim_nombre": round(sim_nombre, 3)}
    return mejor


# ── Carga desde Odoo ──────────────────────────────────────────────────────────

def cargar_skus_odoo(url: str, db: str, username: str, password: str) -> tuple[list[str], str | None]:
    """Obtiene todos los SKUs (default_code) de Odoo via XML-RPC."""
    try:
        common    = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/common", allow_none=True)
        uid       = common.authenticate(db, username, password, {})
        if not uid:
            return [], "Credenciales inválidas o usuario sin acceso."
        models    = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/object", allow_none=True)
        productos = models.execute_kw(
            db, uid, password,
            "product.template", "search_read",
            [[["default_code", "!=", False]]],
            {"fields": ["default_code"], "limit": 0},
        )
        return [p["default_code"] for p in productos if p.get("default_code")], None
    except Exception as e:
        return [], str(e)


def cargar_todos_productos_odoo(url: str, db: str, username: str, password: str) -> list[dict]:
    """Carga nombre, SKU e imagen de todos los productos con SKU en Odoo."""
    try:
        common = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/common", allow_none=True)
        uid    = common.authenticate(db, username, password, {})
        if not uid:
            return []
        models = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/object", allow_none=True)
        return models.execute_kw(
            db, uid, password,
            "product.template", "search_read",
            [[["default_code", "!=", False]]],
            {"fields": ["default_code", "name", "image_128"], "limit": 0},
        )
    except Exception:
        return []


def cargar_detalle_productos_odoo(skus: list[str]) -> list[dict]:
    """Trae detalle de productos Odoo por lista de SKUs (usando .env)."""
    url      = os.environ.get("ODOO_URL", "")
    db       = os.environ.get("ODOO_DB", "")
    username = os.environ.get("ODOO_USER", "")
    password = os.environ.get("ODOO_PASSWORD", "")
    if not all([url, db, username, password]) or not skus:
        return []
    try:
        common = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/common", allow_none=True)
        uid    = common.authenticate(db, username, password, {})
        if not uid:
            return []
        models = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/object", allow_none=True)
        return models.execute_kw(
            db, uid, password,
            "product.template", "search_read",
            [[["default_code", "in", skus]]],
            {"fields": ["default_code", "name", "description_sale", "categ_id",
                        "image_128", "list_price", "standard_price"], "limit": 0},
        )
    except Exception:
        return []


# ── SKU conflict resolution ───────────────────────────────────────────────────

def aplicar_resoluciones_conflictos(
    productos: list[dict],
    conflictos: list[dict],
    resoluciones: dict,
) -> list[dict]:
    """Aplica decisiones del usuario sobre conflictos SKU a la lista de productos."""
    for conflicto in conflictos:
        idx  = conflicto["idx"]
        res  = resoluciones.get(str(idx), {})
        prod = productos[idx]
        if res.get("tipo") == "mismo":
            prod["sku"]         = res.get("sku_final",   prod.get("sku", ""))
            prod["nombre"]      = res.get("nombre",      prod.get("nombre", ""))
            prod["descripcion"] = res.get("descripcion", prod.get("descripcion", ""))
            prod["categoria"]   = res.get("categoria",   prod.get("categoria", ""))
            prod["atributo"]    = res.get("atributo",    prod.get("atributo", ""))
            if res.get("precio_usd"):
                prod["precio_usd"] = res["precio_usd"]
        else:
            prod["sku"] = conflicto["sku_ajustado"]
    return productos


def validar_sku_vs_odoo(sku_propuesto: str, skus_odoo: list[str]) -> dict:
    """
    Verifica que el número secuencial del SKU propuesto sea mayor al máximo
    existente en Odoo para ese SUBCAT.

    Devuelve:
      conflicto      : bool
      skus_odoo_match: list
      sku_ajustado   : str
      nuevo_num      : int | None  — contador ajustado (el caller puede actualizar session_state)
    """
    m = re.match(r"^([A-Z]{2,4})-(\d{4})-([A-Z]{2,4})$", sku_propuesto)
    if not m:
        return {"conflicto": False, "skus_odoo_match": [], "sku_ajustado": sku_propuesto,
                "nuevo_num": None}

    subcat, num_str, atrib = m.group(1), m.group(2), m.group(3)
    num_propuesto = int(num_str)

    patron_sub = re.compile(rf"^{re.escape(subcat)}-(\d{{4}})-", re.IGNORECASE)
    numeros_odoo = [int(p.group(1)) for s in skus_odoo if (p := patron_sub.match(s))]

    if not numeros_odoo:
        return {"conflicto": False, "skus_odoo_match": [], "sku_ajustado": sku_propuesto,
                "nuevo_num": None}

    max_num_odoo = max(numeros_odoo)
    if num_propuesto <= max_num_odoo:
        nuevo_num     = max_num_odoo + 1
        prefijo       = f"{subcat}-{num_str}-"
        coincidencias = [s for s in skus_odoo if s.upper().startswith(prefijo.upper())]
        sku_ajustado  = f"{subcat}-{nuevo_num:04d}-{atrib}"
        return {"conflicto": True, "skus_odoo_match": coincidencias,
                "sku_ajustado": sku_ajustado, "nuevo_num": nuevo_num}

    return {"conflicto": False, "skus_odoo_match": [], "sku_ajustado": sku_propuesto,
            "nuevo_num": None}


# ── Creación de productos ─────────────────────────────────────────────────────

# Grupos de atributos para clasificación en Odoo
_ATRIB_COLORES    = {"NEG","BLN","GRI","ROJ","AZL","VER","AMA","ROS","NAR","MOR","CAF","BEI","MUL","PLA","DOR"}
_ATRIB_TALLAS     = {"XS","S","M","L","XL","UNI"}
_ATRIB_MATERIALES = {"MAD","MET","TEL","CUE"}


def _cbm_por_pieza_odoo(prod: dict) -> float:
    """Calcula CBM por pieza desde los campos disponibles."""
    cbm_pz  = _safe_float(prod.get("cbm_por_pieza"))
    cbm_mc  = _safe_float(prod.get("cbm_master_carton"))
    cbm_tot = _safe_float(prod.get("cbm_total_sku"))
    pzs_cja = _safe_float(prod.get("piezas_x_caja"))
    pzs_tot = _safe_float(prod.get("piezas_total"))
    if cbm_pz > 0:
        return cbm_pz
    if cbm_mc > 0 and pzs_cja > 0:
        return cbm_mc / pzs_cja
    if cbm_tot > 0 and pzs_tot > 0:
        return cbm_tot / pzs_tot
    return 0.0


def _crear_producto_en_odoo(
    prod: dict,
    image_data: bytes,
    uid, models, db: str, pw: str,
    cat_cache: dict,
    tipo_cambio: float = 19.0,
    costo_por_m3: float = 0.0,
    atributos_map: dict | None = None,
) -> str | None:
    """
    Crea un producto nuevo en Odoo con todos sus datos e imagen.
    cat_cache: dict mutable {nombre_cat: cat_id} para evitar lookups repetidos.
    atributos_map: dict {cod: descripcion} (ATRIBUTOS de app_v3).
    Devuelve mensaje de error o None si fue exitoso.
    """
    _CAT_FORZADA = "PRUEBAS_AGENTE"
    atributos_map = atributos_map or {}

    try:
        nombre_cat = _CAT_FORZADA
        if nombre_cat not in cat_cache:
            ids_cat = models.execute_kw(db, uid, pw, "product.category", "search",
                                        [[["name", "=", nombre_cat]]])
            cat_cache[nombre_cat] = ids_cat[0] if ids_cat else models.execute_kw(
                db, uid, pw, "product.category", "create", [{"name": nombre_cat}]
            )
        cat_id = cat_cache[nombre_cat]

        nombre_prod = (prod.get("nombre") or prod.get("sku", ""))
        if not nombre_prod.endswith("_test"):
            nombre_prod += "_test"

        precio_usd     = _safe_float(prod.get("precio_usd"))
        precio_mxn     = round(precio_usd * tipo_cambio, 2)
        cbm_pz         = _cbm_por_pieza_odoo(prod)
        costo_cbm_pz   = round(cbm_pz * costo_por_m3, 2)
        costo_unitario = round(precio_mxn + costo_cbm_pz, 2)

        def _es_chino(txt: str) -> bool:
            import re as _re
            return bool(_re.search(r'[\u4e00-\u9fff]', txt or ""))

        def _traducir_campo(txt: str) -> str:
            if not txt or not _es_chino(txt):
                return txt
            try:
                _llm = ChatAnthropic(model="claude-haiku-4-5-20251001", max_tokens=200)
                _r   = _llm.invoke([HumanMessage(content=(
                    "Traduce al español de forma clara y comercial el siguiente texto. "
                    "Responde SOLO con la traducción:\n\n" + txt
                ))])
                return _r.content.strip() or txt
            except Exception:
                return txt

        desc_partes = []
        if prod.get("descripcion"):
            desc_partes.append(prod["descripcion"])
        material = _traducir_campo(str(prod.get("material") or ""))
        uso      = _traducir_campo(str(prod.get("uso")      or ""))
        if material:
            desc_partes.append(f"Material: {material}")
        if uso:
            desc_partes.append(f"Uso: {uso}")
        if precio_usd > 0:
            desc_partes.append(f"Precio USD: ${precio_usd:.2f}")
        pzs_caja = prod.get("piezas_x_caja")
        if pzs_caja:
            desc_partes.append(f"Piezas por caja: {int(pzs_caja) if str(pzs_caja).replace('.','').isdigit() else pzs_caja}")
        largo, ancho, alto = prod.get("largo_cm"), prod.get("ancho_cm"), prod.get("alto_cm")
        dims = None
        if any([largo, ancho, alto]):
            dims = " × ".join(f"{v}" for v in [largo, ancho, alto] if v)
            desc_partes.append(f"Dimensiones: {dims} cm")

        notas = []
        if prod.get("nombre_alt"):
            notas.append(f"Nombre alternativo: {prod['nombre_alt']}")
        if prod.get("id_guia"):
            notas.append(f"ID Guía / Referencia: {prod['id_guia']}")
        if prod.get("cajas_master"):
            notas.append(f"Cajas master: {prod['cajas_master']}")
        if prod.get("piezas_total"):
            notas.append(f"Piezas totales en contenedor: {prod['piezas_total']}")
        if prod.get("cbm_por_pieza"):
            notas.append(f"CBM por pieza: {prod['cbm_por_pieza']}")
        if prod.get("cbm_master_carton"):
            notas.append(f"CBM master carton: {prod['cbm_master_carton']}")
        if prod.get("cbm_total_sku"):
            notas.append(f"CBM total SKU: {prod['cbm_total_sku']}")
        if prod.get("cbm_inner_carton"):
            notas.append(f"CBM inner carton: {prod['cbm_inner_carton']}")
        if dims:
            notas.append(f"Dimensiones caja: {dims} cm")
        notas.append(f"Costo flete CBM/pieza: ${costo_cbm_pz:.4f} MXN")

        img_b64    = base64.b64encode(image_data).decode("utf-8")
        volumen_pz = cbm_pz if cbm_pz > 0 else None

        vals = {
            "name":             nombre_prod,
            "default_code":     prod.get("sku", ""),
            "description_sale": "\n".join(desc_partes),
            "description":      "\n".join(notas) if notas else False,
            "categ_id":         cat_id,
            "list_price":       costo_unitario,
            "type":             "product",
            "sale_ok":          True,
            "purchase_ok":      True,
            "image_1920":       img_b64,
        }
        if volumen_pz is not None:
            vals["volume"] = volumen_pz
        prod_id = models.execute_kw(db, uid, pw, "product.template", "create", [vals])

        atrib_cod = (prod.get("atributo_cod") or "EST").strip().upper()
        atrib_val = atributos_map.get(atrib_cod)
        if atrib_val and prod_id:
            if atrib_cod in _ATRIB_COLORES:
                atrib_nombre = "Color"
            elif atrib_cod in _ATRIB_TALLAS:
                atrib_nombre = "Talla"
            elif atrib_cod in _ATRIB_MATERIALES:
                atrib_nombre = "Material"
            else:
                atrib_nombre = "Tipo"
            try:
                attr_ids = models.execute_kw(db, uid, pw, "product.attribute", "search",
                                             [[["name", "=", atrib_nombre]]])
                attr_id  = attr_ids[0] if attr_ids else models.execute_kw(
                    db, uid, pw, "product.attribute", "create", [{"name": atrib_nombre}])
                val_ids  = models.execute_kw(db, uid, pw, "product.attribute.value", "search",
                                             [[["name", "=", atrib_val],
                                               ["attribute_id", "=", attr_id]]])
                val_id   = val_ids[0] if val_ids else models.execute_kw(
                    db, uid, pw, "product.attribute.value", "create",
                    [{"name": atrib_val, "attribute_id": attr_id}])
                models.execute_kw(db, uid, pw, "product.template.attribute.line", "create", [{
                    "product_tmpl_id": prod_id,
                    "attribute_id":    attr_id,
                    "value_ids":       [[4, val_id]],
                }])
            except Exception:
                pass

        return None
    except Exception as e:
        return f"Error al crear producto '{prod.get('sku', '?')}' en Odoo: {e}"


def _subir_productos_a_odoo(
    productos: list[dict],
    tipo_cambio: float,
    costo_por_m3: float,
) -> tuple[list[str], list[str]]:
    """
    Sube imágenes a productos YA EXISTENTES en Odoo (por SKU).
    No crea productos nuevos.
    Devuelve (skus_actualizados, errores).
    """
    subidas: list[str] = []
    errores: list[str] = []

    if not IMAGENES_TEMP_PATH.exists():
        return subidas, errores

    odoo_url  = os.environ.get("ODOO_URL", "")
    odoo_db   = os.environ.get("ODOO_DB", "")
    odoo_user = os.environ.get("ODOO_USER", "")
    odoo_pass = os.environ.get("ODOO_PASSWORD", "")

    if not all([odoo_url, odoo_db, odoo_user, odoo_pass]):
        return subidas, ["Faltan credenciales ODOO en .env — imágenes no subidas"]

    try:
        common       = xmlrpc.client.ServerProxy(f"{odoo_url}/xmlrpc/2/common", allow_none=True)
        uid          = common.authenticate(odoo_db, odoo_user, odoo_pass, {})
        if not uid:
            return subidas, ["Credenciales ODOO inválidas"]
        models_proxy = xmlrpc.client.ServerProxy(f"{odoo_url}/xmlrpc/2/object", allow_none=True)
    except Exception as e:
        return subidas, [f"No se pudo conectar a ODOO: {e}"]

    prod_por_sku = {p.get("sku"): p for p in productos if p.get("sku")}

    for archivo in list(IMAGENES_TEMP_PATH.iterdir()):
        if not archivo.is_file():
            continue
        sku = archivo.stem
        if sku not in prod_por_sku:
            continue
        try:
            image_data = archivo.read_bytes()
        except Exception as e:
            errores.append(f"No se pudo leer imagen '{archivo.name}': {e}")
            continue
        try:
            tmpl_ids = models_proxy.execute_kw(
                odoo_db, uid, odoo_pass, "product.template", "search",
                [[["default_code", "=", sku]]],
            )
            if not tmpl_ids:
                errores.append(f"Producto `{sku}` no encontrado en Odoo")
                continue
            img_b64 = base64.b64encode(image_data).decode("utf-8")
            models_proxy.execute_kw(
                odoo_db, uid, odoo_pass, "product.template", "write",
                [tmpl_ids[:1], {"image_1920": img_b64}])
            subidas.append(sku)
        except Exception as e:
            errores.append(f"Error al subir imagen para `{sku}`: {e}")

    return subidas, errores
