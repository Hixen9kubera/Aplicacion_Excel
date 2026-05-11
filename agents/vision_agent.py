"""
Agente Vision — Fase B de la arquitectura multi-agente FERRAFORME.

Responsabilidades:
- Analizar imágenes de productos con Claude Haiku Vision
- Inferir categoría/atributo para productos sin imagen
- Estrategia dual:
    • Batch API de Anthropic para lotes > BATCH_THRESHOLD (sin rate limit, async)
    • Paralelo con bloques para lotes pequeños (síncrono, más rápido)
- Caché de resultados Vision por phash de imagen (ahorra costo en reórdenes)

Entrada : productos (list[dict]), imagenes ({row_0idx: {data, ext}})
Salida  : datos_vision ({idx: {subcategoria_cod, atributo_cod, titulo, ...}})
"""
from __future__ import annotations

import base64
import concurrent.futures
import json
import os
import pickle
import time
from pathlib import Path
from typing import TypedDict

import anthropic
from langchain_anthropic import ChatAnthropic
from langchain_core.messages import HumanMessage

from agents.utils import _phash_imagen

# Usar Batch API cuando hay más de este número de productos con imagen
BATCH_THRESHOLD = 40


def _comprimir_imagen(data: bytes, max_px: int = 1024, quality: int = 85) -> tuple[bytes, str]:
    """Redimensiona y comprime una imagen a JPEG. Devuelve (bytes, 'image/jpeg')."""
    try:
        from PIL import Image
        import io
        img = Image.open(io.BytesIO(data))
        if img.mode not in ("RGB", "L"):
            img = img.convert("RGB")
        img.thumbnail((max_px, max_px), Image.LANCZOS)
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=quality, optimize=True)
        return buf.getvalue(), "image/jpeg"
    except Exception:
        return data, "image/jpeg"

# Rate limit síncrono (tokens/min ÷ tokens/llamada × 0.8)
_RATE_LIMIT   = 30_000
_TOKENS_CALL  = 2_000
_BLOQUE_SIZE  = max(1, int(_RATE_LIMIT / _TOKENS_CALL * 0.8))   # = 12

# Caché de resultados Vision
_VISION_CACHE_PATH = Path(__file__).parent.parent / "vision_cache.pkl"
_CAMPOS_CACHE      = ("subcategoria_cod", "atributo_cod", "atributo_desc",
                      "titulo", "descripcion", "categoria")
_UMBRAL_CACHE      = 10   # distancia phash máxima para considerar "misma imagen"


# ── Prompt y parseo ───────────────────────────────────────────────────────────

_SUBCAT_OPCIONES = (
    "Muebles/Hogar→MUE,MES,SIL,CAM,EST,ORG,COM,ESCR,BAÑ,JAR,TV,COC,DEC,ILUM,TEX | "
    "Bebés→BEB,CUNA,PAS,CORR,ALIM,HIG,ROBB,SEG,JUG | "
    "Juguetes→JUGU,MUN,PEL,VEH,CONS,JUEG,CART,ELEC,CAS,MONT,DEPO,EDU | "
    "Moda→ROP,CALZ,ACC | Tecnología→TEC,CEL | Herramientas/Oficina→HERR,OFI | "
    "Mascotas→MASC | Varios→LIB,ART,DEP,VIA,VAR"
)
_ATTR_OPCIONES = (
    "Colores→NEG,BLN,GRI,ROJ,AZL,VER,AMA,ROS,NAR,MOR,CAF,BEI,MUL,PLA,DOR | "
    "Tallas→XS,S,M,L,XL,UNI | Materiales→MAD,MET,TEL,CUE | Otros→EST,PRE,ECO,INF,ADU"
)


def _build_vision_prompt(contexto: dict | None = None) -> str:
    ctx_lines = []
    if contexto:
        if contexto.get("nombre"):
            ctx_lines.append(f"- Nombre del producto: {contexto['nombre']}")
        if contexto.get("material"):
            ctx_lines.append(f"- Material: {contexto['material']}")
        if contexto.get("uso"):
            ctx_lines.append(f"- Uso / aplicación: {contexto['uso']}")
        if contexto.get("largo_cm") and contexto.get("ancho_cm") and contexto.get("alto_cm"):
            ctx_lines.append(
                f"- Dimensiones: {contexto['largo_cm']} × {contexto['ancho_cm']} × {contexto['alto_cm']} cm"
            )

    ctx_bloque = ""
    if ctx_lines:
        ctx_bloque = (
            "\n\nDATOS DEL PRODUCTO (combínalos con la imagen para una clasificación precisa):\n"
            + "\n".join(ctx_lines)
            + "\n\nIMPORTANTE para subcategoria_cod: analiza TANTO la imagen como el nombre juntos. "
            "El nombre es clave para entender qué tipo de objeto ES (silla, mesa, andadera, etc.) "
            "mientras que la imagen confirma su forma y características visuales. "
            "Nunca uses ROP/CALZ/ACC para muebles, juguetes o accesorios de bebé solo porque tienen tela o correas."
        )

    return (
        f"Analiza esta imagen de un producto de importación y devuelve ÚNICAMENTE este JSON.\n"
        f"TODOS los campos son obligatorios y nunca pueden quedar vacíos.{ctx_bloque}\n\n"
        "{{\n"
        '  "titulo": "nombre comercial corto en español (máx 60 caracteres)",\n'
        '  "descripcion": "descripción en español destacando material, uso y características principales (máx 120 caracteres)",\n'
        '  "categoria": "categoría principal en español",\n'
        f'  "subcategoria_cod": "SOLO uno de estos códigos — basa tu elección en el nombre del producto Y en la imagen: {_SUBCAT_OPCIONES}",\n'
        f'  "atributo_cod": "SOLO uno de estos códigos. PRIORIDAD: si el producto tiene un color predominante visible usa el código de color (NEG=negro, BLN=blanco, GRI=gris, ROJ=rojo, AZL=azul, VER=verde, AMA=amarillo, ROS=rosa/pink, NAR=naranja, MOR=morado, CAF=café/beige oscuro, BEI=beige, MUL=multicolor, PLA=plateado, DOR=dorado). Usa EST solo si el color no es relevante o el producto es neutro. Opciones completas: {_ATTR_OPCIONES}",\n'
        '  "atributo_desc": "color principal del producto o su característica diferenciadora (ej: \'Rosa\', \'Negro\', \'Azul marino\', \'Metal plateado\')"\n'
        "}}\n"
        "Solo JSON, sin texto adicional."
    )


def _parse_vision_response(texto: str, contexto: dict | None = None) -> dict:
    """Parsea la respuesta JSON de Vision y aplica fallbacks."""
    nombre_ctx = (contexto or {}).get("nombre", "Producto")
    try:
        if texto.startswith("```"):
            partes = texto.split("```")
            texto  = partes[1].lstrip("json").strip() if len(partes) > 1 else texto
        datos = json.loads(texto)
        datos["titulo"]           = datos.get("titulo")           or nombre_ctx
        datos["descripcion"]      = datos.get("descripcion")      or nombre_ctx
        datos["categoria"]        = datos.get("categoria")        or "Varios"
        datos["subcategoria_cod"] = datos.get("subcategoria_cod") or "VAR"
        datos["atributo_cod"]     = datos.get("atributo_cod")     or "EST"
        datos["atributo_desc"]    = datos.get("atributo_desc")    or "Estándar"
        datos["_error"]           = None
        return datos
    except Exception as e:
        return {
            "titulo": nombre_ctx, "descripcion": nombre_ctx,
            "categoria": "Varios", "subcategoria_cod": "VAR",
            "atributo_cod": "EST", "atributo_desc": "Estándar",
            "_error": str(e),
        }


# ── Caché de resultados Vision ────────────────────────────────────────────────

def _cargar_cache_vision() -> dict:
    """
    Carga el caché de Vision desde Supabase (persistente) o archivo local.
    Devuelve {phash_hex: {subcategoria_cod, atributo_cod, ...}}.
    """
    # Supabase (preferido — sobrevive reinicios del servidor)
    try:
        url = os.environ.get("SUPABASE_URL", "")
        key = os.environ.get("SUPABASE_KEY", "")
        if url and key:
            from supabase import create_client
            sb   = create_client(url, key)
            resp = sb.table("vision_cache").select(
                "phash_hex,subcategoria_cod,atributo_cod,atributo_desc,titulo,descripcion,categoria"
            ).execute()
            if resp.data:
                return {r["phash_hex"]: r for r in resp.data}
    except Exception:
        pass

    # Fallback local
    if _VISION_CACHE_PATH.exists():
        try:
            with open(_VISION_CACHE_PATH, "rb") as f:
                return pickle.load(f)
        except Exception:
            pass

    return {}


def _guardar_cache_vision(nuevos: list[dict]) -> None:
    """
    Persiste nuevos resultados Vision en Supabase y en archivo local.
    nuevos: [{phash_hex, subcategoria_cod, atributo_cod, atributo_desc, titulo, descripcion, categoria}]
    """
    if not nuevos:
        return

    # Supabase
    try:
        url = os.environ.get("SUPABASE_URL", "")
        key = os.environ.get("SUPABASE_KEY", "")
        if url and key:
            from supabase import create_client
            sb = create_client(url, key)
            for i in range(0, len(nuevos), 100):
                sb.table("vision_cache").upsert(nuevos[i:i+100]).execute()
    except Exception:
        pass

    # Local
    try:
        cache: dict = {}
        if _VISION_CACHE_PATH.exists():
            with open(_VISION_CACHE_PATH, "rb") as f:
                cache = pickle.load(f)
        for item in nuevos:
            cache[item["phash_hex"]] = item
        with open(_VISION_CACHE_PATH, "wb") as f:
            pickle.dump(cache, f)
    except Exception:
        pass


def _buscar_en_cache(phash_img, cache: dict, umbral: int = _UMBRAL_CACHE) -> dict | None:
    """
    Busca en el caché de Vision por similitud de phash.
    Devuelve el resultado más cercano si dist ≤ umbral, o None.
    """
    if phash_img is None or not cache:
        return None
    try:
        import imagehash
        mejor_dist = umbral + 1
        mejor      = None
        for phash_hex, datos in cache.items():
            ph   = imagehash.hex_to_hash(phash_hex)
            dist = phash_img - ph
            if dist <= umbral and dist < mejor_dist:
                mejor_dist = dist
                mejor      = datos
        return mejor
    except Exception:
        return None


# ── Modo síncrono (individual, con retry) ────────────────────────────────────

def analizar_imagen_claude(image_data: bytes, ext: str, contexto: dict | None = None) -> dict:
    """
    Analiza una imagen de producto con Claude Haiku Vision.
    Incluye retry con backoff exponencial en caso de rate limit.
    """
    media_map = {
        "jpg": "image/jpeg", "jpeg": "image/jpeg",
        "png": "image/png",  "gif": "image/gif",
        "webp": "image/webp", "bmp": "image/png",
    }
    media_type = media_map.get(ext, "image/jpeg")
    prompt     = _build_vision_prompt(contexto)

    try:
        client  = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY", ""))
        image_data, media_type = _comprimir_imagen(image_data)
        img_b64 = base64.b64encode(image_data).decode("utf-8")
        resp    = None
        for _intento in range(4):
            try:
                resp = client.messages.create(
                    model="claude-haiku-4-5-20251001",
                    max_tokens=512,
                    messages=[{"role": "user", "content": [
                        {"type": "image", "source": {
                            "type": "base64", "media_type": media_type, "data": img_b64,
                        }},
                        {"type": "text", "text": prompt},
                    ]}],
                )
                break
            except anthropic.RateLimitError:
                if _intento == 3:
                    raise
                time.sleep(15 * (2 ** _intento))   # 15s, 30s, 60s

        return _parse_vision_response(resp.content[0].text.strip(), contexto)
    except Exception as e:
        nombre_ctx = (contexto or {}).get("nombre", "Producto")
        return {
            "titulo": nombre_ctx, "descripcion": nombre_ctx,
            "categoria": "Varios", "subcategoria_cod": "VAR",
            "atributo_cod": "EST", "atributo_desc": "Estándar",
            "_error": str(e),
        }


# ── Inferencia sin imagen ─────────────────────────────────────────────────────

def _inferir_categoria_manual(nombre: str, material: str, uso: str) -> dict:
    """Infiere subcategoria_cod y atributo_cod de un producto sin imagen usando Claude Haiku."""
    try:
        llm    = ChatAnthropic(model="claude-haiku-4-5-20251001", max_tokens=100)
        prompt = (
            f"Clasifica este producto para generar un SKU.\n"
            f"Nombre: {nombre}\nMaterial: {material or '—'}\nUso: {uso or '—'}\n\n"
            "Devuelve SOLO un JSON válido:\n"
            "{\"subcategoria_cod\": \"XXX\", \"atributo_cod\": \"YYY\"}\n"
            "subcategoria_cod: 2-4 letras mayúsculas (tipo de producto, ej: MES, SIL, CAJ, BOL)\n"
            "atributo_cod: 2-4 letras mayúsculas (color/material, ej: NGR, BLC, NAT, GEN)"
        )
        resp    = llm.invoke([HumanMessage(content=prompt)])
        content = resp.content
        if isinstance(content, list):
            content = " ".join(b.get("text", "") for b in content if isinstance(b, dict))
        return json.loads(content.strip())
    except Exception:
        return {"subcategoria_cod": "VAR", "atributo_cod": "EST"}


# ── Modo Batch API (sin rate limit) ──────────────────────────────────────────

def _agente_vision_batch(
    productos: list[dict],
    imagenes: dict[int, dict],
    on_progress: "callable | None" = None,
) -> dict[int, dict]:
    """
    Procesa imágenes usando la Batch API de Anthropic.
    Sin límite de tokens por minuto — idóneo para 50+ productos.
    Consulta caché por phash antes de agregar al batch.

    on_progress(n_done, n_total) es llamado cada vez que se obtienen resultados parciales.
    Devuelve datos_vision {idx: {...}}.
    """
    media_map = {
        "jpg": "image/jpeg", "jpeg": "image/jpeg",
        "png": "image/png",  "gif": "image/gif",
        "webp": "image/webp", "bmp": "image/png",
    }
    client       = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY", ""))
    cache_vision = _cargar_cache_vision()

    # ── Preparar requests — separar hits de caché vs pendientes de Vision ────
    batch_requests:  list[dict]         = []
    idx_sin_imagen:  list[int]          = []
    phashes_batch:   dict[str, object]  = {}   # custom_id → phash_img
    datos_vision:    dict[int, dict]    = {}
    phash_run:       dict[str, dict]    = {}   # phash_hex → resultado, dedup intra-run

    for i, prod in enumerate(productos):
        row_num = prod.get("fila_excel_0idx", i + 1)
        img     = imagenes.get(row_num)
        if img:
            img_phash = _phash_imagen(img["data"])
            cached    = _buscar_en_cache(img_phash, cache_vision)
            if cached:
                # Hit de caché persistente → resultado gratis
                datos_vision[i] = {k: cached.get(k, "") for k in _CAMPOS_CACHE}
                datos_vision[i]["_error"] = None
                continue

            # Deduplicar dentro del mismo run (misma imagen, distinto producto)
            phash_hex = str(img_phash) if img_phash is not None else None
            if phash_hex and phash_hex in phash_run:
                datos_vision[i] = phash_run[phash_hex].copy()
                continue

            # Miss de caché → comprimir y agregar al batch
            img_data, media_type = _comprimir_imagen(img["data"])
            img_b64 = base64.b64encode(img_data).decode("utf-8")
            ctx        = {
                "nombre":   prod.get("nombre"),
                "material": prod.get("material"),
                "uso":      prod.get("uso"),
            }
            prompt = _build_vision_prompt(ctx)
            batch_requests.append({
                "custom_id": str(i),
                "params": {
                    "model":      "claude-haiku-4-5-20251001",
                    "max_tokens": 512,
                    "messages": [{"role": "user", "content": [
                        {"type": "image", "source": {
                            "type": "base64", "media_type": media_type, "data": img_b64,
                        }},
                        {"type": "text", "text": prompt},
                    ]}],
                },
            })
            if img_phash is not None:
                phashes_batch[str(i)] = img_phash
        else:
            idx_sin_imagen.append(i)

    # ── Procesar productos SIN imagen con Claude Haiku (texto) ───────────────
    for i in idx_sin_imagen:
        prod  = productos[i]
        datos = _inferir_categoria_manual(
            prod.get("nombre", ""), prod.get("material", ""), prod.get("uso", "")
        )
        datos.setdefault("atributo_desc", datos.get("atributo_cod", ""))
        datos_vision[i] = datos

    if not batch_requests:
        return datos_vision

    # ── Enviar en chunks de 50 para no superar el límite de 256 MB ───────────
    _CHUNK = 50
    chunks  = [batch_requests[i:i+_CHUNK] for i in range(0, len(batch_requests), _CHUNK)]
    n_total = len(batch_requests)
    n_done  = 0

    nuevos_cache: list[dict] = []

    for chunk in chunks:
        batch = client.messages.batches.create(requests=chunk)

        poll_interval = 10
        while batch.processing_status == "in_progress":
            time.sleep(poll_interval)
            batch   = client.messages.batches.retrieve(batch.id)
            n_done_ = batch.request_counts.succeeded + batch.request_counts.errored
            if on_progress:
                on_progress(n_done + n_done_, n_total)

        for result in client.messages.batches.results(batch.id):
            idx  = int(result.custom_id)
            prod = productos[idx]
            ctx  = {"nombre": prod.get("nombre")}
            if result.result.type == "succeeded":
                texto = result.result.message.content[0].text.strip()
                datos = _parse_vision_response(texto, ctx)
            else:
                datos = _parse_vision_response("", ctx)
            datos_vision[idx] = datos

            if not datos.get("_error") and (ph := phashes_batch.get(result.custom_id)):
                entry = {"phash_hex": str(ph), **{k: datos.get(k, "") for k in _CAMPOS_CACHE}}
                nuevos_cache.append(entry)
                phash_run[str(ph)] = {k: datos.get(k, "") for k in _CAMPOS_CACHE}

        n_done += len(chunk)

    _guardar_cache_vision(nuevos_cache)
    return datos_vision


# ── Modo paralelo en bloques (síncrono, para lotes pequeños) ─────────────────

def _agente_vision_paralelo(
    productos: list[dict],
    imagenes: dict[int, dict],
) -> dict[int, dict]:
    """
    Procesa imágenes en bloques paralelos con pausa entre bloques para respetar
    el rate limit de 30,000 tokens/minuto de Anthropic.
    Consulta caché por phash antes de llamar a Vision.
    """
    cache_vision = _cargar_cache_vision()

    def _procesar(args):
        i, prod = args
        row_num = prod.get("fila_excel_0idx", i + 1)
        img     = imagenes.get(row_num)
        if img:
            img_phash = _phash_imagen(img["data"])
            cached    = _buscar_en_cache(img_phash, cache_vision)
            if cached:
                datos = {k: cached.get(k, "") for k in _CAMPOS_CACHE}
                datos["_error"] = None
                return i, datos, None   # hit → sin phash que guardar
            ctx   = {"nombre": prod.get("nombre"), "material": prod.get("material"),
                     "uso": prod.get("uso")}
            datos = analizar_imagen_claude(img["data"], img["ext"], ctx)
            return i, datos, img_phash  # miss → pasar phash para guardarlo
        else:
            datos = _inferir_categoria_manual(
                prod.get("nombre", ""), prod.get("material", ""), prod.get("uso", "")
            )
            datos.setdefault("atributo_desc", datos.get("atributo_cod", ""))
            return i, datos, None

    # Pre-calcular phashes para deduplicar intra-run
    phash_run:    dict[str, dict] = {}   # phash_hex → resultado ya obtenido
    phash_by_idx: dict[int, str]  = {}   # idx → phash_hex
    for i, prod in enumerate(productos):
        row_num = prod.get("fila_excel_0idx", i + 1)
        img = imagenes.get(row_num)
        if img:
            ph = _phash_imagen(img["data"])
            if ph is not None:
                phash_by_idx[i] = str(ph)

    datos_vision:  dict[int, dict] = {}
    nuevos_cache:  list[dict]      = []
    bloques = [list(enumerate(productos))[j: j + _BLOQUE_SIZE]
               for j in range(0, len(productos), _BLOQUE_SIZE)]

    for b_idx, bloque in enumerate(bloques):
        # Filtrar los que ya tienen resultado por dedup intra-run
        bloque_filtrado = []
        for idx, prod in bloque:
            ph = phash_by_idx.get(idx)
            if ph and ph in phash_run:
                datos_vision[idx] = phash_run[ph].copy()
            else:
                bloque_filtrado.append((idx, prod))

        if bloque_filtrado:
            with concurrent.futures.ThreadPoolExecutor(max_workers=_BLOQUE_SIZE) as ex:
                for i, datos, img_phash in ex.map(_procesar, bloque_filtrado):
                    datos_vision[i] = datos
                    if img_phash and not datos.get("_error"):
                        entry = {"phash_hex": str(img_phash),
                                 **{k: datos.get(k, "") for k in _CAMPOS_CACHE}}
                        nuevos_cache.append(entry)
                        phash_run[str(img_phash)] = {k: datos.get(k, "") for k in _CAMPOS_CACHE}

        if b_idx < len(bloques) - 1:
            time.sleep(62)

    _guardar_cache_vision(nuevos_cache)
    return datos_vision


# ── Nodo LangGraph ─────────────────────────────────────────────────────────────

def agente_vision(
    productos: list[dict],
    imagenes: dict[int, dict],
    forzar_batch: bool = False,
    on_progress: "callable | None" = None,
) -> dict[int, dict]:
    """
    Nodo LangGraph — selecciona estrategia Vision según el tamaño del lote.

    • lote > BATCH_THRESHOLD o forzar_batch=True → Batch API (sin rate limit)
    • lote <= BATCH_THRESHOLD                    → paralelo en bloques (síncrono)

    En ambos modos consulta el caché Vision (phash) antes de llamar a la API.
    Devuelve datos_vision {idx: {subcategoria_cod, atributo_cod, titulo, descripcion, ...}}.
    """
    n_con_imagen = sum(
        1 for i, prod in enumerate(productos)
        if imagenes.get(prod.get("fila_excel_0idx", i + 1))
    )

    if forzar_batch or n_con_imagen > BATCH_THRESHOLD:
        return _agente_vision_batch(productos, imagenes, on_progress=on_progress)
    else:
        return _agente_vision_paralelo(productos, imagenes)
