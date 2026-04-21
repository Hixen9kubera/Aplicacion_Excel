"""
app_v2.py — Agente conversacional FERRAFORME v2
Uso: streamlit run app_v2.py

Mejoras vs v1:
- Dudas relevantes se resuelven con botones/radio ANTES de generar el Excel
- Dudas menores se muestran como avisos informativos
- Prioridad automática a columnas en inglés cuando hay duplicados
- Migrado a LangChain + LangSmith (desactivado hasta producto final)
"""

import io
import os
import re
import json
import math
import time
import pickle
import zipfile
import base64
import xml.etree.ElementTree as ET
import xmlrpc.client
from datetime import datetime, timedelta
from pathlib import Path
from typing import List

import difflib
import concurrent.futures
import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment, GradientFill
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
try:
    from PIL import Image as PILImage
    _PILLOW_OK = True
except ImportError:
    _PILLOW_OK = False

from langchain_anthropic import ChatAnthropic
from langchain_core.tools import tool
from langchain_core.messages import HumanMessage, AIMessage, SystemMessage, ToolMessage

# ── Cargar .env ────────────────────────────────────────────────────────────────
_env_path = Path(__file__).parent / ".env"
if _env_path.exists():
    for _line in _env_path.read_text(encoding="utf-8").splitlines():
        _line = _line.strip()
        if _line and not _line.startswith("#") and "=" in _line:
            _k, _v = _line.split("=", 1)
            os.environ.setdefault(_k.strip(), _v.strip())

# LangSmith activado — cambiar a "false" para desactivar
os.environ["LANGCHAIN_TRACING_V2"] = "true"

# ── Configuración fija ─────────────────────────────────────────────────────────
TEMPLATE_PATH = Path(__file__).parent / "FORMULA FERRAFORME PRODUCTOS .xlsx"
EMPRESA       = "FERRAFORME MS"

# ── Códigos válidos de SKU (base_sku_fixed) ────────────────────────────────────
SUBCATEGORIAS = {
    # MUEBLES - HOGAR
    "MUE": "Muebles Hogar",   "MES": "Mesa",          "SIL": "Silla",
    "CAM": "Cama",            "EST": "Estantería",     "ORG": "Organizador",
    "COM": "Comedor",         "ESCR": "Escritorio",    "BAÑ": "Mueble Baño",
    "JAR": "Mueble Jardín",   "TV": "Mueble TV",       "COC": "Mueble Cocina",
    "DEC": "Decoración Hogar","ILUM": "Iluminación",   "TEX": "Textiles Hogar",
    # BEBÉS - INFANTIL
    "BEB": "Artículos Bebé",  "CUNA": "Cuna y Catre",  "PAS": "Paseo Bebé",
    "CORR": "Corral y Andadera","ALIM": "Alimentación Bebé","HIG": "Higiene Bebé",
    "ROBB": "Ropa Bebé",      "SEG": "Seguridad Bebé", "JUG": "Juguetes Bebé",
    # JUGUETES - JUEGOS
    "JUGU": "Juguetes",       "MUN": "Muñeca y Figura","PEL": "Peluche",
    "VEH": "Vehículo Juguete","CONS": "Juego Construcción","JUEG": "Juego de Mesa",
    "CART": "Cartas y Mazo",  "ELEC": "Juguete Electrónico","CAS": "Casa de Juguete",
    "MONT": "Juguete Montable","DEPO": "Juguete Deportivo","EDU": "Juguete Educativo",
    # MODA Y ROPA
    "ROP": "Ropa",            "CALZ": "Calzado",       "ACC": "Accesorio Moda",
    # TECNOLOGÍA
    "TEC": "Electrónica",     "CEL": "Celular y Accesorios",
    # HERRAMIENTAS Y OFICINA
    "HERR": "Herramienta",    "OFI": "Oficina",
    # MASCOTAS
    "MASC": "Mascota",
    # VARIOS
    "LIB": "Libro y Papelería","ART": "Arte y Manualidad","DEP": "Deporte y Fitness",
    "VIA": "Viaje y Equipaje", "VAR": "Varios",
}
ATRIBUTOS = {
    # COLORES
    "NEG": "Negro",   "BLN": "Blanco",  "GRI": "Gris",    "ROJ": "Rojo",
    "AZL": "Azul",    "VER": "Verde",   "AMA": "Amarillo","ROS": "Rosa",
    "NAR": "Naranja", "MOR": "Morado",  "CAF": "Café",    "BEI": "Beige",
    "MUL": "Multicolor","PLA": "Plateado","DOR": "Dorado",
    # TALLAS
    "XS": "Extra Small","S": "Small",  "M": "Medium",    "L": "Large",
    "XL": "Extra Large","UNI": "Talla Única",
    # MATERIALES
    "MAD": "Madera",  "MET": "Metal",   "TEL": "Tela",    "CUE": "Cuero",
    # OTROS
    "EST": "Estándar","PRE": "Premium", "ECO": "Económico","INF": "Infantil",
    "ADU": "Adulto",
}
_SUBCAT_DEFAULT = "VAR"
_ATTR_DEFAULT   = "EST"


def extraer_contenedor(filename: str) -> str:
    """
    Extrae el número de contenedor del nombre del archivo.
    Busca el patrón ISO 6346: 4 letras seguidas de 6-7 dígitos (ej. TCKU1234567).
    Si no lo encuentra, devuelve el stem del archivo sin extensión.
    """
    nombre = Path(filename).stem
    # Patrón ISO 6346: 3-4 letras, separador opcional, 5-8 dígitos
    # Ej: TCKU1234567, MSCU 1234567, ABCD-1234567
    match = re.search(r'([A-Z]{3,4})[\s_-]?(\d{5,8})', nombre.upper())
    if match:
        return match.group(1) + match.group(2)
    return nombre


# ── Herramientas del agente ────────────────────────────────────────────────────
# ── System prompt ──────────────────────────────────────────────────────────────
def build_system_prompt(analisis, productos, tipo_cambio, contenedor, respuestas_dudas=None):
    n_prods = len(productos) if productos else 0
    prompt = f"""Eres el agente FERRAFORME, especializado en analizar packing lists de importación \
y generar plantillas de productos Excel para la empresa FERRAFORME MS.

CONFIGURACIÓN ACTUAL:
- Tipo de cambio: ${tipo_cambio:.2f} MXN/USD
- Contenedor: {contenedor}
- Empresa: FERRAFORME MS

ARCHIVO YA PROCESADO:
El packing list ya fue subido y procesado por completo. Los {n_prods} productos están disponibles \
en la sección PRODUCTOS de este mismo mensaje. NUNCA le pidas al usuario que vuelva a subir el \
archivo ni digas que no tienes acceso a él. Todo el contexto necesario está aquí.

HERRAMIENTA DISPONIBLE — generar_excel:
Llámala cuando el usuario confirme proceder o pida generar el Excel.
NO necesitas pasar ningún argumento — los productos ya están almacenados en sesión.

REGLAS DE COMPORTAMIENTO:
1. Las dudas relevantes ya fueron resueltas por el usuario antes de llegar aquí.
2. Cuando el usuario confirme que puede proceder, llama a generar_excel.
3. Si el usuario pide una corrección, modifica los datos y vuelve a llamar generar_excel.
4. Aprende de las correcciones del usuario en esta sesión.
5. Responde siempre en español, de forma concisa y profesional.
6. En tus respuestas y resúmenes NO menciones el tipo de cambio aplicado.
7. Si el usuario pide datos de un producto específico, búscalos en la lista PRODUCTOS de abajo."""

    if respuestas_dudas:
        prompt += f"\n\n--- RESPUESTAS DEL USUARIO A LAS DUDAS ---\n{json.dumps(respuestas_dudas, ensure_ascii=False, indent=2)}"

    if analisis:
        prompt += f"\n\n--- ANÁLISIS DEL PACKING LIST ---\n{json.dumps(analisis, ensure_ascii=False, indent=2)}"

    if productos:
        prompt += f"\n\n--- PRODUCTOS ({len(productos)}) ---\n{json.dumps(productos, ensure_ascii=False, indent=2)}"

    return prompt


def _build_system_fase(fase: str, productos: list, analisis: dict | None = None,
                        extra: str = "") -> str:
    """
    System prompt para las fases intermedias (dudas, duplicados, conflictos).
    Incluye datos completos del packing list para que el agente pueda responder
    cualquier pregunta sobre los productos.
    Al final de cada respuesta debe preguntar si el usuario tiene otra duda
    o si ya puede continuar con el proceso.
    """
    n = len(productos) if productos else 0
    prompt = (
        f"Eres el agente FERRAFORME. En este momento el usuario está en la fase: **{fase}**.\n\n"
        "INSTRUCCIONES:\n"
        "1. Responde la pregunta o aclaración del usuario de forma concisa y en español.\n"
        "2. Si el usuario menciona una corrección, confírmale que la tomará en cuenta.\n"
        "3. Al final de CADA respuesta tuya, pregunta siempre:\n"
        '   "¿Tienes alguna otra duda o ya puedo continuar con el proceso?"\n'
        "4. Tienes acceso completo a todos los datos del packing list — NUNCA digas que no "
        "puedes ver el archivo o los productos.\n"
        "5. Si el usuario dice que ya puede continuar, responde brevemente con un mensaje "
        "de confirmación positivo (ej. 'Perfecto, ¡continuamos!'). "
        "El botón para avanzar sigue visible en la pantalla.\n"
    )
    if extra:
        prompt += f"\nCONTEXTO ADICIONAL:\n{extra}\n"
    if analisis:
        prompt += f"\nANÁLISIS DEL PACKING LIST:\n{json.dumps(analisis, ensure_ascii=False, indent=2)}\n"
    if productos:
        # Limitar a 30 productos para no saturar el contexto
        muestra = productos[:30]
        prompt += f"\nPRODUCTOS ({n} en total, mostrando primeros {len(muestra)}):\n"
        prompt += json.dumps(muestra, ensure_ascii=False, indent=2)
    return prompt


# ── Análisis de encabezados ────────────────────────────────────────────────────
def analizar_encabezados(file_bytes: bytes) -> dict:
    """Claude analiza los encabezados y devuelve mapeo + dudas estructuradas.
    Detecta automáticamente en qué fila comienza la tabla de productos,
    independientemente del formato (algunos packing lists tienen cabeceras
    de empresa en las primeras filas antes de la tabla real).
    """
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    filas_raw = []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i >= 35:
            break
        filas_raw.append([str(v) if v is not None else "" for v in row])

    # Enviar todas las filas numeradas para que Claude identifique la fila de headers
    filas_numeradas = [
        {"fila": i + 1, "valores": fila}
        for i, fila in enumerate(filas_raw)
    ]

    prompt = f"""Analiza este packing list de importación.
IMPORTANTE: Algunos packing lists tienen información de empresa/embarque en las primeras filas
antes de la tabla de productos. Identifica exactamente en qué fila (1-indexed) están los
encabezados de columna de la tabla de productos (la fila que dice cosas como "Description",
"Qty", "CBM", "Price", 数量, 品名, etc.).
Los encabezados pueden estar en cualquier idioma (chino, inglés, español, etc.).

Campos a identificar (índice base 0):
nombre_producto, nombre_producto_alt, id_guia, cajas_master, piezas_x_caja, piezas_total, cbm_por_pieza, cbm_master_carton, cbm_total_sku, precio_usd, largo_cm, ancho_cm, alto_cm, peso_kg, volumen_por_pieza, material, uso

IMPORTANTE para columnas CBM — usa los valores numéricos de las filas de muestra para inferir correctamente:

Definiciones:
- cbm_por_pieza:     CBM de UNA pieza individual         → valor típico 0.0001–0.05
- cbm_master_carton: CBM de UNA caja/master carton        → valor típico 0.02–0.5
- cbm_total_sku:     CBM TOTAL de todas las unidades SKU  → valor típico 0.1–20

Relaciones matemáticas para verificar con los datos de muestra:
  cbm_por_pieza × piezas_total  ≈  cbm_total_sku   (si cbm_por_pieza es por pieza)
  cbm_master_carton × cajas     ≈  cbm_total_sku   (si cbm_master_carton es por caja)
  cbm_master_carton / piezas_x_caja  ≈  cbm_por_pieza

Regla clave para columnas ambiguas (header solo dice "CBM" o "M³" sin especificar):
  1. Toma el valor de muestra de esa columna (ej: 0.085) y los valores de cajas y piezas_total.
  2. Calcula: valor × cajas  vs  valor × piezas_total
  3. El que se acerque al cbm_total (si existe) indica el tipo correcto.
  4. Si no hay cbm_total, usa la magnitud: si valor × piezas_total > 50 m³ es improbable
     para un contenedor estándar (máx ~76 m³) → probablemente es por caja, no por pieza.

Si hay más de una columna CBM, mapéalas todas. Nunca mapees la misma columna en dos campos distintos.

IMPORTANTE para nombre_producto y nombre_producto_alt:
- Si hay DOS columnas de nombre (ej: una en inglés y otra en chino), mapea la principal en nombre_producto y la alternativa en nombre_producto_alt.
- Si solo hay una columna de nombre, deja nombre_producto_alt con confianza "no_encontrado".

DATOS (filas del Excel con su número de fila real):
{json.dumps(filas_numeradas, ensure_ascii=False, indent=2)}

Responde ÚNICAMENTE con este JSON (sin texto adicional):
{{
  "idioma_detectado": "...",
  "fila_encabezado": 1,
  "numero_contenedor": "",
  "columnas": {{
    "nombre_producto":     {{"indice": 0,    "encabezado_original": "...", "valor_muestra": "...", "confianza": "alta",          "nota": ""}},
    "nombre_producto_alt": {{"indice": null,  "encabezado_original": "",    "valor_muestra": "",    "confianza": "no_encontrado", "nota": "nombre en otro idioma, si existe"}},
    "id_guia":             {{"indice": null,  "encabezado_original": "",    "valor_muestra": "",    "confianza": "no_encontrado", "nota": "número de guía o código de carga, opcional"}},
    "cajas_master":        {{"indice": null,  "encabezado_original": "",    "valor_muestra": "",    "confianza": "no_encontrado", "nota": "número de cajas/master cartons para este SKU"}},
    "piezas_x_caja":       {{"indice": null,  "encabezado_original": "",    "valor_muestra": "",    "confianza": "no_encontrado", "nota": "unidades por caja"}},
    "piezas_total":        {{"indice": null,  "encabezado_original": "",    "valor_muestra": "",    "confianza": "no_encontrado", "nota": "total piezas en el contenedor para ese SKU"}},
    "largo_cm":            {{"indice": null,  "encabezado_original": "",    "valor_muestra": "",    "confianza": "no_encontrado", "nota": ""}},
    "ancho_cm":            {{"indice": null,  "encabezado_original": "",    "valor_muestra": "",    "confianza": "no_encontrado", "nota": ""}},
    "alto_cm":             {{"indice": null,  "encabezado_original": "",    "valor_muestra": "",    "confianza": "no_encontrado", "nota": ""}},
    "peso_kg":             {{"indice": null,  "encabezado_original": "",    "valor_muestra": "",    "confianza": "no_encontrado", "nota": "peso en kg por pieza o por caja (especificar en nota)"}},
    "volumen_por_pieza":   {{"indice": null,  "encabezado_original": "",    "valor_muestra": "",    "confianza": "no_encontrado", "nota": "volumen por pieza en m³, si viene separado del CBM"}},
    "cbm_por_pieza":       {{"indice": null,  "encabezado_original": "",    "valor_muestra": "",    "confianza": "no_encontrado", "nota": "CBM de una pieza individual"}},
    "cbm_master_carton":   {{"indice": null,  "encabezado_original": "",    "valor_muestra": "",    "confianza": "no_encontrado", "nota": "CBM de una caja/master carton completa"}},
    "cbm_total_sku":       {{"indice": null,  "encabezado_original": "",    "valor_muestra": "",    "confianza": "no_encontrado", "nota": "CBM total de todas las unidades de este SKU"}},
    "precio_usd":     {{"indice": null,  "encabezado_original": "",    "valor_muestra": "",    "confianza": "no_encontrado", "nota": ""}},
    "material":       {{"indice": null,  "encabezado_original": "",    "valor_muestra": "",    "confianza": "no_encontrado", "nota": ""}},
    "uso":            {{"indice": null,  "encabezado_original": "",    "valor_muestra": "",    "confianza": "no_encontrado", "nota": ""}}
  }},
  "dudas_relevantes": [
    {{
      "id": 0,
      "tipo": "eleccion",
      "campo_afectado": "material",
      "descripcion": "Se encontraron dos columnas para 'material': col 11 en inglés y col 12 en chino. Parecen ser el mismo dato.",
      "pregunta": "¿Cuál columna de material deseas usar?",
      "opciones": ["Inglés — col 11 (MDF)", "Chino — col 12 (高密度板/MDF)"],
      "indices_opciones": [11, 12],
      "default": "Inglés — col 11 (MDF)"
    }}
  ],
  "dudas_menores": [
    "3 productos tienen precio USD vacío — quedarán en blanco en el Excel"
  ]
}}

NOTA sobre fila_encabezado:
- Pon el número de fila (1-indexed) donde están los encabezados de la tabla de productos.
- Si la fila 1 ya es el encabezado, pon 1.
- Si hay filas de información de empresa/embarque antes de la tabla, identifica la fila correcta.
- Los índices de columnas en "columnas" son SIEMPRE base 0 contando desde la izquierda del Excel,
  independientemente de en qué fila estén los encabezados.

NOTA sobre numero_contenedor:
- Busca en las filas previas a fila_encabezado (metadata del embarque) un número de contenedor.
- Patrones comunes: "Container No:", "CONT#", "Container:", "Contenedor:", "CTR", o un código
  alfanumérico que empiece con letras y tenga 11 caracteres (ej: TCKU1234567, MSCU9876543).
- Si encuentras más de uno, ponlos separados por coma.
- Si no encuentras ninguno, deja el campo vacío ("").

TIPOS DE DUDA RELEVANTE:
- "eleccion": el usuario debe elegir entre varias opciones (ej: qué columna usar cuando hay duplicados).
  Siempre pon como default la opción en INGLÉS si aplica.
- "confirmar": el usuario confirma o rechaza algo con Sí/No.
- "texto": el usuario necesita escribir un valor.

CRITERIO para dudas_relevantes: solo las que puedan causar un error o dato incorrecto en el Excel
(columna no encontrada, ambigüedad entre columnas, unidades dudosas, etc.).
CRITERIO para dudas_menores: observaciones informativas que no bloquean la generación
(celdas vacías ocasionales, campos opcionales ausentes, etc.).

Niveles de confianza: alta / media / baja / no_encontrado"""

    llm  = ChatAnthropic(model="claude-sonnet-4-6", max_tokens=4000)
    resp = llm.invoke([HumanMessage(content=prompt)])
    texto = resp.content.strip()
    if texto.startswith("```"):
        partes = texto.split("```")
        texto = partes[1].lstrip("json").strip() if len(partes) > 1 else texto
    if texto.endswith("```"):
        texto = texto[:-3].strip()
    return json.loads(texto)


def aplicar_respuestas(columnas: dict, dudas_relevantes: list, respuestas: dict) -> dict:
    """Actualiza el mapeo de columnas según las respuestas del usuario a las dudas."""
    columnas = dict(columnas)  # copia
    for duda in dudas_relevantes:
        duda_id  = str(duda["id"])
        respuesta = respuestas.get(duda_id)
        if duda["tipo"] == "eleccion" and respuesta and duda.get("indices_opciones"):
            try:
                idx_opcion   = duda["opciones"].index(respuesta)
                indice_nuevo = duda["indices_opciones"][idx_opcion]
                campo        = duda["campo_afectado"]
                if campo in columnas:
                    columnas[campo] = {**columnas[campo], "indice": indice_nuevo, "confianza": "alta"}
            except (ValueError, IndexError):
                pass
    return columnas


def leer_productos(file_bytes: bytes, columnas: dict, fila_encabezado: int = 1) -> tuple[list[dict], list[str]]:
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active

    def idx(campo):
        c = columnas.get(campo, {})
        if c.get("confianza") == "no_encontrado":
            return None
        i = c.get("indice")
        return int(i) if i is not None else None

    nombre_idx = idx("nombre_producto")
    if nombre_idx is None:
        return [], ["No se pudo identificar la columna de nombre de producto."]

    # ── Mapa de celdas combinadas: (row, col_1indexed) → valor de la celda superior ──
    merged_values: dict[tuple[int,int], object] = {}
    # Filas secundarias de celdas combinadas (no son la fila principal del rango)
    secondary_merged_rows: set[int] = set()
    for rng in ws.merged_cells.ranges:
        top_val = ws.cell(rng.min_row, rng.min_col).value
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                merged_values[(r, c)] = top_val
        for r in range(rng.min_row + 1, rng.max_row + 1):
            secondary_merged_rows.add(r)

    # Columnas CBM relevantes para detectar filas combinadas (1-indexed)
    _cbm_cols_1idx = {
        f: (idx(f) + 1)
        for f in ("cbm_master_carton", "cbm_total_sku", "cbm_por_pieza", "id_guia")
        if idx(f) is not None
    }

    def cell_val(row_num: int, col_0idx: int, row_vals: tuple):
        """Devuelve el valor de una celda, resolviendo celdas combinadas."""
        v = row_vals[col_0idx] if col_0idx < len(row_vals) else None
        if v is None:
            v = merged_values.get((row_num, col_0idx + 1))
        return v

    productos, advertencias = [], []
    # Los datos empiezan en la fila siguiente a la de encabezados
    primera_fila_datos = fila_encabezado + 1

    for row_num, row in enumerate(
        ws.iter_rows(min_row=primera_fila_datos, values_only=True),
        start=primera_fila_datos,
    ):
        if len(row) <= nombre_idx or cell_val(row_num, nombre_idx, row) is None:
            continue

        def val(campo):
            i = idx(campo)
            return cell_val(row_num, i, row) if i is not None else None

        prod = {
            "nombre":            val("nombre_producto"),
            "nombre_alt":        val("nombre_producto_alt"),
            "id_guia":           val("id_guia"),
            "cajas_master":      val("cajas_master"),
            "piezas_x_caja":     val("piezas_x_caja"),
            "piezas_total":      val("piezas_total"),
            "largo_cm":          val("largo_cm"),
            "ancho_cm":          val("ancho_cm"),
            "alto_cm":           val("alto_cm"),
            "cbm_por_pieza":     val("cbm_por_pieza"),
            "cbm_master_carton": val("cbm_master_carton"),
            "cbm_total_sku":     val("cbm_total_sku"),
            "precio_usd":        val("precio_usd"),
            "peso_kg":           val("peso_kg"),
            "volumen_por_pieza": val("volumen_por_pieza"),
            "material":          val("material") or "",
            "uso":               val("uso") or "",
            # Fila original del Excel (0-indexed) para emparejar con imágenes extraídas
            "fila_excel_0idx":   row_num - 1,
        }

        # Sanitizar campos numéricos — descartar strings no numéricos (ej. "Country of Origin: China")
        _CAMPOS_NUM = (
            "cajas_master", "piezas_x_caja", "piezas_total",
            "cbm_por_pieza", "cbm_master_carton", "cbm_total_sku",
            "precio_usd", "largo_cm", "ancho_cm", "alto_cm",
            "peso_kg", "volumen_por_pieza",
        )
        for _cn in _CAMPOS_NUM:
            _v = prod.get(_cn)
            if isinstance(_v, str):
                try:
                    prod[_cn] = float(_v.replace(",", ".").strip())
                except (ValueError, TypeError):
                    prod[_cn] = None

        nombre = prod["nombre"]
        # Suprimir advertencias para filas secundarias de celdas combinadas:
        # esos valores se calculan en corregir_cbm_inner
        if row_num not in secondary_merged_rows:
            if prod["precio_usd"] is None:
                advertencias.append(f"Fila {row_num} ({nombre}): precio USD vacío")
            if not any(prod.get(c) for c in ("cbm_por_pieza", "cbm_master_carton", "cbm_total_sku")):
                advertencias.append(f"Fila {row_num} ({nombre}): CBM vacío")

        productos.append(prod)

    # Filtrar filas de totales/notas que se cuelen como productos
    # (ej: "Total Amount:", "Country of Origin: China", "Declaration...")
    _PREFIJOS_NO_PROD = (
        "total", "country of origin", "declaration", "declare",
        "subtotal", "grand total", "注意", "备注", "合计",
    )
    def _es_fila_footer(prod: dict) -> bool:
        nom = str(prod.get("nombre") or "").strip().lower()
        return any(nom.startswith(p) for p in _PREFIJOS_NO_PROD)

    productos = [p for p in productos if not _es_fila_footer(p)]

    # Recortar hasta el último producto con datos clave completos
    # (cajas, piezas o CBM). Descarta filas de totales/notas al final.
    _CAMPOS_CLAVE = ("cajas_master", "piezas_total", "piezas_x_caja",
                     "cbm_por_pieza", "cbm_master_carton", "cbm_total_sku")
    ultimo_completo = -1
    for i, p in enumerate(productos):
        if any(p.get(c) for c in _CAMPOS_CLAVE):
            ultimo_completo = i
    if ultimo_completo >= 0:
        productos = productos[: ultimo_completo + 1]

    return productos, advertencias


def _safe_float(val, default: float = 0.0) -> float:
    """Convierte val a float de forma segura; devuelve default si no es numérico."""
    if val is None:
        return default
    try:
        return float(val)
    except (ValueError, TypeError):
        return default


def corregir_cbm(productos: list[dict], advertencias: list[str]) -> list[dict]:
    """
    Detecta si Gemini mapeó cbm_master_carton como cbm_por_pieza y corrige.

    Lógica: si cbm_total_sku existe, compara cuál de los dos ratios se acerca más a 1:
      - ratio_pieza = cbm_por_pieza × piezas_total / cbm_total_sku  (debe ≈ 1 si es por pieza)
      - ratio_caja  = cbm_por_pieza × cajas_master / cbm_total_sku  (debe ≈ 1 si es por carton)

    Si ratio_caja está mucho más cerca de 1 que ratio_pieza → cbm_por_pieza es realmente
    cbm_master_carton. Se corrige: cbm_master_carton = valor_original, cbm_por_pieza = /pzs_caja.
    """
    _TOL = 0.10   # tolerancia 10%

    for prod in productos:
        cbm_pz  = _safe_float(prod.get("cbm_por_pieza"))
        cbm_mc  = _safe_float(prod.get("cbm_master_carton"))
        cbm_tot = _safe_float(prod.get("cbm_total_sku"))
        cajas   = _safe_float(prod.get("cajas_master"))
        pzs_tot = _safe_float(prod.get("piezas_total"))
        pzs_cja = _safe_float(prod.get("piezas_x_caja"))

        # Solo podemos validar si tenemos cbm_por_pieza y cbm_total_sku
        if cbm_pz <= 0 or cbm_tot <= 0:
            continue

        ratio_pieza = (cbm_pz * pzs_tot) / cbm_tot if pzs_tot > 0 else None
        ratio_caja  = (cbm_pz * cajas)   / cbm_tot if cajas   > 0 else None

        err_pieza = abs(ratio_pieza - 1) if ratio_pieza is not None else 999
        err_caja  = abs(ratio_caja  - 1) if ratio_caja  is not None else 999

        # Si el ratio por caja es mucho más cercano a 1 que el ratio por pieza
        if err_caja < _TOL and err_pieza > _TOL:
            nombre = prod.get("nombre", "?")
            cbm_pz_corr = round(cbm_pz / pzs_cja, 6) if pzs_cja > 0 else None
            advertencias.append(
                f"CBM corregido en '{nombre}': "
                f"la columna CBM ({cbm_pz}) era por carton, no por pieza. "
                f"cbm_por_pieza = {cbm_pz_corr}"
            )
            # Reasignar correctamente
            prod["cbm_master_carton"] = cbm_pz
            # Si no había inner carton, asumir igual al master (un solo tipo de caja)
            if not prod.get("cbm_inner_carton"):
                prod["cbm_inner_carton"] = cbm_pz
            prod["cbm_por_pieza"] = cbm_pz_corr

    return productos


def corregir_cbm_inner(productos: list[dict], advertencias: list[str],
                        file_bytes: bytes, columnas: dict) -> list[dict]:
    """
    Detecta productos que comparten un master carton (celdas combinadas en Excel
    O mismo id_guia) y calcula cbm_inner_carton y cbm_por_pieza correctamente.

    Lógica:
      cbm_inner     = cbm_master / nº productos en el grupo
      cbm_por_pieza = cbm_inner / piezas_x_caja
    """
    if not productos:
        return productos

    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active

    def _idx(campo):
        c = columnas.get(campo, {})
        if c.get("confianza") == "no_encontrado":
            return None
        i = c.get("indice")
        return (int(i) + 1) if i is not None else None  # 1-indexed

    cbm_mc_col = _idx("cbm_master_carton") or _idx("cbm_total_sku") or _idx("cbm_por_pieza")

    # ── Método 1: celdas combinadas en la columna CBM ─────────────────────────
    grupos_merged: dict[int, list[int]] = {}   # fila_top_excel → [idx_producto, ...]
    if cbm_mc_col:
        for rng in ws.merged_cells.ranges:
            if rng.min_col <= cbm_mc_col <= rng.max_col:
                top = rng.min_row
                for r in range(rng.min_row, rng.max_row + 1):
                    grupos_merged.setdefault(top, [])

        # Mapear fila_excel → índice de producto
        fila_a_prod = {}
        for i, p in enumerate(productos):
            fila = p.get("fila_excel_0idx")
            if fila is not None:
                fila_a_prod[fila + 1] = i   # convertir a 1-indexed

        for rng in ws.merged_cells.ranges:
            if cbm_mc_col and rng.min_col <= cbm_mc_col <= rng.max_col:
                top = rng.min_row
                for r in range(rng.min_row, rng.max_row + 1):
                    if r in fila_a_prod:
                        grupos_merged.setdefault(top, []).append(fila_a_prod[r])

    # ── Método 2: mismo id_guia no vacío ─────────────────────────────────────
    grupos_id: dict[str, list[int]] = {}
    for i, p in enumerate(productos):
        gid = str(p.get("id_guia") or "").strip()
        if gid:
            grupos_id.setdefault(gid, []).append(i)

    # Combinar grupos de ambos métodos
    grupos_finales: list[list[int]] = []
    usados: set[int] = set()

    for indices in list(grupos_merged.values()) + list(grupos_id.values()):
        indices = [i for i in indices if i not in usados]
        if len(indices) > 1:
            grupos_finales.append(indices)
            usados.update(indices)

    # ── Aplicar cálculo de inner carton a cada grupo ──────────────────────────
    for grupo in grupos_finales:
        n = len(grupo)
        # Tomar CBM master del primer producto del grupo que lo tenga
        cbm_master = next(
            (float(productos[i].get("cbm_master_carton") or
                   productos[i].get("cbm_total_sku") or
                   productos[i].get("cbm_por_pieza") or 0)
             for i in grupo), 0
        )
        if cbm_master <= 0:
            continue

        cbm_inner = round(cbm_master / n, 6)
        nombres = []
        for i in grupo:
            p = productos[i]
            pzs_caja = float(p.get("piezas_x_caja") or 0)
            cbm_pz   = round(cbm_inner / pzs_caja, 6) if pzs_caja > 0 else None
            p["cbm_master_carton"] = cbm_master
            p["cbm_inner_carton"]  = cbm_inner
            p["cbm_por_pieza"]     = cbm_pz
            nombres.append(p.get("nombre", "?"))

        advertencias.append(
            f"Master carton compartido ({n} productos): {', '.join(nombres)}. "
            f"CBM master={cbm_master} → CBM inner={cbm_inner} por producto"
        )

    return productos


def _tiene_chino(texto: str) -> bool:
    """Devuelve True si el texto contiene caracteres chinos."""
    return bool(re.search(r'[\u4e00-\u9fff\u3400-\u4dbf]', str(texto)))


def normalizar_nombres_productos(productos: list[dict]) -> list[dict]:
    """
    Para cada producto, elige el nombre más descriptivo entre nombre y nombre_alt
    y lo traduce al español si está en chino. Actualiza prod["nombre"] en-lugar.
    Hace una sola llamada batch a Claude para todos los nombres que lo necesiten.
    """
    # ── Guardar nombre chino original ANTES de traducir ──────────────────────
    # Se usa después en detectar_productos_duplicados como señal primaria de
    # "mismo producto". El nombre chino es más específico y descriptivo que el inglés.
    for prod in productos:
        if "nombre_chino_orig" not in prod:
            _n_raw   = str(prod.get("nombre")     or "").strip()
            _n_alt   = str(prod.get("nombre_alt") or "").strip()
            # Guardar el nombre en chino si existe; si no, guardar el principal
            if _tiene_chino(_n_raw):
                prod["nombre_chino_orig"] = _n_raw
            elif _tiene_chino(_n_alt):
                prod["nombre_chino_orig"] = _n_alt
            else:
                prod["nombre_chino_orig"] = _n_raw  # inglés/español — servirá igual

    # Detectar qué productos necesitan resolución
    # Usamos un índice LOCAL (0, 1, 2…) en el JSON para Claude para evitar
    # que Claude confunda los índices reales de la lista con sus propios números.
    # El mapeo local_idx → índice_real se guarda en local_a_real.
    local_a_real: dict[int, int] = {}
    items_para_claude: list[dict] = []

    for real_i, prod in enumerate(productos):
        nombre     = str(prod.get("nombre") or "").strip()
        nombre_alt = str(prod.get("nombre_alt") or "").strip()

        if _tiene_chino(nombre) or _tiene_chino(nombre_alt):
            local_idx = len(items_para_claude)
            local_a_real[local_idx] = real_i
            items_para_claude.append({
                "idx":        local_idx,          # índice local, siempre 0,1,2…
                "nombre":     nombre,
                "nombre_alt": nombre_alt or None,
            })

    if not items_para_claude:
        return productos  # todos en inglés/español, no hay nada que hacer

    # Llamada batch a Claude — índice local garantiza que no haya confusión de filas
    lista_json = json.dumps(items_para_claude, ensure_ascii=False, indent=2)
    prompt = f"""Tienes una lista de productos importados de China. Cada uno puede tener un nombre principal y un nombre alternativo (en inglés, chino o ambos).

Tu tarea para cada producto:
1. Elige el nombre MÁS ESPECÍFICO y DESCRIPTIVO entre "nombre" y "nombre_alt" (si existe).
   Si el nombre en chino describe mejor el producto (incluye color, modelo, batería, etc.) elige el chino.
2. Si el nombre elegido está en chino → tradúcelo al español de forma comercial y clara (máx 80 caracteres).
3. Si el nombre elegido está en inglés → tradúcelo también al español claro.
4. Devuelve EXACTAMENTE los mismos valores de "idx" que recibiste — no los cambies ni renumeres.

Devuelve ÚNICAMENTE un array JSON (sin texto adicional, sin markdown):
[
  {{"idx": <número igual al recibido>, "nombre_final": "<nombre en español>"}},
  ...
]
Debe haber exactamente {len(items_para_claude)} elementos, uno por cada producto recibido.

Productos:
{lista_json}"""

    try:
        # max_tokens suficiente para listas grandes (80 chars × n productos + overhead)
        _max_tok = min(200 * len(items_para_claude) + 500, 8000)
        llm   = ChatAnthropic(model="claude-haiku-4-5-20251001", max_tokens=_max_tok)
        resp  = llm.invoke([HumanMessage(content=prompt)])
        texto = resp.content.strip()
        if texto.startswith("```"):
            partes = texto.split("```")
            texto  = partes[1].lstrip("json").strip() if len(partes) > 1 else texto
        resultados = json.loads(texto)

        for r in resultados:
            local_idx    = r.get("idx")
            nombre_final = str(r.get("nombre_final") or "").strip()
            # Validar que local_idx esté en nuestro mapa — nunca confiar ciegamente
            if local_idx is not None and nombre_final and local_idx in local_a_real:
                productos[local_a_real[local_idx]]["nombre"] = nombre_final
    except Exception:
        # Si falla, dejamos los nombres originales — no bloqueamos el flujo
        pass

    return productos


# Campos determinantes: precio, CBM y piezas/caja (incluye cbm_total_sku para archivos sin CBM por pieza)
_CAMPOS_DETERMINANTES = [
    "precio_usd", "cbm_por_pieza", "cbm_master_carton", "piezas_x_caja", "cbm_total_sku",
]
# Campos informativos para mostrar diferencias
_CAMPOS_DIFF_DISPLAY  = [
    "precio_usd", "cbm_por_pieza", "cbm_master_carton",
    "cbm_total_sku", "piezas_x_caja", "piezas_total", "cajas_master", "material", "uso",
]
_LABELS_DIFF = {
    "precio_usd": "Precio USD", "cbm_por_pieza": "CBM/pieza",
    "cbm_master_carton": "CBM master", "cbm_total_sku": "CBM total",
    "piezas_x_caja": "Pzas/caja", "piezas_total": "Pzas total",
    "cajas_master": "Cajas", "material": "Material", "uso": "Uso",
}


def _det_iguales(a: dict, b: dict) -> bool:
    # Si no hay ningún campo con valor real en ambos productos, no podemos afirmar que son iguales
    tiene_datos = any(
        a.get(c) not in (None, "", 0) and b.get(c) not in (None, "", 0)
        for c in _CAMPOS_DETERMINANTES
    )
    if not tiene_datos:
        return False
    return all(str(a.get(c, "")).strip() == str(b.get(c, "")).strip() for c in _CAMPOS_DETERMINANTES)


def _nombre_norm(p: dict) -> str:
    return str(p.get("nombre","") or "").strip().lower()


def _sim_chino(a: dict, b: dict) -> float:
    """
    Similitud [0–1] entre los nombres chinos originales de dos productos.
    Usa SequenceMatcher sobre los caracteres completos.
    Devuelve 0.0 si alguno no tiene nombre chino.
    """
    ca = str(a.get("nombre_chino_orig") or a.get("nombre", "")).strip()
    cb = str(b.get("nombre_chino_orig") or b.get("nombre", "")).strip()
    if not ca or not cb:
        return 0.0
    return difflib.SequenceMatcher(None, ca, cb).ratio()


# Palabras que indican variante — al quitarlas queda el "nombre base" del producto
# ══════════════════════════════════════════════════════════════════════════════
# Clasificación de palabras de variante en 3 niveles:
#
#  🟢 VERDE  — atributos cosméticos: siempre indican variante, seguros de quitar
#  🟡 AMARILLO — material / presentación: posible variante, requiere confirmación
#  🔴 ROJO   — tipo / especificación / modelo: definen el producto, NO quitar
#
# Regla: si los nombres de dos productos difieren SOLO en palabras VERDES → variante.
#        Si difieren en VERDE+AMARILLO (pero base verde coincide) → posible variante.
#        Si difieren en palabras ROJAS → productos DIFERENTES (no variante).
# ══════════════════════════════════════════════════════════════════════════════

_PALABRAS_VERDE: set[str] = {
    # Colores español
    "rojo","roja","rojos","rojas",
    "azul","azules",
    "verde","verdes",
    "amarillo","amarilla","amarillos","amarillas",
    "negro","negra","negros","negras",
    "blanco","blanca","blancos","blancas",
    "gris","grises",
    "naranja","naranjas",
    "morado","morada","morados","moradas",
    "rosa","rosas",
    "cafe","café",
    "marron","marrón",
    "plateado","plateada","plateados","plateadas",
    "dorado","dorada","dorados","doradas",
    "beige","turquesa","lila","violeta","multicolor",
    "transparente","cromado","cromada",
    # Colores inglés
    "red","blue","yellow","black","white","gray","grey",
    "orange","purple","pink","brown","gold","silver",
    # Tallas cosméticas (sin impacto funcional)
    "xs","sm","md","lg","xl","xxl","xxxl",
}

_PALABRAS_AMARILLO: set[str] = {
    # Tamaños relativos (pueden o no afectar funcionalidad)
    "mini","maxi","micro",
    "pequeño","pequeña","chico","chica",
    "mediano","mediana","grande","grandes",
    "largo","larga","corto","corta",
    "angosto","angosta","compacto","compacta",
    # Materiales (pueden definir funcionalidad, pero a veces son solo variantes)
    "madera","mdf","bambu","bambú","roble","pino","nogal","haya",
    "metal","metalico","metálico","metalica","metálica",
    "acero","inoxidable","aluminio","hierro","fierro",
    "plastico","plástico","plastica","plástica","pvc",
    "tela","lona","cuero","piel","nylon","poliester","poliéster","oxford",
    "vidrio","cristal","ceramica","cerámica","porcelana",
    "silicona","goma","hule","latex","látex",
    # Presentación / pack
    "pack","set","kit","pieza","piezas","unidad","unidades",
    "simple","doble","triple",
}

# 🔴 ROJO — palabras que DEFINEN el producto; si difieren → NO son variantes
_PALABRAS_ROJO: set[str] = {
    # Tipo de operación (cambia totalmente el producto)
    "electrico","eléctrico","electrica","eléctrica",
    "manual","automatico","automático","automatica","automática",
    "inalambrico","inalámbrico","inalambrica","inalámbrica",
    "recargable","a","pilas",
    "neumatico","neumático","hidraulico","hidráulico",
    # Versión / tier (puede ser upgrade, no solo cosmético)
    "pro","lite","plus","max","ultra","prime","elite","sport","deluxe",
    # Plegabilidad / configuración estructural
    "plegable","portatil","portátil","fijo","fija","desmontable",
}

# Unión completa para casos donde solo necesitamos saber si es "de variante"
_PALABRAS_VARIANTE: set[str] = _PALABRAS_VERDE | _PALABRAS_AMARILLO


def _es_tecnico(token: str) -> bool:
    """True si el token es una especificación técnica: 21v, 500w, 2l, 12kg, etc."""
    return bool(re.match(r'^\d+(\.\d+)?[a-z]+$', token))


def _nombre_base_verde(nom: str) -> str:
    """
    Quita SOLO atributos cosméticos (colores, tallas simples).
    Si dos productos tienen la misma base verde → variante SEGURA.
    Conserva especificaciones técnicas (21v, 500w…) y palabras ROJAS (definen producto).
    """
    tokens = []
    for t in nom.split():
        if _es_tecnico(t):
            tokens.append(t)          # siempre conservar specs técnicas
        elif t in _PALABRAS_VERDE:
            pass                       # quitar: cosmético
        elif len(t) > 2:
            tokens.append(t)
    return " ".join(tokens)


def _nombre_base_amarillo(nom: str) -> str:
    """
    Quita atributos cosméticos + materiales / presentación.
    Si dos productos tienen la misma base amarilla (pero bases verdes distintas)
    → posible variante, requiere confirmación por imagen o precio.
    """
    tokens = []
    for t in nom.split():
        if _es_tecnico(t):
            tokens.append(t)
        elif t in _PALABRAS_VERDE or t in _PALABRAS_AMARILLO:
            pass                       # quitar
        elif len(t) > 2:
            tokens.append(t)
    return " ".join(tokens)


def _tiene_diferencia_roja(nom_a: str, nom_b: str) -> bool:
    """
    True si los nombres difieren en al menos una palabra ROJA (tipo, specs técnicas).
    Indica que son productos distintos, no variantes.
    Ejemplo: 'pistola eléctrica' vs 'pistola manual' → True → NO variante.
    """
    tok_a = set(nom_a.split())
    tok_b = set(nom_b.split())
    rojas_a = tok_a & _PALABRAS_ROJO
    rojas_b = tok_b & _PALABRAS_ROJO
    # Palabras rojas en uno pero no en el otro → diferencia definitoria
    if rojas_a.symmetric_difference(rojas_b):
        return True
    # Especificaciones técnicas distintas (21v vs 18v, 500w vs 800w)
    specs_a = {t for t in tok_a if _es_tecnico(t)}
    specs_b = {t for t in tok_b if _es_tecnico(t)}
    if specs_a and specs_b and specs_a != specs_b:
        return True
    return False




def _nombres_son_similares(nom_a: str, nom_b: str, ratio_min: float = 0.78) -> bool:
    """
    True si los dos nombres normalizados son suficientemente similares.
    Usa dos métricas:
    1. SequenceMatcher ratio ≥ ratio_min  (bueno cuando los nombres tienen longitud similar)
    2. Solapamiento de tokens: ≥75% de las palabras del nombre más corto
       aparecen en el más largo  (bueno cuando uno tiene palabras extra como
       "2 Baterías", "Mod A", "Set", etc.)
    """
    if not nom_a or not nom_b or nom_a == nom_b:
        return False
    if difflib.SequenceMatcher(None, nom_a, nom_b).ratio() >= ratio_min:
        return True
    # Fallback: solapamiento de tokens
    tok_a = set(nom_a.split())
    tok_b = set(nom_b.split())
    if not tok_a or not tok_b:
        return False
    shorter = tok_a if len(tok_a) <= len(tok_b) else tok_b
    longer  = tok_b if len(tok_a) <= len(tok_b) else tok_a
    # Ignorar tokens muy cortos (números sueltos, artículos)
    shorter_sig = {t for t in shorter if len(t) > 2}
    if not shorter_sig:
        return False
    overlap = len(shorter_sig & longer) / len(shorter_sig)
    return overlap >= 0.75


def _diffs_display(prods: list[dict]) -> dict[str, list[str]]:
    """Devuelve {campo: [val_prod0, val_prod1, ...]} solo para campos que varían."""
    result = {}
    for campo in _CAMPOS_DIFF_DISPLAY:
        vals = [str(p.get(campo,"") or "—").strip() for p in prods]
        if len(set(vals)) > 1:
            result[campo] = vals
    return result


def detectar_productos_duplicados(
    productos: list[dict],
    imagenes: dict | None = None,
) -> list[dict]:
    """
    FASE 0 — Nombre chino original (señal más fuerte, se corre primero):
      A0. Nombre chino idéntico (≥98%) → "exacto" (auto-fusiona)
      A1. Nombre chino ≥80% + phash imagen ≤ 6 → "exacto" (auto-fusiona)
      A2. Nombre chino ≥82% (sin imagen) → "probable" (muestra al usuario)
      A3. Nombre chino ≥62% → "similar" (posible variante por color/talla)

    FASE 1 — Mismo producto por nombre español/determinantes (auto-fusiona):
      B. Imagen idéntica (phash ≤ 6)  Y  mismo nombre normalizado
      C. Mismo nombre normalizado  (sin imagen)

    FASE 2 — Posibles variantes (se le pregunta al usuario):
      D. Determinantes iguales (precio/CBM/piezas) + nombres distintos → "similar"
      E. Mismo nombre base (sin colores/tallas/modelo) → "nombre_similar"
      F. Imagen similar (phash ≤ 10) + nombre parcialmente similar (≥50% tokens)
      G. Nombre fuzzy similar (≥78%) o solapamiento de tokens ≥75%

    Post-procesamiento: fusionar grupos de variantes que comparten nombre base o fuzzy.
    """
    grupos: list[dict] = []
    usados: set[int]   = set()
    _FUZZY_RATIO = 0.78

    def _agregar(indices: list[int], tipo: str) -> None:
        prods = [productos[k] for k in indices]
        grupos.append({
            "id":        len(grupos),
            "indices":   indices,
            "productos": prods,
            "tipo":      tipo,
            "diffs":     _diffs_display(prods),
        })

    # Precalcular phashes si hay imágenes
    phashes: dict[int, object] = {}
    if imagenes:
        for idx, prod in enumerate(productos):
            fila = prod.get("fila_excel_0idx")
            img  = imagenes.get(fila + 1) if fila is not None else None
            if img:
                ph = _phash_imagen(img["data"])
                if ph is not None:
                    phashes[idx] = ph

    # ══════════════════════════════════════════════════════════════════════════
    # FASE 0 — NOMBRE CHINO ORIGINAL (señal más fuerte, prioritaria)
    # El nombre chino es más específico que el inglés: incluye color, modelo,
    # batería, etc. Si coincide a nivel alto, probablemente es el mismo producto.
    # ══════════════════════════════════════════════════════════════════════════

    # A0: Nombre chino IDÉNTICO o casi (≥98%) → mismo producto, auto-fusiona
    _CHINO_EXACTO   = 0.98
    _CHINO_PROBABLE = 0.82  # mismo producto con ligera variación ortográfica
    _CHINO_VARIANTE = 0.62  # misma línea de producto, atributo distinto (color/talla)

    _hay_chino = any(_tiene_chino(str(p.get("nombre_chino_orig", ""))) for p in productos)

    if _hay_chino:
        # A0 — Idéntico (auto-fusiona)
        for i in range(len(productos)):
            if i in usados:
                continue
            if not _tiene_chino(str(productos[i].get("nombre_chino_orig", ""))):
                continue
            grp = [i]
            for j in range(i + 1, len(productos)):
                if j in usados:
                    continue
                if _sim_chino(productos[i], productos[j]) >= _CHINO_EXACTO:
                    grp.append(j)
            if len(grp) > 1:
                for k in grp:
                    usados.add(k)
                _agregar(grp, "exacto")

        # A1 — Nombre chino muy similar + imagen también similar → auto-fusiona
        for i in range(len(productos)):
            if i in usados or i not in phashes:
                continue
            if not _tiene_chino(str(productos[i].get("nombre_chino_orig", ""))):
                continue
            grp = [i]
            for j in range(i + 1, len(productos)):
                if j in usados or j not in phashes:
                    continue
                sim_n = _sim_chino(productos[i], productos[j])
                sim_i = phashes[i] - phashes[j]
                if sim_n >= 0.80 and sim_i <= 6:
                    grp.append(j)
            if len(grp) > 1:
                for k in grp:
                    usados.add(k)
                _agregar(grp, "exacto")

        # A2 — Nombre chino muy similar (sin imagen) → probable mismo producto
        for i in range(len(productos)):
            if i in usados:
                continue
            if not _tiene_chino(str(productos[i].get("nombre_chino_orig", ""))):
                continue
            grp = [i]
            for j in range(i + 1, len(productos)):
                if j in usados:
                    continue
                if _CHINO_PROBABLE <= _sim_chino(productos[i], productos[j]) < _CHINO_EXACTO:
                    grp.append(j)
            if len(grp) > 1:
                for k in grp:
                    usados.add(k)
                _agregar(grp, "probable")

        # A3 — Nombre chino moderadamente similar → posible variante (color/talla)
        for i in range(len(productos)):
            if i in usados:
                continue
            if not _tiene_chino(str(productos[i].get("nombre_chino_orig", ""))):
                continue
            grp = [i]
            for j in range(i + 1, len(productos)):
                if j in usados:
                    continue
                sim = _sim_chino(productos[i], productos[j])
                if _CHINO_VARIANTE <= sim < _CHINO_PROBABLE:
                    grp.append(j)
            if len(grp) > 1:
                for k in grp:
                    usados.add(k)
                _agregar(grp, "similar")

    # ══════════════════════════════════════════════════════════════════════════
    # FASE 1 — MISMO PRODUCTO (exacto) → auto-fusiona
    # ══════════════════════════════════════════════════════════════════════════

    # B: misma foto (phash ≤ 2) + mismo nombre exacto → auto-fusiona sin preguntar
    # phash ≤ 2 = literalmente la misma imagen (mismo pixel hash).
    # Necesitamos AMBAS señales para evitar falsos positivos.
    for i in range(len(productos)):
        if i in usados or i not in phashes:
            continue
        nom_i = _nombre_norm(productos[i])
        grp = [i]
        for j in range(i + 1, len(productos)):
            if j in usados or j not in phashes:
                continue
            if phashes[i] - phashes[j] <= 2 and _nombre_norm(productos[j]) == nom_i:
                grp.append(j)
        if len(grp) > 1:
            for k in grp: usados.add(k)
            _agregar(grp, "exacto")

    # C: misma foto (phash ≤ 2) + nombre distinto → pregunta al usuario
    # Ejemplo: mismo zapato, distinta talla en el nombre → probable variante
    for i in range(len(productos)):
        if i in usados or i not in phashes:
            continue
        grp      = [i]
        noms_grp = [_nombre_norm(productos[i])]
        for j in range(i + 1, len(productos)):
            if j in usados or j not in phashes:
                continue
            if any(phashes[j] - phashes[k] <= 2 for k in grp if k in phashes):
                grp.append(j)
                noms_grp.append(_nombre_norm(productos[j]))
        if len(grp) > 1:
            for k in grp: usados.add(k)
            _agregar(grp, "variante_imagen")

    # D-esp: mismo nombre español exacto (sin imagen idéntica) → pregunta al usuario
    for i in range(len(productos)):
        if i in usados:
            continue
        nom_i = _nombre_norm(productos[i])
        grp = [i]
        for j in range(i + 1, len(productos)):
            if j in usados:
                continue
            if _nombre_norm(productos[j]) == nom_i:
                grp.append(j)
        if len(grp) > 1:
            for k in grp: usados.add(k)
            _agregar(grp, "probable")

    # ══════════════════════════════════════════════════════════════════════════
    # FASE 2 — POSIBLES VARIANTES → pregunta al usuario
    # Regla de 3 niveles (Verde / Amarillo / Rojo):
    #   Si nombres difieren en palabras ROJAS (tipo, specs) → NO son variantes → skip
    #   Si base VERDE coincide → variante segura (solo difieren en color/talla)
    #   Si base AMARILLA coincide (pero verde no) → posible variante (requiere imagen o precio)
    # ══════════════════════════════════════════════════════════════════════════

    # D: base VERDE igual + sin diferencia ROJA → variante segura por nombre
    # Ej: "Pistola agua Azul" vs "Pistola agua Rosa" → base verde = "Pistola agua"
    for i in range(len(productos)):
        if i in usados:
            continue
        nom_i   = _nombre_norm(productos[i])
        base_vi = _nombre_base_verde(nom_i)
        if len(base_vi.split()) < 2:
            continue
        grp = [i]
        for j in range(i + 1, len(productos)):
            if j in usados:
                continue
            nom_j   = _nombre_norm(productos[j])
            base_vj = _nombre_base_verde(nom_j)
            if (base_vj == base_vi
                    and nom_i != nom_j
                    and not _tiene_diferencia_roja(nom_i, nom_j)):
                grp.append(j)
        if len(grp) > 1:
            for k in grp:
                usados.add(k)
            _agregar(grp, "nombre_similar")

    # E: base AMARILLA igual + imagen similar → posible variante por material/tamaño
    # Ej: "Mesa Madera Grande" vs "Mesa Metal Grande" — mismo producto, distinto material
    for i in range(len(productos)):
        if i in usados or i not in phashes:
            continue
        nom_i   = _nombre_norm(productos[i])
        base_ai = _nombre_base_amarillo(nom_i)
        if len(base_ai.split()) < 2:
            continue
        grp = [i]
        for j in range(i + 1, len(productos)):
            if j in usados or j not in phashes:
                continue
            nom_j   = _nombre_norm(productos[j])
            base_aj = _nombre_base_amarillo(nom_j)
            img_ok  = any(phashes[j] - phashes[k] <= 10 for k in grp if k in phashes)
            if (base_aj == base_ai
                    and nom_i != nom_j
                    and not _tiene_diferencia_roja(nom_i, nom_j)
                    and img_ok):
                grp.append(j)
        if len(grp) > 1:
            for k in grp:
                usados.add(k)
            _agregar(grp, "similar")

    # F: determinantes iguales (precio, CBM, piezas) + sin diferencia ROJA → variante
    # Fallback cuando no hay nombre descriptivo pero los datos físicos/precio coinciden
    for i in range(len(productos)):
        if i in usados:
            continue
        nom_i = _nombre_norm(productos[i])
        grp = [i]
        for j in range(i + 1, len(productos)):
            if j in usados:
                continue
            nom_j = _nombre_norm(productos[j])
            if (_det_iguales(productos[i], productos[j])
                    and nom_i != nom_j
                    and not _tiene_diferencia_roja(nom_i, nom_j)):
                grp.append(j)
        if len(grp) > 1:
            for k in grp:
                usados.add(k)
            _agregar(grp, "similar")

    # G: imagen similar (phash ≤ 10) + nombre parcialmente similar (≥50% tokens)
    # + sin diferencia ROJA → variante por imagen (cuando el nombre no es concluyente)
    for i in range(len(productos)):
        if i in usados or i not in phashes:
            continue
        nom_i    = _nombre_norm(productos[i])
        grp      = [i]
        noms_grp = [nom_i]
        for j in range(i + 1, len(productos)):
            if j in usados or j not in phashes:
                continue
            nom_j  = _nombre_norm(productos[j])
            img_ok = any(phashes[j] - phashes[k] <= 10 for k in grp if k in phashes)
            nom_ok = any(_nombres_son_similares(nom_j, nom_k, 0.50) for nom_k in noms_grp)
            if img_ok and nom_ok and not _tiene_diferencia_roja(nom_i, nom_j):
                grp.append(j)
                noms_grp.append(nom_j)
        if len(grp) > 1:
            for k in grp: usados.add(k)
            _agregar(grp, "variante_imagen")

    # H: nombre fuzzy similar (≥78%) + sin diferencia ROJA
    for i in range(len(productos)):
        if i in usados:
            continue
        nom_i = _nombre_norm(productos[i])
        if not nom_i:
            continue
        grp      = [i]
        noms_grp = [nom_i]
        for j in range(i + 1, len(productos)):
            if j in usados:
                continue
            nom_j = _nombre_norm(productos[j])
            if not nom_j:
                continue
            if (any(_nombres_son_similares(nom_j, nom_k, _FUZZY_RATIO) for nom_k in noms_grp)
                    and not _tiene_diferencia_roja(nom_i, nom_j)):
                grp.append(j)
                noms_grp.append(nom_j)
        if len(grp) > 1:
            for k in grp: usados.add(k)
            _agregar(grp, "nombre_similar")

    # ── Post-procesamiento: fusionar grupos de variantes relacionados ──────────
    # Comprueba nombre fuzzy, base, phash Y nombre chino entre productos de distintos grupos.
    # Esto garantiza que si {A,B} = mismo producto y {C,D} = variantes, pero C también
    # es variante de A, los cuatro quedan en un solo grupo.
    _PRIO = {"exacto": 4, "probable": 3, "similar": 3, "variante_imagen": 2, "nombre_similar": 1}

    def _grupos_relacionados(ga: dict, gb: dict) -> bool:
        _noms_a  = [_nombre_norm(p) for p in ga["productos"]]
        _noms_b  = [_nombre_norm(p) for p in gb["productos"]]
        _bases_va = [_nombre_base_verde(n) for n in _noms_a if len(_nombre_base_verde(n).split()) >= 2]
        _bases_vb = [_nombre_base_verde(n) for n in _noms_b if len(_nombre_base_verde(n).split()) >= 2]
        _bases_aa = [_nombre_base_amarillo(n) for n in _noms_a if len(_nombre_base_amarillo(n).split()) >= 2]
        _bases_ab = [_nombre_base_amarillo(n) for n in _noms_b if len(_nombre_base_amarillo(n).split()) >= 2]
        # 1. Nombre fuzzy similar
        if any(_nombres_son_similares(na, nb, _FUZZY_RATIO) for na in _noms_a for nb in _noms_b):
            return True
        # 2. Mismo nombre base (verde o amarillo)
        if any(bva == bvb for bva in _bases_va for bvb in _bases_vb):
            return True
        if any(baa == bab for baa in _bases_aa for bab in _bases_ab):
            return True
        # 3. Nombre chino similar (≥ umbral variante)
        _chin_a = [p for p in ga["productos"] if _tiene_chino(str(p.get("nombre_chino_orig", "")))]
        _chin_b = [p for p in gb["productos"] if _tiene_chino(str(p.get("nombre_chino_orig", "")))]
        if _chin_a and _chin_b:
            if any(_sim_chino(pa, pb) >= _CHINO_VARIANTE for pa in _chin_a for pb in _chin_b):
                return True
        # 4. Phash similar entre cualquier par de productos de ambos grupos
        _ph_a = {idx: phashes[idx] for idx in ga["indices"] if idx in phashes}
        _ph_b = {idx: phashes[idx] for idx in gb["indices"] if idx in phashes}
        if _ph_a and _ph_b:
            if any(pha - phb <= 10 for pha in _ph_a.values() for phb in _ph_b.values()):
                # Solo fusionar si además los nombres no tienen diferencia roja
                if not any(
                    _tiene_diferencia_roja(_nombre_norm(pa), _nombre_norm(pb))
                    for pa in ga["productos"] for pb in gb["productos"]
                ):
                    return True
        return False

    _fusionado = True
    while _fusionado:
        _fusionado = False
        for _a in range(len(grupos)):
            if grupos[_a]["tipo"] == "exacto":
                continue
            for _b in range(_a + 1, len(grupos)):
                if grupos[_b]["tipo"] == "exacto":
                    continue
                if _grupos_relacionados(grupos[_a], grupos[_b]):
                    _tipo_m  = min(grupos[_a]["tipo"], grupos[_b]["tipo"], key=lambda t: _PRIO.get(t, 0))
                    _all_idx = grupos[_a]["indices"] + grupos[_b]["indices"]
                    _all_p   = grupos[_a]["productos"] + grupos[_b]["productos"]
                    grupos[_a] = {
                        "id":        grupos[_a]["id"],
                        "indices":   _all_idx,
                        "productos": _all_p,
                        "tipo":      _tipo_m,
                        "diffs":     _diffs_display(_all_p),
                    }
                    grupos.pop(_b)
                    _fusionado = True
                    break
            if _fusionado:
                break

    return grupos


def aplicar_resolucion_duplicados(productos: list[dict], grupos: list[dict], respuestas: dict) -> list[dict]:
    """
    respuestas: {str(gid): "diferente" | dict}
    - "diferente"                             → todos independientes
    - {"tipo": "mismo",    "sel": [ci,...]}   → fusiona los seleccionados; resto independientes
    - {"tipo": "variantes","sel": [ci,...]}   → seleccionados comparten base de SKU; resto independientes
    """
    _CAMPOS_SUMA = ("piezas_total", "cajas_master", "cbm_total_sku")
    indices_a_eliminar: set[int] = set()

    def _fusionar(indices_sub: list[int]) -> None:
        if len(indices_sub) < 2:
            return
        base = productos[indices_sub[0]]
        for campo in _CAMPOS_SUMA:
            total = sum(float(productos[idx].get(campo) or 0) for idx in indices_sub)
            if total > 0:
                base[campo] = total
        for idx in indices_sub[1:]:
            stem = productos[idx].get("imagen_temp_stem")
            if stem and IMAGENES_TEMP_PATH.exists():
                for f in IMAGENES_TEMP_PATH.iterdir():
                    if f.stem == stem:
                        try:
                            f.unlink()
                        except Exception:
                            pass
                        break
            indices_a_eliminar.add(idx)

    for grupo in grupos:
        gid_str  = str(grupo["id"])
        decision = respuestas.get(gid_str, "diferente")

        if decision == "diferente":
            pass  # todos independientes

        elif isinstance(decision, dict):
            tipo = decision.get("tipo", "variantes")

            # Soporte para sub-grupos: {"tipo": ..., "subgrupos": [[ci,...], [ci,...]]}
            # o formato antiguo:       {"tipo": ..., "sel": [ci,...]}
            if "subgrupos" in decision:
                subgrupos_ci = [sg for sg in decision["subgrupos"] if sg]
            else:
                sel = decision.get("sel", list(range(len(grupo["indices"]))))
                subgrupos_ci = [sel]

            for sg_idx, sg_cis in enumerate(subgrupos_ci):
                indices_sel = [grupo["indices"][ci] for ci in sg_cis if ci < len(grupo["indices"])]
                if not indices_sel:
                    continue
                if tipo == "mismo":
                    _fusionar(indices_sel)
                elif tipo == "variantes" and len(indices_sel) > 1:
                    gid_v = f"var_{grupo['id']}_{sg_idx}"
                    for idx in indices_sel:
                        productos[idx]["_grupo_variante"] = gid_v
                # Si solo 1 producto en el subgrupo → independiente

    return [p for i, p in enumerate(productos) if i not in indices_a_eliminar]


# ══════════════════════════════════════════════════════════════════════════════
# ══════════════════════════════════════════════════════════════════════════════
# Cache local de datos ODOO
# ══════════════════════════════════════════════════════════════════════════════

CACHE_PATH = Path(__file__).parent / "odoo_cache.pkl"
CACHE_MAX_HORAS = 24  # horas antes de considerar el cache desactualizado
IMAGENES_TEMP_PATH = Path(__file__).parent / "imagenes_temp"


def guardar_cache_odoo(skus: list[str], productos: list[dict], phashes: dict) -> None:
    """Guarda SKUs, productos (sin imagen) y phashes en cache local."""
    # Guardar phashes como strings hex para serialización segura
    phashes_hex = {k: str(v) for k, v in phashes.items()}
    # Quitar image_128 de productos para mantener el cache liviano
    prods_sin_img = [
        {k: v for k, v in p.items() if k != "image_128"}
        for p in productos
    ]
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
    Devuelve dict con skus, productos, phashes (objetos imagehash) o None.
    """
    if not CACHE_PATH.exists():
        return None
    try:
        with open(CACHE_PATH, "rb") as f:
            cache = pickle.load(f)
        ts = datetime.fromisoformat(cache["timestamp"])
        if datetime.now() - ts > timedelta(hours=CACHE_MAX_HORAS):
            return None  # cache expirado
        # Convertir phashes de hex string a objetos imagehash
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


# ══════════════════════════════════════════════════════════════════════════════
# Supabase — cache persistente de productos ODOO
# ══════════════════════════════════════════════════════════════════════════════

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
    Devuelve dict con skus, productos, phashes, imagen_urls o None si no hay datos.
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
        skus        = [r["sku"] for r in rows]
        productos   = [{"default_code": r["sku"], "name": r.get("nombre", "")} for r in rows]
        phashes     = {}
        imagen_urls = {}
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
        return {
            "timestamp":   ts,
            "skus":        skus,
            "productos":   productos,
            "phashes":     phashes,
            "imagen_urls": imagen_urls,
        }
    except Exception:
        return None


def _sincronizar_a_supabase(prods_odoo: list[dict]) -> dict:
    """
    Sube productos de ODOO a Supabase:
    - Calcula phash para cada imagen
    - Sube la imagen al bucket 'odoo-imagenes'
    - Hace upsert en tabla 'odoo_productos'
    Devuelve dict con phashes e imagen_urls calculados.
    """
    sb          = _supabase_client()
    phashes     = {}
    imagen_urls = {}
    if not sb:
        return {"phashes": phashes, "imagen_urls": imagen_urls}

    bucket = sb.storage.from_("odoo-imagenes")

    # Calcular phashes primero (local, sin red)
    rows_to_upsert = []
    prods_con_img  = []
    for p in prods_odoo:
        sku    = p.get("default_code", "")
        nombre = p.get("name", "")
        if not sku:
            continue
        phash_hex  = None
        img_b64    = p.get("image_128")
        img_bytes  = None
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

    # Subir imágenes en paralelo (20 workers)
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

    # Actualizar imagen_url en rows antes del upsert
    for row in rows_to_upsert:
        row["imagen_url"] = imagen_urls.get(row["sku"])

    # Upsert en batches de 100
    try:
        for i in range(0, len(rows_to_upsert), 100):
            sb.table("odoo_productos").upsert(rows_to_upsert[i:i+100]).execute()
    except Exception:
        pass

    return {"phashes": phashes, "imagen_urls": imagen_urls}


# ══════════════════════════════════════════════════════════════════════════════
# Similitud de imágenes y nombres
# ══════════════════════════════════════════════════════════════════════════════

def _phash_imagen(image_data: bytes):
    """Genera perceptual hash de una imagen. Devuelve objeto imagehash o None."""
    try:
        import imagehash
        from PIL import Image
        img = Image.open(io.BytesIO(image_data)).convert("RGB")
        return imagehash.phash(img)
    except Exception:
        return None


def _similitud_nombres(a: str, b: str) -> float:
    from difflib import SequenceMatcher
    return SequenceMatcher(None, a.lower().strip(), b.lower().strip()).ratio()


def buscar_similares_odoo(
    imagen_data: bytes | None,
    nombre_producto: str,
    productos_odoo: list[dict],
    phashes_odoo: dict,
    umbral_phash: int = 12,
    umbral_nombre: float = 0.70,
) -> list[dict]:
    """
    Busca productos similares en ODOO por imagen (perceptual hash) y por nombre (fuzzy).
    Devuelve lista ordenada por relevancia con los campos:
      sku, nombre, similitud_imagen, similitud_nombre, por_imagen, por_nombre, producto_odoo
    """
    phash_nuevo = _phash_imagen(imagen_data) if imagen_data else None
    similares   = []

    for prod in productos_odoo:
        sku        = prod.get("default_code", "")
        nombre_odo = prod.get("name", "")
        if not sku:
            continue

        sim_nombre = _similitud_nombres(nombre_producto, nombre_odo)
        sim_img    = None
        por_imagen = False

        if phash_nuevo is not None and sku in phashes_odoo:
            dist      = phash_nuevo - phashes_odoo[sku]
            sim_img   = round(1.0 - dist / 64.0, 3)
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


# ══════════════════════════════════════════════════════════════════════════════
# ODOO — Carga y validación de SKUs existentes
# ══════════════════════════════════════════════════════════════════════════════

def cargar_skus_odoo(url: str, db: str, username: str, password: str) -> tuple[list[str], str | None]:
    """
    Obtiene todos los SKUs (default_code) de productos en ODOO via XML-RPC.
    Devuelve (lista_de_skus, mensaje_error_o_None).
    """
    try:
        common = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/common", allow_none=True)
        uid = common.authenticate(db, username, password, {})
        if not uid:
            return [], "Credenciales inválidas o usuario sin acceso."
        models = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/object", allow_none=True)
        productos = models.execute_kw(
            db, uid, password,
            "product.template", "search_read",
            [[["default_code", "!=", False]]],
            {"fields": ["default_code"], "limit": 0},
        )
        skus = [p["default_code"] for p in productos if p.get("default_code")]
        return skus, None
    except Exception as e:
        return [], str(e)


def cargar_todos_productos_odoo(url: str, db: str, username: str, password: str) -> list[dict]:
    """
    Carga nombre, SKU e imagen de TODOS los productos con SKU en ODOO.
    Se usa para pre-calcular phashes y hacer búsqueda de similitud.
    """
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
    """
    Trae nombre, descripción, categoría e imagen de productos ODOO por lista de SKUs.
    Usa las credenciales del .env directamente.
    """
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


def aplicar_resoluciones_conflictos(productos: list[dict], conflictos: list[dict], resoluciones: dict) -> list[dict]:
    """
    Aplica las decisiones del usuario sobre los conflictos SKU a la lista de productos.
    resoluciones: {str(idx): {tipo, sku_final, nombre, descripcion, categoria, atributo, precio_usd}}
    """
    for conflicto in conflictos:
        idx = conflicto["idx"]
        res = resoluciones.get(str(idx), {})
        prod = productos[idx]
        if res.get("tipo") == "mismo":
            prod["sku"]         = res.get("sku_final", prod.get("sku", ""))
            prod["nombre"]      = res.get("nombre",      prod.get("nombre", ""))
            prod["descripcion"] = res.get("descripcion", prod.get("descripcion", ""))
            prod["categoria"]   = res.get("categoria",   prod.get("categoria", ""))
            prod["atributo"]    = res.get("atributo",    prod.get("atributo", ""))
            if res.get("precio_usd"):
                prod["precio_usd"] = res["precio_usd"]
        else:
            # Producto diferente → SKU ajustado (max+1) ya calculado
            prod["sku"] = conflicto["sku_ajustado"]
    return productos


def sincronizar_contadores_con_odoo(skus_odoo: list[str]) -> None:
    """
    Actualiza sku_contadores en session_state para que los nuevos SKUs
    partan del máximo que ya existe en ODOO por subcategoría.
    """
    contadores = st.session_state.setdefault("sku_contadores", {})
    patron = re.compile(r"^([A-Z]{2,4})-(\d{4})-")
    for sku in skus_odoo:
        m = patron.match(sku)
        if m:
            sub, num = m.group(1), int(m.group(2))
            if contadores.get(sub, 0) < num:
                contadores[sub] = num


def validar_sku_vs_odoo(sku_propuesto: str, skus_odoo: list[str]) -> dict:
    """
    Verifica que el número secuencial del SKU propuesto sea mayor al máximo
    que existe en ODOO para ese SUBCAT, independientemente del atributo.

    Regla: OFI-0001-NEG y OFI-0001-BEI comparten el mismo número de producto.
    Si ODOO ya tiene OFI-####-XXX con #### >= num_propuesto, ajustar a max+1.

    Devuelve:
      conflicto      : bool  — True si el número propuesto ya está ocupado en ODOO
      skus_odoo_match: list  — SKUs de ODOO con el mismo SUBCAT-####
      sku_ajustado   : str   — SKU final (max+1 si hubo conflicto)
    """
    m = re.match(r"^([A-Z]{2,4})-(\d{4})-([A-Z]{2,4})$", sku_propuesto)
    if not m:
        return {"conflicto": False, "skus_odoo_match": [], "sku_ajustado": sku_propuesto}

    subcat, num_str, atrib = m.group(1), m.group(2), m.group(3)
    num_propuesto = int(num_str)

    # Buscar el máximo número existente en ODOO para este SUBCAT (cualquier atributo)
    patron_sub = re.compile(rf"^{re.escape(subcat)}-(\d{{4}})-", re.IGNORECASE)
    numeros_odoo = [
        int(p.group(1)) for s in skus_odoo if (p := patron_sub.match(s))
    ]

    if not numeros_odoo:
        # Este SUBCAT no existe en ODOO → SKU válido tal cual
        return {"conflicto": False, "skus_odoo_match": [], "sku_ajustado": sku_propuesto}

    max_num_odoo = max(numeros_odoo)

    if num_propuesto <= max_num_odoo:
        # El número propuesto ya está ocupado o es menor al máximo → ajustar
        nuevo_num   = max_num_odoo + 1
        prefijo     = f"{subcat}-{num_str}-"
        coincidencias = [s for s in skus_odoo if s.upper().startswith(prefijo.upper())]
        st.session_state.setdefault("sku_contadores", {})[subcat] = nuevo_num
        sku_ajustado = f"{subcat}-{nuevo_num:04d}-{atrib}"
        return {"conflicto": True, "skus_odoo_match": coincidencias, "sku_ajustado": sku_ajustado}

    # El número propuesto es mayor al máximo de ODOO → sin conflicto
    return {"conflicto": False, "skus_odoo_match": [], "sku_ajustado": sku_propuesto}


# ══════════════════════════════════════════════════════════════════════════════
# AGENTE SKU — Extracción de imágenes + Claude Vision + generación de SKU
# ══════════════════════════════════════════════════════════════════════════════

def extraer_imagenes_excel(file_bytes: bytes) -> dict[int, dict]:
    """
    Extrae imágenes flotantes del Excel leyendo el XML de posicionamiento.
    Devuelve dict {row_0indexed: {data, ext}} donde row_0indexed es la fila
    de inicio del ancla de la imagen (0 = fila 1 de Excel = encabezado,
    1 = fila 2 = primer producto, etc.).
    """
    NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
    NS_A   = "http://schemas.openxmlformats.org/drawingml/2006/main"
    NS_R   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    resultado: dict[int, dict] = {}

    try:
        with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
            nombres = set(z.namelist())

            # Buscar todos los archivos drawing*.xml
            drawing_paths = sorted(
                n for n in nombres
                if re.match(r"xl/drawings/drawing\d+\.xml$", n)
            )

            for drawing_path in drawing_paths:
                # Archivo de relaciones del drawing
                parts      = drawing_path.rsplit("/", 1)
                rels_path  = f"{parts[0]}/_rels/{parts[1]}.rels"
                if rels_path not in nombres:
                    continue

                # rId → ruta en xl/media/
                rels_tree  = ET.fromstring(z.read(rels_path))
                rid_a_ruta = {}
                for rel in rels_tree:
                    rid    = rel.get("Id")
                    target = rel.get("Target", "")
                    if "media" in target:
                        ruta = target.replace("../", "xl/")
                        rid_a_ruta[rid] = ruta

                # Parsear anchors para obtener fila → rId
                drawing_tree = ET.fromstring(z.read(drawing_path))
                for anchor in drawing_tree:
                    # Fila de inicio (from/row) — funciona para twoCellAnchor y oneCellAnchor
                    row_el = anchor.find(f"{{{NS_XDR}}}from/{{{NS_XDR}}}row")
                    if row_el is None:
                        continue
                    row_num = int(row_el.text)

                    # rId de la imagen embebida en el blip
                    blip = anchor.find(
                        f".//{{{NS_XDR}}}pic/{{{NS_XDR}}}blipFill/{{{NS_A}}}blip"
                    )
                    if blip is None:
                        continue
                    rid = blip.get(f"{{{NS_R}}}embed")
                    if not rid or rid not in rid_a_ruta:
                        continue

                    ruta = rid_a_ruta[rid]
                    if ruta not in nombres:
                        continue
                    ext = Path(ruta).suffix.lower().lstrip(".")
                    if ext not in ("png", "jpg", "jpeg", "gif", "bmp", "webp"):
                        continue

                    # Si hay varias imágenes en la misma fila, guardar solo la primera
                    if row_num not in resultado:
                        resultado[row_num] = {"data": z.read(ruta), "ext": ext}

    except Exception:
        pass

    return resultado


def analizar_imagen_claude(image_data: bytes, ext: str, contexto: dict | None = None) -> dict:
    """
    Analiza una imagen de producto con Claude Haiku Vision y devuelve metadatos estructurados.
    contexto: datos del packing list para apoyar el análisis (nombre, material, uso, etc.)
    """
    media_map = {
        "jpg": "image/jpeg", "jpeg": "image/jpeg",
        "png": "image/png",  "gif": "image/gif",
        "webp": "image/webp", "bmp": "image/png",
    }
    media_type = media_map.get(ext, "image/jpeg")

    # Construir sección de contexto — nombre e imagen se consideran juntos para la categoría
    nombre_ctx = (contexto or {}).get("nombre", "")
    ctx_lines  = []
    if contexto:
        if nombre_ctx:
            ctx_lines.append(f"- Nombre del producto: {nombre_ctx}")
        if contexto.get("material"):
            ctx_lines.append(f"- Material: {contexto['material']}")
        if contexto.get("uso"):
            ctx_lines.append(f"- Uso / aplicación: {contexto['uso']}")
        if contexto.get("largo_cm") and contexto.get("ancho_cm") and contexto.get("alto_cm"):
            ctx_lines.append(f"- Dimensiones: {contexto['largo_cm']} × {contexto['ancho_cm']} × {contexto['alto_cm']} cm")

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

    prompt = f"""Analiza esta imagen de un producto de importación y devuelve ÚNICAMENTE este JSON.
TODOS los campos son obligatorios y nunca pueden quedar vacíos.{ctx_bloque}

{{
  "titulo": "nombre comercial corto en español (máx 60 caracteres)",
  "descripcion": "descripción en español destacando material, uso y características principales (máx 120 caracteres)",
  "categoria": "categoría principal en español",
  "subcategoria_cod": "SOLO uno de estos códigos — basa tu elección en el nombre del producto Y en la imagen: Muebles/Hogar→MUE,MES,SIL,CAM,EST,ORG,COM,ESCR,BAÑ,JAR,TV,COC,DEC,ILUM,TEX | Bebés→BEB,CUNA,PAS,CORR,ALIM,HIG,ROBB,SEG,JUG | Juguetes→JUGU,MUN,PEL,VEH,CONS,JUEG,CART,ELEC,CAS,MONT,DEPO,EDU | Moda→ROP,CALZ,ACC | Tecnología→TEC,CEL | Herramientas/Oficina→HERR,OFI | Mascotas→MASC | Varios→LIB,ART,DEP,VIA,VAR",
  "atributo_cod": "SOLO uno de estos códigos (elige según color, material o característica principal): Colores→NEG,BLN,GRI,ROJ,AZL,VER,AMA,ROS,NAR,MOR,CAF,BEI,MUL,PLA,DOR | Tallas→XS,S,M,L,XL,UNI | Materiales→MAD,MET,TEL,CUE | Otros→EST,PRE,ECO,INF,ADU",
  "atributo_desc": "descripción del atributo principal: color, material o característica"
}}
Solo JSON, sin texto adicional."""

    try:
        import anthropic
        client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY", ""))
        img_b64 = base64.b64encode(image_data).decode("utf-8")
        resp = None
        for _intento in range(4):
            try:
                resp = client.messages.create(
                    model="claude-haiku-4-5-20251001",
                    max_tokens=512,
                    messages=[{
                        "role": "user",
                        "content": [
                            {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": img_b64}},
                            {"type": "text",  "text": prompt},
                        ],
                    }],
                )
                break
            except anthropic.RateLimitError:
                if _intento == 3:
                    raise
                time.sleep(15 * (2 ** _intento))  # 15s, 30s, 60s
        texto = resp.content[0].text.strip()
        if texto.startswith("```"):
            partes = texto.split("```")
            texto  = partes[1].lstrip("json").strip() if len(partes) > 1 else texto
        datos = json.loads(texto)
        # Garantizar que ningún campo quede vacío usando el contexto como fallback
        nombre_ctx = (contexto or {}).get("nombre", "Producto")
        datos["titulo"]           = datos.get("titulo")           or nombre_ctx
        datos["descripcion"]      = datos.get("descripcion")      or nombre_ctx
        datos["categoria"]        = datos.get("categoria")        or "Varios"
        datos["subcategoria_cod"] = datos.get("subcategoria_cod") or "VAR"
        datos["atributo_cod"]     = datos.get("atributo_cod")     or "EST"
        datos["atributo_desc"]    = datos.get("atributo_desc")    or "Estándar"
        datos["_error"] = None
        return datos
    except Exception as e:
        nombre_ctx = (contexto or {}).get("nombre", "Producto")
        return {
            "titulo":          nombre_ctx,
            "descripcion":     nombre_ctx,
            "categoria":       "Varios",
            "subcategoria_cod": "VAR",
            "_error": str(e),
            "atributo_cod":    "EST",
            "atributo_desc":   "Estándar",
        }


def generar_sku(subcategoria_cod: str, atributo_cod: str) -> str:
    """Genera un SKU único con formato SUBCAT-####-ATRIBUTO usando contadores en session_state.
    Valida que los códigos estén en las listas definidas; usa VAR/EST como fallback.
    """
    sub = subcategoria_cod.strip().upper() if subcategoria_cod else _SUBCAT_DEFAULT
    att = atributo_cod.strip().upper()     if atributo_cod     else _ATTR_DEFAULT
    if sub not in SUBCATEGORIAS:
        sub = _SUBCAT_DEFAULT
    if att not in ATRIBUTOS:
        att = _ATTR_DEFAULT
    contadores = st.session_state.setdefault("sku_contadores", {})
    contadores[sub] = contadores.get(sub, 0) + 1
    sku = f"{sub}-{contadores[sub]:04d}-{att}"
    if st.session_state.get("modo_prueba"):
        sku += "_test"
    return sku


def _sku_mismo_numero(sub: str, numero: int, atributo_cod: str) -> str:
    """Construye un SKU con el mismo número base que otro producto del grupo variante.
    NO incrementa el contador — solo cambia el atributo/color.
    """
    att = atributo_cod.strip().upper() if atributo_cod else _ATTR_DEFAULT
    if att not in ATRIBUTOS:
        att = _ATTR_DEFAULT
    sku = f"{sub}-{numero:04d}-{att}"
    if st.session_state.get("modo_prueba"):
        sku += "_test"
    return sku


def _sanitizar_nombre_archivo(nombre: str) -> str:
    """Convierte nombre de producto a nombre de archivo seguro."""
    nombre = nombre.strip().lower()
    nombre = re.sub(r"[^\w\s-]", "", nombre, flags=re.UNICODE)
    nombre = re.sub(r"[\s]+", "_", nombre)
    return nombre[:100] or "producto"


def guardar_imagenes_temp(imagenes: dict[int, dict], productos: list[dict]) -> tuple[int, list[str]]:
    """
    Guarda las imágenes extraídas del Excel en IMAGENES_TEMP_PATH.
    Usa el índice del producto como prefijo para garantizar nombres únicos aunque
    varios productos tengan el mismo nombre.
    Almacena en prod["imagen_temp_stem"] el stem del archivo guardado para poder
    encontrarlo después al renombrar con el SKU.
    Devuelve (n_guardadas, errores).
    """
    errores: list[str] = []
    try:
        IMAGENES_TEMP_PATH.mkdir(exist_ok=True)
    except Exception as e:
        return 0, [f"No se pudo crear carpeta temporal: {e}"]

    # Construir lookup: fila_excel_0idx → índice en lista (después de filtrar duplicados)
    fila_a_idx = {p.get("fila_excel_0idx"): i for i, p in enumerate(productos)
                  if p.get("fila_excel_0idx") is not None}

    guardadas = 0
    for row_num, img in imagenes.items():
        prod_idx = fila_a_idx.get(row_num)
        if prod_idx is None:
            continue
        nombre = productos[prod_idx].get("nombre", f"producto_{prod_idx + 1}")
        nombre_archivo = f"{row_num}_{_sanitizar_nombre_archivo(nombre)}"
        ruta = IMAGENES_TEMP_PATH / f"{nombre_archivo}.{img['ext']}"
        try:
            ruta.write_bytes(img["data"])
            productos[prod_idx]["imagen_temp_stem"] = nombre_archivo
            productos[prod_idx]["imagen_temp_path"] = str(ruta)
            guardadas += 1
        except Exception as e:
            errores.append(f"No se pudo guardar imagen de '{nombre}': {e}")
    return guardadas, errores


def renombrar_imagenes_con_sku(productos: list[dict]) -> tuple[int, list[str]]:
    """
    Renombra las imágenes en IMAGENES_TEMP_PATH usando el stem guardado en
    prod["imagen_temp_stem"] → SKU definitivo.
    Devuelve (n_renombradas, errores).
    """
    if not IMAGENES_TEMP_PATH.exists():
        return 0, ["La carpeta de imágenes temporales no existe — no se extrajeron imágenes del Excel."]

    errores: list[str] = []
    renombradas = 0
    for prod in productos:
        stem_buscado = prod.get("imagen_temp_stem")
        sku          = prod.get("sku", "")
        if not stem_buscado or not sku:
            continue
        encontrado = False
        for archivo in IMAGENES_TEMP_PATH.iterdir():
            if archivo.stem == stem_buscado:
                nueva_ruta = archivo.with_name(f"{sku}{archivo.suffix}")
                if nueva_ruta.exists():
                    # Ya existe imagen para este SKU (producto unificado) — eliminar la extra
                    try:
                        archivo.unlink()
                    except Exception:
                        pass
                    prod["imagen_temp_stem"] = sku
                    encontrado = True
                    break
                try:
                    archivo.rename(nueva_ruta)
                    prod["imagen_temp_stem"] = sku  # actualizar por si se llama de nuevo
                    renombradas += 1
                except Exception as e:
                    errores.append(f"No se pudo renombrar imagen '{stem_buscado}' → '{sku}': {e}")
                encontrado = True
                break
        if not encontrado:
            errores.append(f"No se encontró imagen para SKU '{sku}' (buscada: {stem_buscado})")
    return renombradas, errores


def _reportar_subida_imagenes(subidas: list[str], errores: list[str]) -> None:
    """Agrega al chat un resumen de imágenes subidas a productos en ODOO."""
    lineas = []
    if subidas:
        lineas.append(f"**Imágenes subidas a Odoo ({len(subidas)}):**")
        for sku in subidas:
            lineas.append(f"- ✅ `{sku}`")
    if errores:
        lineas.append(f"\n**Problemas al subir imágenes ({len(errores)}):**")
        for e in errores:
            lineas.append(f"- ❌ {e}")
    if lineas:
        st.session_state.chat.append({"role": "assistant", "content": "\n".join(lineas)})


def _reportar_errores_imagenes(errores: list[str], etapa: str) -> None:
    """Agrega errores de imagen al chat y a session_state para que sean visibles."""
    if not errores:
        return
    lineas = "\n".join(f"- {e}" for e in errores)
    msg = f"⚠️ **Problemas en {etapa}:**\n{lineas}"
    st.session_state.setdefault("imagenes_errores", []).extend(errores)
    st.session_state.chat.append({"role": "assistant", "content": msg})




_ATRIB_COLORES   = {"NEG","BLN","GRI","ROJ","AZL","VER","AMA","ROS","NAR","MOR","CAF","BEI","MUL","PLA","DOR"}
_ATRIB_TALLAS    = {"XS","S","M","L","XL","UNI"}
_ATRIB_MATERIALES= {"MAD","MET","TEL","CUE"}

def _crear_producto_en_odoo(prod: dict, image_data: bytes,
                            uid, models, db: str, pw: str,
                            cat_cache: dict,
                            tipo_cambio: float = 19.0,
                            costo_por_m3: float = 0.0) -> str | None:
    """
    Crea un producto nuevo en ODOO con todos sus datos e imagen.
    cat_cache: dict mutable {nombre_cat: cat_id} para evitar lookups repetidos.
    Devuelve mensaje de error o None si fue exitoso.
    """
    # TODO: mientras se prueba, forzar categoría PRUEBAS_AGENTE y sufijo _test en nombre
    _CAT_FORZADA = "PRUEBAS_AGENTE"
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

        # ── Cálculo de precios y costo ────────────────────────────────────────
        precio_usd     = _safe_float(prod.get("precio_usd"))
        precio_mxn     = round(precio_usd * tipo_cambio, 2)
        cbm_pz         = _cbm_por_pieza(prod)
        costo_cbm_pz   = round(cbm_pz * costo_por_m3, 2)
        costo_unitario = round(precio_mxn + costo_cbm_pz, 2)

        # ── Descripción enriquecida ───────────────────────────────────────────
        def _es_chino(txt: str) -> bool:
            return bool(re.search(r'[\u4e00-\u9fff]', txt or ""))

        def _traducir_campo(txt: str) -> str:
            """Traduce al español si contiene caracteres chinos, usando Claude Haiku."""
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
        # Dimensiones en descripción de venta
        largo, ancho, alto = prod.get("largo_cm"), prod.get("ancho_cm"), prod.get("alto_cm")
        if any([largo, ancho, alto]):
            dims = " × ".join(f"{v}" for v in [largo, ancho, alto] if v)
            desc_partes.append(f"Dimensiones: {dims} cm")

        # ── Notas internas (descripción técnica) ─────────────────────────────
        notas = []
        if prod.get("nombre_alt"):
            notas.append(f"Nombre alternativo: {prod['nombre_alt']}")
        if prod.get("id_guia"):
            notas.append(f"ID Guía / Referencia: {prod['id_guia']}")
        if prod.get("cajas_master"):
            notas.append(f"Cajas master: {int(prod['cajas_master']) if str(prod['cajas_master']).replace('.','').isdigit() else prod['cajas_master']}")
        if prod.get("piezas_total"):
            notas.append(f"Piezas totales en contenedor: {int(prod['piezas_total']) if str(prod['piezas_total']).replace('.','').isdigit() else prod['piezas_total']}")
        if prod.get("cbm_por_pieza"):
            notas.append(f"CBM por pieza: {prod['cbm_por_pieza']}")
        if prod.get("cbm_master_carton"):
            notas.append(f"CBM master carton: {prod['cbm_master_carton']}")
        if prod.get("cbm_total_sku"):
            notas.append(f"CBM total SKU: {prod['cbm_total_sku']}")
        if prod.get("cbm_inner_carton"):
            notas.append(f"CBM inner carton: {prod['cbm_inner_carton']}")
        if any([largo, ancho, alto]):
            notas.append(f"Dimensiones caja: {dims} cm")
        notas.append(f"Costo flete CBM/pieza: ${costo_cbm_pz:.4f} MXN")

        img_b64 = base64.b64encode(image_data).decode("utf-8")

        # Volumen por pieza en m³ (CBM por pieza)
        volumen_pz = cbm_pz if cbm_pz > 0 else None

        vals = {
            "name":             nombre_prod,
            "default_code":     prod.get("sku", ""),
            "description_sale": "\n".join(desc_partes),
            "description":      "\n".join(notas) if notas else False,
            "categ_id":         cat_id,
            "list_price":       precio_mxn,       # precio de venta en MXN
            "standard_price":   costo_unitario,   # costo calculado (precio_mxn + flete CBM)
            "type":             "product",
            "sale_ok":          True,
            "purchase_ok":      True,
            "image_1920":       img_b64,
        }
        if volumen_pz is not None:
            vals["volume"] = volumen_pz
        prod_id = models.execute_kw(db, uid, pw, "product.template", "create", [vals])

        # ── Forzar standard_price en product.product (evita que AVCO/FIFO lo ignore) ──
        if prod_id and costo_unitario > 0:
            try:
                pp_ids = models.execute_kw(db, uid, pw, "product.product", "search",
                                           [[["product_tmpl_id", "=", prod_id]]])
                if pp_ids:
                    models.execute_kw(db, uid, pw, "product.product", "write",
                                      [pp_ids, {"standard_price": costo_unitario}])
            except Exception:
                pass  # No crítico — el producto ya fue creado

        # ── Atributo del producto (color / talla / material / tipo) ──────────
        atrib_cod = (prod.get("atributo_cod") or "EST").strip().upper()
        atrib_val = ATRIBUTOS.get(atrib_cod)
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
                attr_id = attr_ids[0] if attr_ids else models.execute_kw(
                    db, uid, pw, "product.attribute", "create", [{"name": atrib_nombre}])
                val_ids = models.execute_kw(db, uid, pw, "product.attribute.value", "search",
                                            [[["name", "=", atrib_val],
                                              ["attribute_id", "=", attr_id]]])
                val_id = val_ids[0] if val_ids else models.execute_kw(
                    db, uid, pw, "product.attribute.value", "create",
                    [{"name": atrib_val, "attribute_id": attr_id}])
                models.execute_kw(db, uid, pw, "product.template.attribute.line", "create", [{
                    "product_tmpl_id": prod_id,
                    "attribute_id":    attr_id,
                    "value_ids":       [[4, val_id]],
                }])
            except Exception:
                pass  # Atributo no crítico — producto ya fue creado

        return None
    except Exception as e:
        return f"Error al crear producto '{prod.get('sku', '?')}' en ODOO: {e}"




def _subir_productos_a_odoo(productos: list[dict],
                             tipo_cambio: float,
                             costo_por_m3: float) -> tuple[list[str], list[str]]:
    """
    Sube imágenes a productos YA EXISTENTES en Odoo (buscados por SKU).
    No crea productos nuevos — la creación ocurrió al confirmar la clasificación.
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
        common = xmlrpc.client.ServerProxy(f"{odoo_url}/xmlrpc/2/common", allow_none=True)
        uid    = common.authenticate(odoo_db, odoo_user, odoo_pass, {})
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
            # Buscar producto existente por SKU (default_code en product.product o product.template)
            tmpl_ids = models_proxy.execute_kw(
                odoo_db, uid, odoo_pass, "product.template", "search",
                [[["default_code", "=", sku]]])
            if not tmpl_ids:
                # Buscar también en product.product (variante con ese SKU)
                var_ids = models_proxy.execute_kw(
                    odoo_db, uid, odoo_pass, "product.product", "search",
                    [[["default_code", "=", sku]]])
                if var_ids:
                    var_data = models_proxy.execute_kw(
                        odoo_db, uid, odoo_pass, "product.product", "read",
                        [var_ids[:1]], {"fields": ["product_tmpl_id"]})
                    tmpl_ids = [var_data[0]["product_tmpl_id"][0]] if var_data else []

            if not tmpl_ids:
                errores.append(f"SKU `{sku}` no encontrado en Odoo — imagen no subida")
                continue

            img_b64 = base64.b64encode(image_data).decode("utf-8")
            models_proxy.execute_kw(
                odoo_db, uid, odoo_pass, "product.template", "write",
                [tmpl_ids[:1], {"image_1920": img_b64}])
            subidas.append(sku)

        except Exception as e:
            errores.append(f"Error al subir imagen para `{sku}`: {e}")

    return subidas, errores


def enriquecer_con_imagenes(file_bytes: bytes, productos: list[dict]) -> tuple[list[dict], int, list[dict]]:
    """
    Extrae imágenes del Excel, analiza cada una con Claude Haiku Vision,
    genera SKU validado contra ODOO, busca similares por imagen/nombre y enriquece los productos.

    Devuelve (productos_enriquecidos, n_procesadas, conflictos).
    """
    imagenes       = extraer_imagenes_excel(file_bytes)
    skus_odoo      = st.session_state.get("odoo_skus", [])
    prods_odoo     = st.session_state.get("odoo_productos", [])
    phashes_odoo   = st.session_state.get("odoo_phashes", {})

    if skus_odoo:
        sincronizar_contadores_con_odoo(skus_odoo)

    if not imagenes:
        return productos, 0, []

    # Guardar imágenes en carpeta temporal con nombre del producto
    _, errores_guardado = guardar_imagenes_temp(imagenes, productos)
    if errores_guardado:
        st.session_state.setdefault("imagenes_errores", []).extend(errores_guardado)

    procesadas  = 0
    conflictos  = []
    errores_gemini: list[str] = []
    skus_ya_en_conflicto: set[int] = set()
    # Grupos de variantes: gv_id → (sub, numero) del primer producto procesado
    grupos_variante_sku: dict[str, tuple[str, int]] = {}

    for i, prod in enumerate(productos):
        # Usar fila_excel_0idx para buscar la imagen — coincide con las claves de
        # extraer_imagenes_excel (0-indexed desde el XML del Excel).
        # NO usar i+1 porque si el header no está en fila 1 los índices desfasan.
        row_num = prod.get("fila_excel_0idx")
        if row_num is None:
            row_num = i + 1  # fallback para productos sin fila conocida (ej. Tab 2)
        img = imagenes.get(row_num)
        if img is None:
            prod.setdefault("sin_imagen", True)
            prod.setdefault("fila_sin_imagen", row_num + 1)
            continue

        contexto = {
            "nombre":   prod.get("nombre"),
            "material": prod.get("material"),
            "uso":      prod.get("uso"),
            "largo_cm": prod.get("largo_cm"),
            "ancho_cm": prod.get("ancho_cm"),
            "alto_cm":  prod.get("alto_cm"),
        }
        datos = analizar_imagen_claude(img["data"], img["ext"], contexto)
        if datos.get("_error"):
            errores_gemini.append(f"Fila {row_num + 1} ({prod.get('nombre', '?')}): {datos['_error']}")

        sub_cod = datos.get("subcategoria_cod", "VAR")
        att_cod = datos.get("atributo_cod", "EST")
        gv = prod.get("_grupo_variante")
        if gv and gv in grupos_variante_sku:
            # Variante de un grupo ya numerado → mismo sub y número, solo cambia atributo
            base_sub, base_num = grupos_variante_sku[gv]
            sku_inicial = _sku_mismo_numero(base_sub, base_num, att_cod)
        else:
            sku_inicial = generar_sku(sub_cod, att_cod)
            if gv:
                # Primer producto del grupo → guardar su sub y número para los demás
                partes = sku_inicial.replace("_test", "").split("-")
                if len(partes) >= 2:
                    try:
                        grupos_variante_sku[gv] = (partes[0], int(partes[1]))
                    except (ValueError, IndexError):
                        pass

        conflicto_entry = None

        if skus_odoo:
            # ── 1. Conflicto por prefijo SKU ──────────────────────────────────
            validacion = validar_sku_vs_odoo(sku_inicial, skus_odoo)
            sku_final  = validacion["sku_ajustado"]
            if validacion["conflicto"]:
                conflicto_entry = {
                    "idx":              i,
                    "nombre":           prod.get("nombre", f"Producto {i+1}"),
                    "sku_propuesto":    sku_inicial,
                    "sku_ajustado":     sku_final,
                    "datos_gemini":     datos,
                    "skus_odoo_match":  validacion["skus_odoo_match"],
                    "productos_odoo":   [],
                    "razon":            "sku",
                    "similares_odoo":   [],
                }

            # ── 2. Similitud por imagen, nombre y RAG ────────────────────────
            if prods_odoo:
                similares = buscar_similares_odoo(
                    img["data"], prod.get("nombre", ""),
                    prods_odoo, phashes_odoo,
                )
            else:
                similares = []

            if similares and conflicto_entry is None:
                razones = []
                if any(s.get("por_imagen") for s in similares):
                    razones.append("imagen similar")
                if any(s.get("por_nombre") for s in similares):
                    razones.append("nombre similar")
                conflicto_entry = {
                    "idx":              i,
                    "nombre":           prod.get("nombre", f"Producto {i+1}"),
                    "sku_propuesto":    sku_inicial,
                    "sku_ajustado":     sku_inicial,
                    "datos_gemini":     datos,
                    "skus_odoo_match":  [s["sku"] for s in similares],
                    "productos_odoo":   [],
                    "razon":            ", ".join(razones) if razones else "similitud detectada",
                    "similares_odoo":   similares,
                }
            elif similares and conflicto_entry is not None:
                conflicto_entry["similares_odoo"] = similares

            if conflicto_entry:
                conflictos.append(conflicto_entry)
                skus_ya_en_conflicto.add(i)
        else:
            sku_final = sku_inicial

        prod["sku"]           = sku_final if skus_odoo else sku_inicial
        prod["titulo_imagen"] = datos.get("titulo")       or prod.get("nombre", "")
        prod["descripcion"]   = datos.get("descripcion")  or prod.get("descripcion", "")
        prod["categoria"]     = datos.get("categoria")    or prod.get("categoria", "Varios")
        prod["atributo"]      = datos.get("atributo_desc") or prod.get("atributo", "Estándar")
        prod["atributo_cod"]  = datos.get("atributo_cod")  or "EST"
        procesadas += 1

    # Enriquecer conflictos con detalles completos de ODOO (1 sola llamada en batch)
    if conflictos:
        todos_skus  = list({s for c in conflictos for s in c["skus_odoo_match"]})
        detalles    = cargar_detalle_productos_odoo(todos_skus)
        detalle_map = {p["default_code"]: p for p in detalles}
        for c in conflictos:
            c["productos_odoo"] = [detalle_map[s] for s in c["skus_odoo_match"] if s in detalle_map]

    # Guardar errores de Gemini en session_state para mostrarlos en UI
    if errores_gemini:
        st.session_state["gemini_errores"] = errores_gemini

    return productos, procesadas, conflictos


def generar_excel(productos: list[dict], tipo_cambio: float, contenedor: str) -> bytes:
    _no_fill   = PatternFill(fill_type=None)
    _no_border = Border(
        left=Side(style=None), right=Side(style=None),
        top=Side(style=None),  bottom=Side(style=None),
    )
    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws = wb.active
    ultima = 3 + len(productos) - 1

    for r in range(3, ultima + 1):
        for c in range(1, ws.max_column + 1):
            ws.cell(r, c).value = None
    for r in range(ultima + 1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(r, c)
            cell.value = None
            cell.fill = _no_fill
            cell.border = _no_border
    if "Sheet2" in wb.sheetnames:
        del wb["Sheet2"]

    # Eliminar columnas largo, ancho, alto (originalmente cols 4, 5, 6)
    ws.delete_cols(4, 3)

    # Actualizar encabezados tras el delete
    ws.cell(2, 3).value  = "Imagen"
    ws.cell(2, 13).value = "Costo Landed USD"

    # Tamaño fijo para la columna de imagen (col 3) y filas de datos
    IMG_W_PX  = 120   # ancho imagen en píxeles
    IMG_H_PX  = 120   # alto imagen en píxeles
    IMG_COL_W = 18    # ancho columna 3 en unidades Excel (~120px)
    ROW_H_PT  = 90    # altura de fila en puntos (~120px)
    ws.column_dimensions["C"].width = IMG_COL_W

    # Después del delete_cols las columnas quedan:
    # 1:nombre  2:sku  3:imagen  4:cbm_por_pieza  5:precio_usd  6:None
    # 7:piezas_x_caja  8:cbm_caja  9:tipo_producto  10:EMPRESA  11:None
    # 12:contenedor  13:costo_landed_usd  14:descripcion  15:categoria
    # 16:atributo  17-20:None
    for i, prod in enumerate(productos):
        r   = i + 3
        sku = prod.get("sku", "")
        ws.cell(r, 1).value  = prod.get("nombre")
        ws.cell(r, 2).value  = sku

        # Imagen incrustada en col 3
        _img_incrustada = False
        if sku and IMAGENES_TEMP_PATH.exists():
            _archivo_img = next(
                (f for f in IMAGENES_TEMP_PATH.iterdir() if f.stem == sku),
                None,
            )
            if _archivo_img and _PILLOW_OK:
                try:
                    pil_img = PILImage.open(_archivo_img).convert("RGB")
                    pil_img.thumbnail((IMG_W_PX, IMG_H_PX), PILImage.LANCZOS)
                    _buf_img = io.BytesIO()
                    pil_img.save(_buf_img, format="PNG")
                    _buf_img.seek(0)
                    xl_img = XLImage(_buf_img)
                    xl_img.width  = IMG_W_PX
                    xl_img.height = IMG_H_PX
                    xl_img.anchor = f"C{r}"
                    ws.add_image(xl_img)
                    ws.row_dimensions[r].height = ROW_H_PT
                    _img_incrustada = True
                except Exception:
                    pass
        if not _img_incrustada:
            ws.cell(r, 3).value = ""
        cbm = ws.cell(r, 4)
        _cu = _cbm_por_pieza(prod)
        cbm.value = _cu if _cu > 0 else None
        if cbm.value is not None:
            cbm.number_format = "0.000000"
        ws.cell(r, 5).value  = prod.get("precio_usd")
        ws.cell(r, 6).value  = prod.get("piezas_total")   # col "Unidades"
        ws.cell(r, 7).value  = prod.get("piezas_x_caja")
        cbm_caja = ws.cell(r, 8)
        px = prod.get("piezas_x_caja")
        cbm_caja.value = round(float(px) * _cu, 6) if px and _cu > 0 else None
        if cbm_caja.value is not None:
            cbm_caja.number_format = "0.000000"
        ws.cell(r, 9).value  = prod.get("tipo_producto", "Producto almacenable")
        ws.cell(r, 10).value = EMPRESA
        ws.cell(r, 11).value = None
        ws.cell(r, 12).value = contenedor
        ws.cell(r, 13).value = float(prod["precio_usd"]) if prod.get("precio_usd") else None
        ws.cell(r, 14).value = prod.get("descripcion")
        ws.cell(r, 15).value = prod.get("categoria")
        ws.cell(r, 16).value = prod.get("atributo")
        ws.cell(r, 17).value = None
        ws.cell(r, 18).value = None
        ws.cell(r, 19).value = None
        ws.cell(r, 20).value = None

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def _cbm_total_fila(prod: dict) -> float:
    """
    Deriva el CBM total de una fila según lo que traiga el packing list.
    Prioridad: cbm_total_sku > cbm_master_carton × cajas > cbm_por_pieza × piezas_total.
    """
    cbm_pz  = _safe_float(prod.get("cbm_por_pieza"))
    cbm_mc  = _safe_float(prod.get("cbm_master_carton"))
    cbm_tot = _safe_float(prod.get("cbm_total_sku"))
    cajas   = _safe_float(prod.get("cajas_master"))
    pzs_caja= _safe_float(prod.get("piezas_x_caja"))
    pzs_tot = _safe_float(prod.get("piezas_total")) or pzs_caja

    if cbm_tot > 0:
        return cbm_tot
    if cbm_mc > 0:
        return cbm_mc * cajas if cajas > 0 else cbm_mc * (pzs_tot / pzs_caja if pzs_caja > 0 else 1)
    if cbm_pz > 0:
        return cbm_pz * pzs_tot
    return 0.0


def _cbm_por_pieza(prod: dict) -> float:
    """Deriva CBM por pieza individual desde lo que traiga el packing list."""
    cbm_pz = _safe_float(prod.get("cbm_por_pieza"))
    cbm_mc = _safe_float(prod.get("cbm_master_carton"))
    cbm_tot= _safe_float(prod.get("cbm_total_sku"))
    pzs_caja= _safe_float(prod.get("piezas_x_caja"))
    pzs_tot = _safe_float(prod.get("piezas_total")) or pzs_caja

    if cbm_pz > 0:
        return cbm_pz
    if cbm_mc > 0 and pzs_caja > 0:
        return cbm_mc / pzs_caja
    if cbm_tot > 0 and pzs_tot > 0:
        return cbm_tot / pzs_tot
    return 0.0


def generar_excel_master(productos: list[dict], tipo_cambio: float,
                         costo_contenedor: float, nombre_packing: str) -> bytes:
    """
    Genera el Excel maestro de costos con fórmulas visibles.
    Columnas (igual que ejemplo_excel2.csv, más SKU):
      A  Packing List
      B  SKU
      C  Producto
      D  Cajas
      E  Piezas Total
      F  Piezas x Caja
      G  CBM Total          ← cbm_total_sku (o derivado)
      H  CBM Master Carton
      I  CBM Inner Carton
      J  CBM por Pieza      ← cbm_por_pieza (o derivado)
      K  Tipo de Cambio     ← =$U$2
      L  Precio USD
      M  Precio MXN         ← =L*K
      N  Costo Contenedor   ← =$U$3
      O  Costo por M³       ← =$U$5
      P  Costo CBM/Pieza    ← =J*O
      Q  Costo Unitario     ← =M+P
      R  Costo Total        ← =Q*E

    Parámetros (cols T-U, editables):
      U2 = Tipo de cambio
      U3 = Costo contenedor
      U4 = CBM Contenedor   ← =SUM(G_datos)
      U5 = Costo por M³     ← =U3/U4
    """
    n = len(productos)
    DATA_START = 3
    DATA_END   = DATA_START + n - 1
    TOTAL_ROW  = DATA_END + 1

    # Parámetros en columnas T=20, U=21
    T, U = 20, 21
    Tc = get_column_letter(T)   # "T"
    Uc = get_column_letter(U)   # "U"

    REF_TC  = f"${Uc}$2"   # Tipo de cambio
    REF_CC  = f"${Uc}$3"   # Costo contenedor
    REF_CBM = f"${Uc}$4"   # CBM Contenedor total
    REF_M3  = f"${Uc}$5"   # Costo por M³

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Master Costos"

    # ── Estilos compartidos ───────────────────────────────────────────────────
    # Grupos de color por tipo de columna
    _col_color = {}
    for color, cols in [
        ("1F3864", [1, 2, 3, 4, 5, 6]),   # identificación — azul oscuro
        ("1B6B3A", [7, 8, 9, 10]),          # CBM — verde oscuro
        ("2E4057", [11, 12]),               # parámetros precio — azul medio
        ("7B2D00", [13, 14, 15, 16, 17, 18]),  # costos — rojo oscuro
    ]:
        for c in cols:
            _col_color[c] = color

    HEADERS = [
        "Packing List",       # A  1
        "SKU",                # B  2
        "Producto",           # C  3
        "Cajas",              # D  4
        "Piezas Total",       # E  5
        "Piezas x Caja",      # F  6
        "CBM Total",          # G  7
        "CBM Master Carton",  # H  8
        "CBM Inner Carton",   # I  9
        "CBM por Pieza",      # J  10
        "Tipo de Cambio",     # K  11
        "Precio USD",         # L  12
        "Precio MXN",         # M  13
        "Costo Contenedor",   # N  14
        "Costo por M³",       # O  15
        "Costo CBM/Pieza",    # P  16
        "Costo Unitario",     # Q  17
        "Costo Total",        # R  18
    ]
    COL_WIDTHS = [26, 16, 34, 8, 13, 13, 14, 18, 16, 14, 13, 12, 14, 16, 14, 15, 14, 14]

    FMT_MXN  = '"$"#,##0.00'
    FMT_USD  = '#,##0.00'
    FMT_CBM  = '0.000000'
    FMT_CBM4 = '0.0000'
    FMT_INT  = '#,##0'
    FMT_NUM  = '#,##0.00'

    COL_FMT = {
        4:  FMT_INT,   # Cajas
        5:  FMT_INT,   # Piezas Total
        6:  FMT_INT,   # Piezas x Caja
        7:  FMT_CBM,   # CBM Total
        8:  FMT_CBM,   # CBM Master Carton
        9:  FMT_CBM,   # CBM Inner Carton
        10: FMT_CBM,   # CBM por Pieza
        11: FMT_NUM,   # Tipo de Cambio
        12: FMT_USD,   # Precio USD
        13: FMT_MXN,   # Precio MXN
        14: FMT_MXN,   # Costo Contenedor
        15: FMT_MXN,   # Costo por M³
        16: FMT_MXN,   # Costo CBM/Pieza
        17: FMT_MXN,   # Costo Unitario
        18: FMT_MXN,   # Costo Total
    }

    FILL_ODD        = PatternFill("solid", fgColor="F0F4FA")
    FILL_EVEN       = PatternFill("solid", fgColor="FFFFFF")
    FILL_HDR        = PatternFill("solid", fgColor="0D1F3C")
    FILL_TOTAL      = PatternFill("solid", fgColor="0D1F3C")
    FILL_PARAM_LBL  = PatternFill("solid", fgColor="2E4057")
    FILL_PARAM_VAL  = PatternFill("solid", fgColor="F0F4FA")
    FILL_PARAM_CALC = PatternFill("solid", fgColor="E8F5E9")

    FONT_TITLE = Font(name="Calibri", bold=True, size=13, color="FFFFFF")
    FONT_HDR   = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
    FONT_DATA  = Font(name="Calibri", size=10)
    FONT_TOTAL = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
    FONT_PLBL  = Font(name="Calibri", bold=True, size=9,  color="FFFFFF")
    FONT_PVAL  = Font(name="Calibri", size=10)
    FONT_PCALC = Font(name="Calibri", italic=True, size=10, color="1B6B3A")

    ALIGN_CTR = Alignment(horizontal="center", vertical="center")
    ALIGN_LFT = Alignment(horizontal="left",   vertical="center")
    ALIGN_RGT = Alignment(horizontal="right",  vertical="center")

    border_hdr  = Border(left=Side(style="thin",   color="FFFFFF"),
                         right=Side(style="thin",  color="FFFFFF"),
                         bottom=Side(style="medium", color="FFFFFF"))
    border_data = Border(bottom=Side(style="thin", color="D9D9D9"),
                         right=Side(style="thin",  color="D9D9D9"))

    # ── Fila 1: Título ────────────────────────────────────────────────────────
    ws.merge_cells(f"A1:{get_column_letter(len(HEADERS))}1")
    c = ws["A1"]
    c.value     = f"MASTER COSTOS — {nombre_packing}"
    c.font      = FONT_TITLE
    c.fill      = FILL_HDR
    c.alignment = ALIGN_CTR
    ws.row_dimensions[1].height = 22

    # ── Fila 2: Encabezados de datos ──────────────────────────────────────────
    for ci, header in enumerate(HEADERS, start=1):
        c = ws.cell(row=2, column=ci, value=header)
        c.fill      = PatternFill("solid", fgColor=_col_color.get(ci, "2E4057"))
        c.font      = FONT_HDR
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border    = border_hdr
    ws.row_dimensions[2].height = 30

    for ci, ancho in enumerate(COL_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(ci)].width = ancho

    # ── Sección PARÁMETROS (cols T-U, filas 1-6) ─────────────────────────────
    ws.column_dimensions[Tc].width = 26
    ws.column_dimensions[Uc].width = 18

    ws.merge_cells(f"{Tc}1:{Uc}1")
    c = ws.cell(row=1, column=T, value="⚙ PARÁMETROS")
    c.font = FONT_TITLE; c.fill = PatternFill("solid", fgColor="1F3864")
    c.alignment = ALIGN_CTR

    params = [
        ("Tipo de cambio (MXN/USD)",  tipo_cambio,       FMT_NUM,  FILL_PARAM_VAL,  FONT_PVAL,  False),
        ("Costo contenedor (MXN)",    costo_contenedor,  FMT_MXN,  FILL_PARAM_VAL,  FONT_PVAL,  False),
        ("CBM Contenedor (M³)",
         f"=SUM(G{DATA_START}:G{DATA_END})",
         FMT_CBM4,  FILL_PARAM_CALC, FONT_PCALC, True),
        ("Costo por M³ (MXN)",
         f"={Uc}3/{Uc}4",
         FMT_MXN,   FILL_PARAM_CALC, FONT_PCALC, True),
    ]
    for pi, (label, value, fmt, fill_v, font_v, _) in enumerate(params, start=2):
        lc = ws.cell(row=pi, column=T, value=label)
        lc.font = FONT_PLBL; lc.fill = FILL_PARAM_LBL
        lc.alignment = ALIGN_LFT
        lc.border = Border(bottom=Side(style="thin", color="FFFFFF"),
                           right=Side(style="thin",  color="FFFFFF"))
        vc = ws.cell(row=pi, column=U, value=value)
        vc.font = font_v; vc.fill = fill_v
        vc.alignment = ALIGN_RGT
        vc.number_format = fmt
        vc.border = Border(bottom=Side(style="thin", color="BBBBBB"))

    nota_row = 6
    ws.merge_cells(f"{Tc}{nota_row}:{Uc}{nota_row}")
    nc = ws.cell(row=nota_row, column=T,
                 value="U4 y U5 se calculan con fórmula — editables U2 y U3")
    nc.font      = Font(name="Calibri", italic=True, size=8, color="666666")
    nc.alignment = ALIGN_CTR

    # ── Filas de datos con fórmulas ───────────────────────────────────────────
    for row_num, prod in enumerate(productos, start=DATA_START):
        r = row_num

        pzs_caja  = _safe_float(prod.get("piezas_x_caja"))
        pzs_tot   = _safe_float(prod.get("piezas_total")) or pzs_caja
        cajas     = _safe_float(prod.get("cajas_master"))
        precio    = _safe_float(prod.get("precio_usd"))
        cbm_pz    = _safe_float(prod.get("cbm_por_pieza"))
        cbm_mc    = _safe_float(prod.get("cbm_master_carton"))
        cbm_ic    = _safe_float(prod.get("cbm_inner_carton"))
        cbm_tot   = _safe_float(prod.get("cbm_total_sku"))

        # Derivar cbm_por_pieza y cbm_total según disponibilidad
        if cbm_pz > 0:
            cbm_por_pieza_val  = cbm_pz
            cbm_total_fila_val = cbm_tot if cbm_tot > 0 else round(cbm_pz * pzs_tot, 6)
        elif cbm_mc > 0 and pzs_caja > 0:
            cbm_por_pieza_val  = round(cbm_mc / pzs_caja, 6)
            cbm_total_fila_val = cbm_tot if cbm_tot > 0 else (
                round(cbm_mc * cajas, 6) if cajas > 0 else round(cbm_por_pieza_val * pzs_tot, 6)
            )
        elif cbm_tot > 0 and pzs_tot > 0:
            cbm_por_pieza_val  = round(cbm_tot / pzs_tot, 6)
            cbm_total_fila_val = cbm_tot
        else:
            cbm_por_pieza_val  = 0
            cbm_total_fila_val = 0

        # Valores directos del packing list (columnas A-L)
        valores_directos = {
            1:  nombre_packing,                                              # A Packing List
            2:  prod.get("sku", ""),                                         # B SKU
            3:  prod.get("nombre", ""),                                      # C Producto
            4:  int(cajas)    if cajas    == int(cajas)    else cajas,       # D Cajas
            5:  int(pzs_tot)  if pzs_tot  == int(pzs_tot)  else pzs_tot,   # E Piezas Total
            6:  int(pzs_caja) if pzs_caja == int(pzs_caja) else pzs_caja,  # F Piezas x Caja
            7:  round(cbm_total_fila_val, 6) or None,                       # G CBM Total
            8:  round(cbm_mc, 6) or None,                                   # H CBM Master Carton
            9:  round(cbm_ic, 6) or None,                                   # I CBM Inner Carton
            10: round(cbm_por_pieza_val, 6) or None,                        # J CBM por Pieza
            # K-R: fórmulas
            12: precio if precio > 0 else None,                             # L Precio USD
        }

        # Fórmulas (columnas K, M-R)
        formulas = {
            11: f"={REF_TC}",                              # K Tipo de Cambio
            13: f"=IF(L{r}=\"\",0,L{r}*{REF_TC})",        # M Precio MXN
            14: f"={REF_CC}",                              # N Costo Contenedor
            15: f"={REF_M3}",                              # O Costo por M³
            16: f"=J{r}*{REF_M3}",                        # P Costo CBM/Pieza
            17: f"=M{r}+P{r}",                             # Q Costo Unitario
            18: f"=Q{r}*E{r}",                             # R Costo Total
        }

        fill = FILL_ODD if row_num % 2 == 1 else FILL_EVEN
        ws.row_dimensions[r].height = 18

        for ci in range(1, 19):
            valor = formulas.get(ci) if ci in formulas else valores_directos.get(ci)
            cell  = ws.cell(row=r, column=ci, value=valor)
            cell.fill      = fill
            cell.font      = FONT_DATA
            cell.border    = border_data
            cell.alignment = ALIGN_LFT if ci == 3 else ALIGN_CTR
            if ci in COL_FMT:
                cell.number_format = COL_FMT[ci]

    # ── Fila de totales ───────────────────────────────────────────────────────
    ws.row_dimensions[TOTAL_ROW].height = 20
    totales_vals = {
        1:  "TOTALES",
        5:  f"=SUM(E{DATA_START}:E{DATA_END})",   # Piezas Total
        7:  f"={REF_CBM}",                         # CBM Total = CBM Contenedor
        18: f"=SUM(R{DATA_START}:R{DATA_END})",   # Costo Total
    }
    totales_fmt = {5: FMT_INT, 7: FMT_CBM4, 18: FMT_MXN}

    for ci in range(1, 19):
        cell = ws.cell(row=TOTAL_ROW, column=ci, value=totales_vals.get(ci, ""))
        cell.fill      = FILL_TOTAL
        cell.font      = FONT_TOTAL
        cell.alignment = ALIGN_LFT if ci == 1 else ALIGN_CTR
        if ci in totales_fmt:
            cell.number_format = totales_fmt[ci]

    ws.freeze_panes = f"A{DATA_START}"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


PURCHASE_TEMPLATE_PATH = Path(__file__).parent / "Purchase.xlsx"


def generar_excel_purchase(productos: list[dict], tipo_cambio: float,
                           costo_contenedor: float) -> bytes:
    """
    Genera el Excel de Purchase Order a partir del template Purchase.xlsx.
    Llena: PRODUCT (SKU), QUANTITY (piezas_total), DESCRIPTION, PRICE (costo_unitario).
    Las columnas fijas (PURCHASE ID, SUPPLIER, CURRENCY, UOM, DATE) se toman del template.
    """
    # Pre-calcular costo_por_m³ igual que en generar_excel_master
    cbm_total    = sum(_cbm_total_fila(p) for p in productos)
    costo_por_m3 = costo_contenedor / cbm_total if cbm_total > 0 else 0.0

    wb = openpyxl.load_workbook(PURCHASE_TEMPLATE_PATH)
    ws = wb.active

    # Leer valores fijos de la fila de ejemplo (fila 2) del template
    _purchase_id = ws.cell(2, 1).value or ""
    _supplier    = ws.cell(2, 2).value or EMPRESA
    _currency    = ws.cell(2, 3).value or "MXN"
    _uom         = ws.cell(2, 6).value or "Unidades"
    _date        = ws.cell(2, 10).value

    # Limpiar fila de ejemplo
    for c in range(1, 11):
        ws.cell(2, c).value = None

    for i, prod in enumerate(productos):
        r = 2 + i

        precio_usd     = _safe_float(prod.get("precio_usd"))
        precio_mxn     = round(precio_usd * tipo_cambio, 2)
        cbm_pz         = _cbm_por_pieza(prod)
        costo_cbm_pz   = round(cbm_pz * costo_por_m3, 2)
        costo_unitario = round(precio_mxn + costo_cbm_pz, 2)

        piezas_total = prod.get("piezas_total")
        try:
            piezas_total = int(float(piezas_total)) if piezas_total else None
        except (ValueError, TypeError):
            piezas_total = None

        ws.cell(r, 1).value  = _purchase_id
        ws.cell(r, 2).value  = _supplier
        ws.cell(r, 3).value  = _currency
        ws.cell(r, 4).value  = prod.get("sku", "")           # PRODUCT
        ws.cell(r, 5).value  = piezas_total                  # QUANTITY
        ws.cell(r, 6).value  = _uom
        ws.cell(r, 7).value  = prod.get("descripcion", "")   # DESCRIPTION
        ws.cell(r, 8).value  = costo_unitario if costo_unitario > 0 else None  # PRICE
        ws.cell(r, 9).value  = None
        ws.cell(r, 10).value = _date

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ══════════════════════════════════════════════════════════════════════════════
# Clasificación padre / variante — flujo unificado
# ══════════════════════════════════════════════════════════════════════════════

def _generar_sku_padre(subcategoria_cod: str) -> str:
    """Genera SKU de producto padre: SUBCAT-NNNN (sin atributo)."""
    sub = subcategoria_cod.strip().upper() if subcategoria_cod else _SUBCAT_DEFAULT
    if sub not in SUBCATEGORIAS:
        sub = _SUBCAT_DEFAULT
    contadores = st.session_state.setdefault("sku_contadores", {})
    contadores[sub] = contadores.get(sub, 0) + 1
    sku = f"{sub}-{contadores[sub]:04d}"
    if st.session_state.get("modo_prueba"):
        sku += "_test"
    return sku


def _atributo_cod_desde_valor(atributo_valor: str | None) -> str | None:
    """Mapea atributo_valor al código de SKU del dict ATRIBUTOS. Fallback: 3 letras."""
    if not atributo_valor:
        return None
    val_lower = atributo_valor.lower().strip()
    for cod, desc in ATRIBUTOS.items():
        if desc.lower() == val_lower or cod.lower() == val_lower:
            return cod
    return atributo_valor[:3].upper()


def _extraer_nombre_base_atributo_batch(productos: list[dict]) -> list[dict]:
    """
    Llama a Claude Haiku una sola vez para extraer nombre_base, atributo_tipo
    y atributo_valor de todos los productos.
    Devuelve lista de dicts en el mismo orden que 'productos'.
    """
    if not productos:
        return []
    lineas = []
    for i, p in enumerate(productos):
        nombre = p.get("nombre") or p.get("titulo") or f"Producto {i+1}"
        lineas.append(f"{i}: {nombre}")
    prompt = (
        "Analiza estos nombres de productos de una importación de contenedor.\n"
        "Para cada uno extrae:\n"
        "- nombre_base: nombre base SIN color, talla o material específico\n"
        "- atributo_tipo: \"Color\", \"Talla\", \"Material\", o null si no hay atributo distinguible\n"
        "- atributo_valor: valor concreto (\"Azul\", \"XL\", \"Madera\") o null\n\n"
        "IMPORTANTE: Si dos o más productos son el mismo tipo base con distinto color/talla/material "
        "(ej: \"Mesa Negra\" y \"Mesa Blanca\"), deben tener el MISMO nombre_base.\n\n"
        "Responde SOLO con JSON array (mismo orden, sin texto extra):\n"
        "[{\"nombre_base\":\"...\",\"atributo_tipo\":\"...\",\"atributo_valor\":\"...\"}, ...]\n\n"
        "Productos:\n" + "\n".join(lineas)
    )
    try:
        llm  = ChatAnthropic(model="claude-haiku-4-5-20251001", max_tokens=2000)
        resp = llm.invoke([HumanMessage(content=prompt)])
        content = resp.content
        if isinstance(content, list):
            content = " ".join(b.get("text", "") for b in content if isinstance(b, dict))
        content = content.strip()
        start = content.find("[")
        end   = content.rfind("]") + 1
        if start >= 0 and end > start:
            content = content[start:end]
        resultados = json.loads(content)
        while len(resultados) < len(productos):
            resultados.append({"nombre_base": None, "atributo_tipo": None, "atributo_valor": None})
        return resultados[:len(productos)]
    except Exception:
        return [{"nombre_base": None, "atributo_tipo": None, "atributo_valor": None}
                for _ in productos]


def _buscar_padre_en_odoo_por_nombre(nombre_base: str, prods_odoo: list[dict],
                                      phashes_odoo: dict,
                                      imagen_data: bytes | None = None,
                                      umbral: float = 0.75) -> dict | None:
    """
    Busca en prods_odoo el mejor candidato a template padre por nombre similar.
    Opcionalmente pondera con similitud de imagen (phash).
    Devuelve dict {prod, score, sim_nombre} o None.
    """
    if not nombre_base or not prods_odoo:
        return None
    phash_nuevo = _phash_imagen(imagen_data) if imagen_data else None
    mejor       = None
    mejor_score = 0.0
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


def analizar_clasificacion_packing(file_bytes: bytes, productos: list[dict]) -> list[dict]:
    """
    Análisis unificado padre/variante para todos los productos del packing list.
    Reemplaza la detección de duplicados + la detección de conflictos SKU.

    Flujo:
    1. Extrae imágenes del Excel y las guarda en temp
    2. Claude Vision por producto (subcategoria_cod, atributo_cod, descripcion…)
    3. Extrae nombre_base + atributo en batch (un solo call a Claude Haiku)
    4. Busca padres en Odoo por nombre
    5. Agrupa por nombre_base (detecta variantes y duplicados locales)
    6. Genera propuesta de SKU y acción para cada producto

    Devuelve lista de propuestas (una por producto del packing list).
    """
    imagenes     = extraer_imagenes_excel(file_bytes)
    prods_odoo   = st.session_state.get("odoo_productos", [])
    phashes_odoo = st.session_state.get("odoo_phashes", {})
    skus_odoo    = st.session_state.get("odoo_skus", [])

    if skus_odoo:
        sincronizar_contadores_con_odoo(skus_odoo)
    if imagenes:
        guardar_imagenes_temp(imagenes, productos)

    # ── Paso 1: Claude Vision por producto ────────────────────────────────────
    # Fórmula de bloques: floor(rate_limit / tokens_por_llamada × 0.8)
    # 30,000 tokens/min ÷ 2,000 tokens/llamada × 0.8 margen = 12 llamadas/bloque
    _RATE_LIMIT   = 30_000
    _TOKENS_CALL  = 2_000
    _BLOQUE_SIZE  = max(1, int(_RATE_LIMIT / _TOKENS_CALL * 0.8))  # = 12
    _n_con_img    = sum(1 for p in productos if imagenes.get(p.get("fila_excel_0idx", 0)))
    _num_bloques  = max(1, math.ceil(_n_con_img / _BLOQUE_SIZE))
    _llamadas_bloque = 0

    datos_vision: dict[int, dict] = {}
    for i, prod in enumerate(productos):
        row_num = prod.get("fila_excel_0idx", i + 1)
        img     = imagenes.get(row_num)
        if img:
            ctx   = {"nombre": prod.get("nombre"), "material": prod.get("material"),
                     "uso": prod.get("uso")}
            datos = analizar_imagen_claude(img["data"], img["ext"], ctx)
            _llamadas_bloque += 1
            # Pausa entre bloques para no exceder el rate limit
            if _llamadas_bloque >= _BLOQUE_SIZE and i < len(productos) - 1:
                time.sleep(62)
                _llamadas_bloque = 0
        else:
            datos = _inferir_categoria_manual(
                prod.get("nombre", ""), prod.get("material", ""), prod.get("uso", "")
            )
            datos.setdefault("atributo_desc", datos.get("atributo_cod", ""))
        datos_vision[i] = datos

    # Liberar bytes de imágenes — ya están en disco y ya se usaron para Vision
    del imagenes

    # ── Paso 2: Nombre base + atributo en batch ───────────────────────────────
    clasificaciones = _extraer_nombre_base_atributo_batch(productos)

    # ── Paso 3: Construir propuestas ──────────────────────────────────────────
    # Registro local: nombre_base → {sku_padre, subcod, numero, odoo_id, odoo_nombre,
    #                                 padre_idx, variantes: {att_cod: {sku, idx}}}
    registro: dict[str, dict] = {}
    propuestas: list[dict] = []

    ACCION_LABELS = {
        "crear_padre_y_variante": "🆕 Padre nuevo + variante nueva",
        "crear_padre_solo":       "🆕 Padre nuevo (sin variante)",
        "crear_variante":         "➕ Variante nueva en padre existente",
        "reutilizar":             "♻️ Reutilizar producto existente",
        "duplicado":              "🔁 Duplicado — mismo producto ya registrado",
    }

    for i, prod in enumerate(productos):
        datos   = datos_vision[i]
        clas    = clasificaciones[i] if i < len(clasificaciones) else {}
        row_num = prod.get("fila_excel_0idx", i + 1)
        nombre      = prod.get("nombre") or f"Producto {i+1}"
        nombre_base = (clas.get("nombre_base") or nombre).strip()
        att_tipo    = clas.get("atributo_tipo")
        att_valor   = clas.get("atributo_valor")
        att_cod     = _atributo_cod_desde_valor(att_valor) or datos.get("atributo_cod")
        sub_cod     = datos.get("subcategoria_cod") or "VAR"

        prop: dict = {
            "idx":             i,
            "producto":        prod,
            "nombre":          nombre,
            "nombre_base":     nombre_base,
            "atributo_tipo":   att_tipo,
            "atributo_valor":  att_valor,
            "atributo_cod":    att_cod,
            "subcategoria_cod": sub_cod,
            "datos_vision":    datos,
            "imagen_temp_path": prod.get("imagen_temp_path"),
            "padre_fuente":    "nuevo",
            "padre_odoo_id":   None,
            "padre_odoo_nombre": "",
            "padre_sku":       "",
            "sku":             "",
            "accion":          "crear_padre_y_variante",
            "accion_display":  "",
            "confianza_padre": 0.0,
            "requiere_revision": False,
            "nota_revision":   "",
            "duplicado_de_idx": None,
        }

        if nombre_base in registro:
            # ── Padre ya en registro local ─────────────────────────────────
            entrada = registro[nombre_base]
            prop.update({
                "padre_fuente":      "local",
                "padre_sku":         entrada["sku_padre"],
                "padre_odoo_id":     entrada["odoo_id"],
                "padre_odoo_nombre": entrada["odoo_nombre"],
                "confianza_padre":   1.0,
            })
            if not att_cod:
                # Sin atributo → duplicado del padre
                prop["accion"]           = "duplicado"
                prop["duplicado_de_idx"] = entrada.get("padre_idx")
                prop["sku"]              = entrada["sku_padre"]
            elif att_cod in entrada["variantes"]:
                # Misma variante ya existe en este packing list
                prop["accion"]           = "duplicado"
                prop["duplicado_de_idx"] = entrada["variantes"][att_cod]["idx"]
                prop["sku"]              = entrada["variantes"][att_cod]["sku"]
            else:
                # Nueva variante del mismo padre
                prop["accion"] = "crear_variante"
                sku_var = _sku_mismo_numero(entrada["subcod"], entrada["numero"], att_cod)
                if skus_odoo:
                    sku_var = validar_sku_vs_odoo(sku_var, skus_odoo)["sku_ajustado"]
                prop["sku"] = sku_var
                entrada["variantes"][att_cod] = {"sku": sku_var, "idx": i}
        else:
            # ── Buscar en Odoo ─────────────────────────────────────────────
            candidato = None
            if prods_odoo:
                _img_path = prod.get("imagen_temp_path")
                _img_bytes = Path(_img_path).read_bytes() if _img_path and Path(_img_path).exists() else None
                candidato = _buscar_padre_en_odoo_por_nombre(
                    nombre_base, prods_odoo, phashes_odoo, _img_bytes
                )
                del _img_bytes

            if candidato:
                prod_odoo  = candidato["prod"]
                sku_odoo   = prod_odoo.get("default_code", "")
                partes     = sku_odoo.replace("_test", "").split("-")
                try:
                    num_odoo = int(partes[1]) if len(partes) >= 2 else 1
                    sub_odoo = partes[0]
                except (ValueError, IndexError):
                    num_odoo, sub_odoo = 1, sub_cod
                prop.update({
                    "padre_fuente":      "odoo",
                    "padre_odoo_id":     prod_odoo.get("id"),
                    "padre_odoo_nombre": prod_odoo.get("name", ""),
                    "padre_sku":         sku_odoo,
                    "confianza_padre":   candidato["score"],
                })
                if att_cod:
                    sku_var = _sku_mismo_numero(sub_odoo, num_odoo, att_cod)
                    if skus_odoo:
                        sku_var = validar_sku_vs_odoo(sku_var, skus_odoo)["sku_ajustado"]
                    prop["sku"]    = sku_var
                    prop["accion"] = "crear_variante"
                else:
                    prop["sku"]    = sku_odoo
                    prop["accion"] = "reutilizar"
                registro[nombre_base] = {
                    "sku_padre": sku_odoo, "subcod": sub_odoo, "numero": num_odoo,
                    "odoo_id": prod_odoo.get("id"), "odoo_nombre": prod_odoo.get("name", ""),
                    "padre_idx": i,
                    "variantes": {att_cod: {"sku": prop["sku"], "idx": i}} if att_cod else {},
                }
            else:
                # ── Crear padre nuevo ──────────────────────────────────────
                sku_padre = _generar_sku_padre(sub_cod)
                partes    = sku_padre.replace("_test", "").split("-")
                try:
                    num_nuevo = int(partes[1]) if len(partes) >= 2 else 1
                except (ValueError, IndexError):
                    num_nuevo = 1
                prop["padre_sku"] = sku_padre
                if att_cod:
                    sku_var = _sku_mismo_numero(sub_cod, num_nuevo, att_cod)
                    if skus_odoo:
                        sku_var = validar_sku_vs_odoo(sku_var, skus_odoo)["sku_ajustado"]
                    prop["sku"]    = sku_var
                    prop["accion"] = "crear_padre_y_variante"
                else:
                    prop["sku"]    = sku_padre
                    prop["accion"] = "crear_padre_solo"
                registro[nombre_base] = {
                    "sku_padre": sku_padre, "subcod": sub_cod, "numero": num_nuevo,
                    "odoo_id": None, "odoo_nombre": "",
                    "padre_idx": i,
                    "variantes": {att_cod: {"sku": prop["sku"], "idx": i}} if att_cod else {},
                }

        prop["accion_display"] = ACCION_LABELS.get(prop["accion"], prop["accion"])

        # Enriquecer el producto con datos de visión para el Excel
        prod.update({
            "sku":             prop["sku"],
            "subcategoria_cod": sub_cod,
            "atributo_cod":    att_cod or datos.get("atributo_cod", "EST"),
            "atributo_desc":   att_valor or datos.get("atributo_desc", ""),
            "atributo":        att_valor or datos.get("atributo_desc", ""),
            "descripcion":     datos.get("descripcion", ""),
            "categoria":       datos.get("categoria", ""),
            "titulo":          datos.get("titulo") or nombre,
        })
        propuestas.append(prop)

    del datos_vision
    return propuestas


def generar_reporte_clasificacion(propuestas: list[dict], nombre_packing: str = "") -> bytes:
    """
    Genera Excel de reporte de clasificación para revisión de bodega.
    Columnas: Fila PL, Nombre, Nombre Base, SKU, Padre SKU, Atributo, Acción,
              Padre en Odoo, Requiere Revisión, Nota.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Clasificación Productos"

    from openpyxl.styles import PatternFill as _PF, Font as _Ft, Alignment as _Al

    HEADERS = [
        "Fila PL", "Nombre producto", "Nombre base", "SKU propuesto",
        "SKU padre", "Tipo atributo", "Valor atributo",
        "Acción", "Padre en Odoo", "Requiere revisión", "Nota",
    ]
    hdr_fill = _PF("solid", fgColor="1F3864")
    hdr_font = _Ft(color="FFFFFF", bold=True, size=10)
    aln_ctr  = _Al(horizontal="center", vertical="center", wrap_text=True)
    aln_lft  = _Al(horizontal="left",   vertical="center", wrap_text=True)

    for c, h in enumerate(HEADERS, 1):
        cell = ws.cell(1, c, h)
        cell.fill      = hdr_fill
        cell.font      = hdr_font
        cell.alignment = aln_ctr

    FILL_REV  = _PF("solid", fgColor="FFF2CC")   # amarillo — requiere revisión
    FILL_DUP  = _PF("solid", fgColor="E2EFDA")   # verde claro — duplicado
    FILL_ODOO = _PF("solid", fgColor="DAE8FC")   # azul claro — desde Odoo
    FILL_NONE = _PF(fill_type=None)

    for r, prop in enumerate(propuestas, 2):
        prod   = prop["producto"]
        fila   = str((prod.get("fila_excel_0idx") or 0) + 1)
        accion = prop["accion"]
        if prop["requiere_revision"]:
            fill = FILL_REV
        elif accion == "duplicado":
            fill = FILL_DUP
        elif prop["padre_fuente"] == "odoo":
            fill = FILL_ODOO
        else:
            fill = FILL_NONE

        att_str = ""
        if prop.get("atributo_tipo") and prop.get("atributo_valor"):
            att_tipo  = prop["atributo_tipo"]
            att_valor = prop["atributo_valor"]
        else:
            att_tipo  = prop.get("atributo_tipo") or "—"
            att_valor = prop.get("atributo_valor") or "—"

        valores = [
            fila,
            prop["nombre"],
            prop["nombre_base"],
            prop["sku"],
            prop["padre_sku"] or "—",
            att_tipo,
            att_valor,
            prop["accion_display"],
            prop["padre_odoo_nombre"] or "—",
            "⚠️ SÍ" if prop["requiere_revision"] else "",
            prop["nota_revision"] or "",
        ]
        for c, v in enumerate(valores, 1):
            cell = ws.cell(r, c, v)
            cell.fill      = fill
            cell.alignment = aln_lft if c in (2, 3, 11) else aln_ctr

    widths = [8, 30, 25, 16, 14, 14, 16, 26, 25, 16, 30]
    for c, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def crear_clasificacion_en_odoo(propuestas: list[dict],
                                 tipo_cambio: float,
                                 costo_contenedor: float) -> tuple[list[str], list[str]]:
    """
    Crea templates padre y variantes en Odoo según las propuestas confirmadas.
    Llena: description_sale (material, uso, precio USD, dimensiones, piezas/caja),
           description/notas internas (CBM, guía, cajas, flete), volume (inventario).
    Respeta modo_prueba: nombre con sufijo _test y categoría PRUEBAS_AGENTE.
    Devuelve (resultados_ok, errores).
    """
    odoo_url  = os.environ.get("ODOO_URL", "")
    odoo_db   = os.environ.get("ODOO_DB", "")
    odoo_user = os.environ.get("ODOO_USER", "")
    odoo_pass = os.environ.get("ODOO_PASSWORD", "")
    if not all([odoo_url, odoo_db, odoo_user, odoo_pass]):
        return [], ["Faltan credenciales ODOO en .env"]
    try:
        common = xmlrpc.client.ServerProxy(f"{odoo_url}/xmlrpc/2/common", allow_none=True)
        uid    = common.authenticate(odoo_db, odoo_user, odoo_pass, {})
        models = xmlrpc.client.ServerProxy(f"{odoo_url}/xmlrpc/2/object", allow_none=True)
    except Exception as e:
        return [], [f"No se pudo conectar a ODOO: {e}"]

    modo_prueba  = st.session_state.get("modo_prueba", False)
    cbm_total    = sum(_cbm_total_fila(p["producto"])
                       for p in propuestas if p["accion"] != "duplicado")
    costo_por_m3 = costo_contenedor / cbm_total if cbm_total > 0 else 0.0

    # Categoría
    cat_cache: dict[str, int] = {}
    def _get_cat(nombre_cat: str) -> int | None:
        if nombre_cat in cat_cache:
            return cat_cache[nombre_cat]
        try:
            ids = models.execute_kw(odoo_db, uid, odoo_pass, "product.category", "search",
                                    [[["name", "=", nombre_cat]]])
            cid = ids[0] if ids else models.execute_kw(
                odoo_db, uid, odoo_pass, "product.category", "create", [{"name": nombre_cat}])
            cat_cache[nombre_cat] = cid
            return cid
        except Exception:
            return None

    templates_creados: dict[str, int] = {}  # sku_padre → tmpl_id
    attrs_cache:       dict[str, int] = {}
    attr_vals_cache:   dict[tuple, int] = {}
    tags_cache:        dict[str, int] = {}

    def _get_or_create_tag(nombre: str) -> int | None:
        if nombre in tags_cache:
            return tags_cache[nombre]
        try:
            ids = models.execute_kw(odoo_db, uid, odoo_pass, "product.tag", "search",
                                    [[["name", "=", nombre]]])
            tid = ids[0] if ids else models.execute_kw(
                odoo_db, uid, odoo_pass, "product.tag", "create", [{"name": nombre}])
            tags_cache[nombre] = tid
            return tid
        except Exception:
            return None

    def _get_or_create_attr(nombre: str) -> int:
        if nombre in attrs_cache:
            return attrs_cache[nombre]
        ids = models.execute_kw(odoo_db, uid, odoo_pass, "product.attribute", "search",
                                [[["name", "=", nombre]]])
        aid = ids[0] if ids else models.execute_kw(
            odoo_db, uid, odoo_pass, "product.attribute", "create", [{"name": nombre}])
        attrs_cache[nombre] = aid
        return aid

    def _get_or_create_val(attr_id: int, valor: str) -> int:
        key = (attr_id, valor)
        if key in attr_vals_cache:
            return attr_vals_cache[key]
        ids = models.execute_kw(odoo_db, uid, odoo_pass, "product.attribute.value", "search",
                                [[["attribute_id", "=", attr_id], ["name", "=", valor]]])
        vid = ids[0] if ids else models.execute_kw(
            odoo_db, uid, odoo_pass, "product.attribute.value", "create",
            [{"name": valor, "attribute_id": attr_id}])
        attr_vals_cache[key] = vid
        return vid

    def _build_vals(prod: dict, nombre_tmpl: str, precio_usd: float,
                    precio_mxn: float, costo_unitario: float,
                    sku_direct: str | None = None) -> dict:
        """Construye el dict de valores para product.template con todos los campos."""
        cbm_pz       = _cbm_por_pieza(prod)
        costo_cbm_pz = round(cbm_pz * costo_por_m3, 2)
        largo  = prod.get("largo_cm")
        ancho  = prod.get("ancho_cm")
        alto   = prod.get("alto_cm")
        dims_str = (" × ".join(str(v) for v in [largo, ancho, alto] if v)
                    if any([largo, ancho, alto]) else None)

        # ── Descripción de venta (pestaña Ventas / visible al cliente) ─────────
        desc_partes = []
        if prod.get("descripcion"):
            desc_partes.append(prod["descripcion"])
        if prod.get("material"):
            desc_partes.append(f"Material: {prod['material']}")
        if prod.get("uso"):
            desc_partes.append(f"Uso: {prod['uso']}")
        if precio_usd > 0:
            desc_partes.append(f"Precio USD: ${precio_usd:.2f}")
        pzs = prod.get("piezas_x_caja")
        if pzs:
            try:
                desc_partes.append(f"Piezas por caja: {int(float(pzs))}")
            except (ValueError, TypeError):
                desc_partes.append(f"Piezas por caja: {pzs}")
        if dims_str:
            desc_partes.append(f"Dimensiones: {dims_str} cm")

        # ── Notas internas (pestaña Notas / internal notes) ────────────────────
        notas = []
        if prod.get("nombre_alt"):
            notas.append(f"Nombre alternativo: {prod['nombre_alt']}")
        if prod.get("id_guia"):
            notas.append(f"ID Guía / Referencia: {prod['id_guia']}")
        if prod.get("cajas_master"):
            try:
                notas.append(f"Cajas master: {int(float(prod['cajas_master']))}")
            except (ValueError, TypeError):
                notas.append(f"Cajas master: {prod['cajas_master']}")
        if prod.get("piezas_total"):
            try:
                notas.append(f"Piezas totales en contenedor: {int(float(prod['piezas_total']))}")
            except (ValueError, TypeError):
                notas.append(f"Piezas totales en contenedor: {prod['piezas_total']}")
        if cbm_pz > 0:
            notas.append(f"CBM por pieza: {cbm_pz:.6f} m³")
        if prod.get("cbm_master_carton"):
            notas.append(f"CBM master carton: {prod['cbm_master_carton']}")
        if prod.get("cbm_total_sku"):
            notas.append(f"CBM total SKU: {prod['cbm_total_sku']}")
        if prod.get("cbm_inner_carton"):
            notas.append(f"CBM inner carton: {prod['cbm_inner_carton']}")
        if dims_str:
            notas.append(f"Dimensiones caja: {dims_str} cm")
        if costo_cbm_pz > 0:
            notas.append(f"Costo flete CBM/pieza: ${costo_cbm_pz:.4f} MXN")

        vals: dict = {
            "name":             nombre_tmpl,
            "type":             "product",
            "sale_ok":          True,
            "purchase_ok":      True,
            "list_price":       precio_mxn,
            "standard_price":   costo_unitario,
        }
        if desc_partes:
            vals["description_sale"] = "\n".join(desc_partes)
        if notas:
            vals["description"] = "\n".join(notas)

        # ── Pestaña Inventario / Empaque ──────────────────────────────────────
        # volume: preferir volumen_por_pieza si viene explícito, si no usar cbm_pz
        vol_pz = _safe_float(prod.get("volumen_por_pieza")) or cbm_pz
        if vol_pz > 0:
            vals["volume"]          = vol_pz
            vals["cbm_per_product"] = vol_pz
        peso = _safe_float(prod.get("peso_kg"))
        if peso > 0:
            vals["weight"] = peso
        if precio_usd > 0:
            vals["costo_usd"] = precio_usd
        cbm_mc = _safe_float(prod.get("cbm_master_carton"))
        if cbm_mc > 0:
            vals["cbm_master_box"] = cbm_mc        # CBM por caja master
        pzs_caja = prod.get("piezas_x_caja")
        if pzs_caja:
            try:
                vals["units_per_master_box"] = int(float(pzs_caja))
            except (ValueError, TypeError):
                pass
        if largo:
            try:
                vals["length"] = float(largo)
            except (ValueError, TypeError):
                pass
        if ancho:
            try:
                vals["width"] = float(ancho)
            except (ValueError, TypeError):
                pass
        if alto:
            try:
                vals["height"] = float(alto)
            except (ValueError, TypeError):
                pass
        # Número de contenedor (campo texto)
        _contenedor_num = st.session_state.get("contenedor_val", "")
        if _contenedor_num:
            vals["container_numbers"] = _contenedor_num

        if sku_direct:
            vals["default_code"] = sku_direct
        # Categoría: prueba → PRUEBAS_AGENTE, producción → Productos Agente
        cat_nombre = "PRUEBAS_AGENTE" if modo_prueba else "Productos Agente"
        cid = _get_cat(cat_nombre)
        if cid:
            vals["categ_id"] = cid
        return vals

    def _apply_rev_tag_and_note(tmpl_id: int, nota: str) -> None:
        tag_id = _get_or_create_tag("Requiere Revisión")
        update: dict = {}
        if tag_id:
            update["tag_ids"] = [(4, tag_id)]
        if nota:
            existing = models.execute_kw(
                odoo_db, uid, odoo_pass, "product.template", "read",
                [[tmpl_id]], {"fields": ["description"]})
            prev = (existing[0].get("description") or "") if existing else ""
            update["description"] = f"⚠️ REQUIERE REVISIÓN: {nota}\n{prev}".strip()
        if update:
            models.execute_kw(odoo_db, uid, odoo_pass,
                              "product.template", "write", [[tmpl_id], update])

    resultados: list[str] = []
    errores:    list[str] = []

    for prop in propuestas:
        if prop["accion"] == "duplicado":
            continue
        prod           = prop["producto"]
        accion         = prop["accion"]
        requiere_rev   = bool(prop.get("requiere_revision", False))
        nota_rev       = str(prop.get("nota_revision", "") or "")
        sku            = prop["sku"]
        sku_padre      = prop["padre_sku"]
        att_tipo       = prop["atributo_tipo"] or "Variante"
        att_valor      = prop["atributo_valor"] or "Estándar"
        nombre_base    = prop["nombre_base"]
        nombre         = prop["nombre"]
        precio_usd     = _safe_float(prod.get("precio_usd"))
        precio_mxn     = round(precio_usd * tipo_cambio, 2)
        cbm_pz         = _cbm_por_pieza(prod)
        costo_unitario = round(precio_mxn + cbm_pz * costo_por_m3, 2)

        # Nombre del template: sufijo _test en modo prueba
        nombre_tmpl = nombre_base + ("_test" if modo_prueba else "")

        try:
            if accion in ("crear_padre_y_variante", "crear_padre_solo"):
                # Para padre+variante: poner sku_padre en default_code del template
                # Para padre solo: poner el sku directamente
                sku_direct = sku_padre if accion == "crear_padre_y_variante" else sku
                vals = _build_vals(prod, nombre_tmpl, precio_usd, precio_mxn,
                                   costo_unitario, sku_direct=sku_direct)
                tmpl_id = models.execute_kw(
                    odoo_db, uid, odoo_pass, "product.template", "create", [vals])
                templates_creados[sku_padre] = tmpl_id

                if requiere_rev:
                    _apply_rev_tag_and_note(tmpl_id, nota_rev)

                if accion == "crear_padre_solo":
                    # Forzar standard_price en product.product
                    pp_ids = models.execute_kw(odoo_db, uid, odoo_pass, "product.product",
                                               "search", [[["product_tmpl_id", "=", tmpl_id]]])
                    if pp_ids:
                        models.execute_kw(odoo_db, uid, odoo_pass, "product.product", "write",
                                          [pp_ids, {"standard_price": costo_unitario}])
                    _rev_sfx = " ⚠️ rev" if requiere_rev else ""
                    resultados.append(f"✅ Padre creado: `{sku}` — {nombre_tmpl}{_rev_sfx}")
                else:
                    attr_id = _get_or_create_attr(att_tipo)
                    val_id  = _get_or_create_val(attr_id, att_valor)
                    models.execute_kw(odoo_db, uid, odoo_pass,
                                      "product.template.attribute.line", "create",
                                      [{"product_tmpl_id": tmpl_id, "attribute_id": attr_id,
                                        "value_ids": [(4, val_id)]}])
                    # Re-write volume/weight/costo_usd: Odoo puede resetearlos al crear variantes
                    _rewrite: dict = {}
                    _vol_pz = _safe_float(prod.get("volumen_por_pieza")) or cbm_pz
                    if _vol_pz > 0:
                        _rewrite["volume"] = _vol_pz
                    _peso = _safe_float(prod.get("peso_kg"))
                    if _peso > 0:
                        _rewrite["weight"] = _peso
                    if precio_usd > 0:
                        _rewrite["costo_usd"] = precio_usd
                    if _rewrite:
                        models.execute_kw(odoo_db, uid, odoo_pass,
                                          "product.template", "write", [[tmpl_id], _rewrite])
                    var_ids = models.execute_kw(odoo_db, uid, odoo_pass,
                                                "product.product", "search",
                                                [[["product_tmpl_id", "=", tmpl_id]]])
                    if var_ids:
                        models.execute_kw(odoo_db, uid, odoo_pass, "product.product", "write",
                                          [var_ids, {"default_code": sku,
                                                     "standard_price": costo_unitario}])
                    _rev_sfx = " ⚠️ rev" if requiere_rev else ""
                    resultados.append(f"✅ Padre + variante: `{sku_padre}` / `{sku}` — {nombre_tmpl}{_rev_sfx}")

            elif accion == "crear_variante":
                tmpl_id = (templates_creados.get(sku_padre)
                           or prop.get("padre_odoo_id"))
                if not tmpl_id:
                    errores.append(f"⚠️ No se encontró template padre para `{sku}` ({nombre})")
                    continue
                attr_id = _get_or_create_attr(att_tipo)
                val_id  = _get_or_create_val(attr_id, att_valor)
                lineas  = models.execute_kw(
                    odoo_db, uid, odoo_pass, "product.template.attribute.line", "search_read",
                    [[["product_tmpl_id", "=", tmpl_id], ["attribute_id", "=", attr_id]]],
                    {"fields": ["id"]})
                if lineas:
                    models.execute_kw(odoo_db, uid, odoo_pass,
                                      "product.template.attribute.line", "write",
                                      [[lineas[0]["id"]], {"value_ids": [(4, val_id)]}])
                else:
                    models.execute_kw(odoo_db, uid, odoo_pass,
                                      "product.template.attribute.line", "create",
                                      [{"product_tmpl_id": tmpl_id, "attribute_id": attr_id,
                                        "value_ids": [(4, val_id)]}])
                var_ids = models.execute_kw(
                    odoo_db, uid, odoo_pass, "product.product", "search",
                    [[["product_tmpl_id", "=", tmpl_id],
                      ["product_template_attribute_value_ids"
                       ".product_attribute_value_id", "=", val_id]]])
                if var_ids:
                    models.execute_kw(odoo_db, uid, odoo_pass, "product.product", "write",
                                      [var_ids, {"default_code": sku,
                                                 "standard_price": costo_unitario}])
                if requiere_rev:
                    _apply_rev_tag_and_note(tmpl_id, nota_rev)
                _rev_sfx = " ⚠️ rev" if requiere_rev else ""
                resultados.append(f"✅ Variante creada: `{sku}` — {nombre}{_rev_sfx}")

            elif accion == "reutilizar":
                resultados.append(f"♻️ Reutilizado: `{sku}` — {nombre}")

        except Exception as e:
            errores.append(f"❌ Error en `{sku}` ({nombre}): {e}")

    return resultados, errores


# ── Herramienta LangChain ──────────────────────────────────────────────────────
@tool
def generar_excel_tool(productos: List[dict] | None = None) -> str:
    """Genera o regenera el archivo Excel de FERRAFORME.
    Llama esta herramienta cuando el usuario confirme proceder o pida generar el Excel.
    No necesitas pasar productos — se usan los que ya están cargados en sesión."""
    tipo_cambio      = st.session_state.get("tipo_cambio", 19.0)
    contenedor       = st.session_state.get("contenedor_val", "")
    costo_contenedor = st.session_state.get("costo_contenedor", 525000.0)
    nombre_packing   = st.session_state.get("filename", "packing_list.xlsx")

    prods = productos if productos else st.session_state.get("productos", [])
    if not prods:
        return "Error: no hay productos cargados en sesión."

    # costo_por_m3 se calcula una vez y lo usan tanto el master Excel como ODOO
    cbm_total    = sum(_cbm_total_fila(p) for p in prods)
    costo_por_m3 = costo_contenedor / cbm_total if cbm_total > 0 else 0.0

    # Generar Excels y subir a ODOO en paralelo
    with concurrent.futures.ThreadPoolExecutor(max_workers=3) as executor:
        fut_excel  = executor.submit(generar_excel,        prods, tipo_cambio, contenedor)
        fut_master = executor.submit(generar_excel_master, prods, tipo_cambio, costo_contenedor, nombre_packing)
        fut_odoo   = executor.submit(_subir_productos_a_odoo, prods, tipo_cambio, costo_por_m3)

        excel_bytes        = fut_excel.result()
        master_bytes       = fut_master.result()
        subidas, errs_odoo = fut_odoo.result()

    # Limpiar imágenes temporales restantes (fallidas o sin ODOO)
    # Drive ya las tiene; la carpeta temp ya no es necesaria
    if IMAGENES_TEMP_PATH.exists():
        for _f in IMAGENES_TEMP_PATH.iterdir():
            try:
                _f.unlink()
            except Exception:
                pass

    st.session_state.excel_bytes  = excel_bytes
    st.session_state.master_bytes = master_bytes
    st.session_state.productos    = prods

    _reportar_subida_imagenes(subidas, errs_odoo)

    resumen_odoo = (
        f" {len(subidas)} productos subidos a ODOO." if subidas
        else (f" Advertencia ODOO: {errs_odoo[0]}" if errs_odoo else "")
    )
    return f"Excel FERRAFORME y Excel master generados con {len(prods)} productos.{resumen_odoo}"


def llamar_agente(user_input: str, system: str, lc_messages: list) -> str:
    """Llama al agente con LangChain bind_tools. Devuelve el texto de respuesta."""
    llm = ChatAnthropic(model="claude-sonnet-4-6", max_tokens=8000)
    llm_con_tools = llm.bind_tools([generar_excel_tool])

    mensajes = [SystemMessage(content=system)] + lc_messages + [HumanMessage(content=user_input)]

    while True:
        respuesta = llm_con_tools.invoke(mensajes)
        mensajes.append(respuesta)

        if not respuesta.tool_calls:
            break

        for tc in respuesta.tool_calls:
            if tc["name"] == "generar_excel_tool":
                resultado = generar_excel_tool.invoke(tc["args"])
            else:
                resultado = "Herramienta no reconocida."
            mensajes.append(ToolMessage(content=resultado, tool_call_id=tc["id"]))

    # Normalizar: Claude puede devolver lista de bloques en lugar de string
    content = respuesta.content
    if isinstance(content, list):
        partes = []
        for bloque in content:
            if isinstance(bloque, dict) and bloque.get("type") == "text":
                partes.append(bloque.get("text", ""))
            elif isinstance(bloque, str):
                partes.append(bloque)
        return "\n".join(partes)
    return content


# ══════════════════════════════════════════════════════════════════════════════
# STREAMLIT UI
# ══════════════════════════════════════════════════════════════════════════════

st.set_page_config(page_title="FERRAFORME — Agente v2", page_icon="📦", layout="wide")

defaults = {
    "chat":              [],
    "lc_messages":       [],   # historial LangChain (HumanMessage / AIMessage)
    "chat_dudas":        [],   # mensajes de aclaración durante la fase de dudas
    "chat_fase":         [],   # mensajes de aclaración durante duplicados/conflictos
    "analisis":          None,
    "productos":         [],
    "excel_bytes":       None,
    "master_bytes":      None,
    "purchase_bytes":    None,
    "archivo_id":        None,
    "dudas_relevantes":  [],
    "dudas_menores":     [],
    "respuestas_dudas":  {},
    "esperando_dudas":   False,
    "_quick_reply":      None,   # respuesta rápida desde botón
    "sku_contadores":          {},
    "odoo_skus":               [],
    "odoo_productos":          [],   # todos los productos ODOO con name (sin image_128)
    "odoo_phashes":            {},   # {sku: phash}
    "odoo_imagen_urls":        {},   # {sku: url_publica_supabase}
    "odoo_conectado":          False,
    "esperando_duplicados":    False,
    "dup_paso":                1,
    "dup_respuestas":          {},
    # ── Clasificación padre/variante ──────────────────────────────────────────
    "clasificacion_activa":    False,
    "clasificacion_propuestas": [],
    "clasificacion_reporte_bytes": None,
    "duplicados_pendientes":   [],
    "dup_grupos_segunda":      [],      # grupos encontrados en segunda vuelta de detección
    "imagenes_excel":          {},      # {row_0idx: {data, ext}} extraídas antes de duplicados
    "historial_duplicados":    [],      # grupos resueltos para consulta posterior
    "dup_snapshot":            None,   # snapshot visual del último análisis de duplicados
    "esperando_conflictos":    False,
    "conflictos_pendientes":   [],
    "resoluciones_conflictos": {},
    # ── Agregar productos en lote ──────────────────────────────────────────────
    "agregar_lote_activo": False,
    "agregar_lote_paso":   1,       # 1=subir imgs, 2=procesar, 3=revisar/confirmar
    "agregar_lote_n":      1,
    "agregar_lote_imgs":   {},      # {idx: {"data": bytes, "ext": str, "nombre": str}}
    "agregar_lote_prods":  [],      # productos procesados (con SKU)
    "agregar_lote_conf":   [],      # conflictos ODOO para estos productos
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ── Auto-cargar cache ODOO al inicio ───────────────────────────────────────────
if not st.session_state.odoo_conectado and not st.session_state.odoo_skus:
    _cache_auto = cargar_cache_odoo()
    if _cache_auto:
        st.session_state.odoo_skus      = _cache_auto["skus"]
        st.session_state.odoo_productos = _cache_auto["productos"]
        st.session_state.odoo_phashes   = _cache_auto["phashes"]
        st.session_state.odoo_conectado = True
        sincronizar_contadores_con_odoo(_cache_auto["skus"])

# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Configuración")
    tipo_cambio = st.number_input("Tipo de cambio USD → MXN", value=19.00, step=0.50, format="%.2f", key="tipo_cambio")
    costo_contenedor = st.number_input("Costo contenedor (MXN)", value=525000.0, step=1000.0, format="%.0f", key="costo_contenedor")

    if "contenedor_val" not in st.session_state:
        st.session_state.contenedor_val = ""
    contenedor_input = st.text_input("Número de contenedor", value=st.session_state.contenedor_val)
    st.session_state.contenedor_val = contenedor_input
    contenedor = contenedor_input

    st.divider()
    modo_prueba = st.toggle(
        "🧪 Modo prueba",
        value=st.session_state.get("modo_prueba", False),
        help="Los SKUs generados llevarán el sufijo _test para no mezclar con productos reales",
    )
    st.session_state.modo_prueba = modo_prueba
    if modo_prueba:
        st.caption("⚠️ SKUs con sufijo `_test` — solo para pruebas")

    st.divider()
    st.subheader("🔗 ODOO")

    # Mostrar info del cache si existe
    info_cache = info_cache_odoo()
    if info_cache:
        st.caption(f"💾 Cache: {info_cache}")

    col_odoo1, col_odoo2 = st.columns(2)
    with col_odoo1:
        cargar_btn = st.button("🔄 Cargar SKUs", width="stretch",
                               help="Usa Supabase si tiene datos; si no, carga desde ODOO")
    with col_odoo2:
        forzar_btn = st.button("🔃 Actualizar", width="stretch",
                               help="Fuerza recarga desde ODOO y sincroniza a Supabase")

    def _aplicar_datos_odoo(skus, productos, phashes, imagen_urls=None, desde_cache=False):
        st.session_state.odoo_skus       = skus
        st.session_state.odoo_productos  = productos
        st.session_state.odoo_phashes    = phashes
        st.session_state.odoo_imagen_urls = imagen_urls or {}
        st.session_state.odoo_conectado  = True
        sin_img = len(skus) - len(phashes)
        origen  = "Supabase" if desde_cache else "ODOO"
        st.success(f"✅ {len(skus)} SKUs · {len(phashes)} imágenes · {sin_img} sin imagen  —  {origen}")

    def _cargar_desde_odoo(forzar=False):
        odoo_url  = os.environ.get("ODOO_URL", "")
        odoo_db   = os.environ.get("ODOO_DB", "")
        odoo_user = os.environ.get("ODOO_USER", "")
        odoo_pass = os.environ.get("ODOO_PASSWORD", "")
        if not all([odoo_url, odoo_db, odoo_user, odoo_pass]):
            st.warning("Faltan credenciales ODOO en el .env")
            return

        # 1. Intentar Supabase primero (salvo recarga forzada)
        if not forzar:
            # 1a. pkl local primero (instantáneo)
            cache = cargar_cache_odoo()
            if cache:
                _aplicar_datos_odoo(
                    cache["skus"], cache["productos"], cache["phashes"],
                    desde_cache=True,
                )
                return
            # 1b. Supabase como fallback (cuando no hay pkl — ej. Digital Ocean)
            with st.spinner("Cargando desde Supabase..."):
                sb_data = _cargar_desde_supabase()
            if sb_data:
                _aplicar_datos_odoo(
                    sb_data["skus"], sb_data["productos"], sb_data["phashes"],
                    imagen_urls=sb_data["imagen_urls"], desde_cache=True,
                )
                return

        # 2. Cargar desde ODOO
        with st.spinner("Conectando a ODOO..."):
            skus, error = cargar_skus_odoo(odoo_url, odoo_db, odoo_user, odoo_pass)
        if error:
            st.error(f"Error: {error}")
            return
        with st.spinner("Descargando productos e imágenes desde ODOO..."):
            prods_odoo = cargar_todos_productos_odoo(odoo_url, odoo_db, odoo_user, odoo_pass)

        # 3. Sincronizar a Supabase (calcula phashes, sube imágenes, guarda filas)
        with st.spinner("Sincronizando con Supabase..."):
            sb_result = _sincronizar_a_supabase(prods_odoo)
        phashes     = sb_result["phashes"]
        imagen_urls = sb_result["imagen_urls"]

        # Si Supabase no está configurado, calcular phashes localmente y guardar pkl
        if not phashes:
            with st.spinner("Calculando hashes de imágenes..."):
                for p in prods_odoo:
                    if p.get("image_128"):
                        try:
                            h = _phash_imagen(base64.b64decode(p["image_128"]))
                            if h is not None:
                                phashes[p["default_code"]] = h
                        except Exception:
                            pass
            guardar_cache_odoo(skus, prods_odoo, phashes)

        # Quitar image_128 de memoria antes de guardar en session_state
        prods_sin_img = [{k: v for k, v in p.items() if k != "image_128"} for p in prods_odoo]
        _aplicar_datos_odoo(skus, prods_sin_img, phashes, imagen_urls=imagen_urls, desde_cache=False)

    if cargar_btn:
        _cargar_desde_odoo(forzar=False)
    if forzar_btn:
        _cargar_desde_odoo(forzar=True)

    if st.session_state.odoo_conectado:
        st.success(f"✅ ODOO activo — {len(st.session_state.odoo_skus)} SKUs en sesión")
    else:
        st.warning("⚠️ ODOO no conectado — presiona **Cargar SKUs**")

    # ── Limpiar productos _test ────────────────────────────────────────────────
    if st.button("🗑️ Eliminar productos _test", width="stretch",
                 help="Elimina de ODOO todos los productos de la categoría PRUEBAS_AGENTE"):
        odoo_url  = os.environ.get("ODOO_URL", "")
        odoo_db   = os.environ.get("ODOO_DB", "")
        odoo_user = os.environ.get("ODOO_USER", "")
        odoo_pass = os.environ.get("ODOO_PASSWORD", "")
        if not all([odoo_url, odoo_db, odoo_user, odoo_pass]):
            st.warning("Faltan credenciales ODOO en el .env")
        else:
            try:
                common = xmlrpc.client.ServerProxy(f"{odoo_url}/xmlrpc/2/common", allow_none=True)
                uid    = common.authenticate(odoo_db, odoo_user, odoo_pass, {})
                models = xmlrpc.client.ServerProxy(f"{odoo_url}/xmlrpc/2/object", allow_none=True)

                # Buscar la categoría PRUEBAS_AGENTE
                cat_ids = models.execute_kw(odoo_db, uid, odoo_pass,
                    "product.category", "search",
                    [[["name", "=", "PRUEBAS_AGENTE"]]])

                if not cat_ids:
                    st.warning("No se encontró la categoría PRUEBAS_AGENTE en ODOO.")
                else:
                    ids_a_borrar = models.execute_kw(odoo_db, uid, odoo_pass,
                        "product.template", "search",
                        [[["categ_id", "in", cat_ids]]])

                    if not ids_a_borrar:
                        st.info("No hay productos en PRUEBAS_AGENTE para eliminar.")
                    else:
                        # Intentar borrar; si hay restricciones de FK, archivar en su lugar
                        borrados  = 0
                        archivados = 0
                        errores_ind = []
                        for pid in ids_a_borrar:
                            try:
                                models.execute_kw(odoo_db, uid, odoo_pass,
                                    "product.template", "unlink", [[pid]])
                                borrados += 1
                            except Exception:
                                try:
                                    models.execute_kw(odoo_db, uid, odoo_pass,
                                        "product.template", "write",
                                        [[pid], {"active": False}])
                                    archivados += 1
                                except Exception as e2:
                                    errores_ind.append(str(e2))
                        partes = []
                        if borrados:
                            partes.append(f"{borrados} eliminado(s)")
                        if archivados:
                            partes.append(f"{archivados} archivado(s)")
                        if partes:
                            st.success(f"✅ {', '.join(partes)} de PRUEBAS_AGENTE.")
                        if errores_ind:
                            st.warning(f"⚠️ {len(errores_ind)} producto(s) no pudieron procesarse: {errores_ind[0]}")
            except Exception as e:
                st.error(f"Error al eliminar: {e}")

    # ── Generar Master Costos rápido (sin pasar por todo el flujo) ───────────
    st.divider()
    st.caption("🧪 Prueba rápida de costos")
    _prods_test = st.session_state.get("productos", [])
    if _prods_test:
        if st.button("📊 Generar Master Costos", width="stretch",
                     help="Genera el Excel de costos con los productos ya cargados"):
            try:
                _tc   = st.session_state.get("tipo_cambio", 19.0)
                _cc   = st.session_state.get("costo_contenedor", 525000.0)
                _npk  = st.session_state.get("filename", "packing_list.xlsx")
                _mbytes = generar_excel_master(_prods_test, _tc, _cc, _npk)
                st.session_state.master_bytes = _mbytes
                st.rerun()
            except Exception as _e:
                st.error(f"Error: {_e}")
    else:
        st.caption("_(carga un packing list primero)_")

    st.divider()
    archivo = st.file_uploader("📂 Packing list (.xlsx)", type=["xlsx"])

    st.divider()
    if st.session_state.excel_bytes:
        filename   = st.session_state.get("filename", "packing_list.xlsx")
        nombre_out = filename.replace(".xlsx", "") + "_FERRAFORME.xlsx"
        st.success("✅ Excels listos para descargar")
        st.download_button(
            label="⬇️ Descargar Excel FERRAFORME",
            data=st.session_state.excel_bytes,
            file_name=nombre_out,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )
    if st.session_state.master_bytes:
        filename     = st.session_state.get("filename", "packing_list.xlsx")
        nombre_master = filename.replace(".xlsx", "") + "_MASTER.xlsx"
        st.download_button(
            label="⬇️ Descargar Excel Master Costos",
            data=st.session_state.master_bytes,
            file_name=nombre_master,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )
        if st.button("🔄 Nuevo packing list", width="stretch"):
            # Limpiar carpeta de imágenes temporales
            if IMAGENES_TEMP_PATH.exists():
                for _f in IMAGENES_TEMP_PATH.iterdir():
                    try:
                        _f.unlink()
                    except Exception:
                        pass
            _chat_hist   = st.session_state.chat[:]
            _lc_hist     = st.session_state.lc_messages[:]
            _odoo_s      = st.session_state.odoo_skus
            _odoo_p      = st.session_state.odoo_productos
            _odoo_ph     = st.session_state.odoo_phashes
            _odoo_c      = st.session_state.odoo_conectado
            for k in defaults:
                st.session_state[k] = defaults[k]
            st.session_state.chat           = _chat_hist
            st.session_state.lc_messages    = _lc_hist
            st.session_state.odoo_skus      = _odoo_s
            st.session_state.odoo_productos = _odoo_p
            st.session_state.odoo_phashes   = _odoo_ph
            st.session_state.odoo_conectado = _odoo_c
            st.session_state.pop("filename", None)
            st.rerun()

    # ── Historial de duplicados/variantes ─────────────────────────────────────
    _hist_dup = st.session_state.get("historial_duplicados", [])
    if _hist_dup:
        st.divider()
        with st.expander(f"📋 Historial de similares/variantes ({sum(len(e['grupos']) for e in _hist_dup)} grupos)"):
            _TIPO_LABEL = {
                "exacto":         "🔁 Exacto",
                "probable":       "❓ Probable",
                "similar":        "📦 Datos iguales",
                "nombre_similar": "🔤 Nombre parecido",
            }
            for entrada in reversed(_hist_dup):
                st.caption(f"**{entrada['archivo']}** — {entrada['fecha']}")
                for g in entrada["grupos"]:
                    tipo_lbl = _TIPO_LABEL.get(g["tipo"], g["tipo"])
                    dec = entrada["respuestas"].get(str(g["id"]))
                    if dec == "diferente":
                        dec_lbl = "🔗 Independientes"
                    elif isinstance(dec, dict):
                        tipo_d  = dec.get("tipo", "variantes")
                        n_sel   = len(dec.get("sel", []))
                        n_indep = len(g["indices"]) - n_sel
                        if tipo_d == "mismo":
                            dec_lbl = f"🔀 Fusionados ({n_sel})" + (f" · 🔗 {n_indep} indep." if n_indep else "")
                        else:
                            dec_lbl = f"📂 Variantes ({n_sel})" + (f" · 🔗 {n_indep} indep." if n_indep else "")
                    else:
                        dec_lbl = "📂 Variantes"
                    nombres = [p.get("nombre", "—") for p in g["productos"]]
                    filas   = [str(p.get("fila_excel_0idx", g["indices"][ci]) + 1) for ci, p in enumerate(g["productos"])]
                    st.markdown(
                        f"- {tipo_lbl} · {dec_lbl} · Filas {', '.join(filas)}:  \n"
                        + "  ".join(f"`{n}`" for n in nombres)
                    )
                st.divider()

def _inferir_categoria_manual(nombre: str, material: str, uso: str) -> dict:
    """Infiere subcategoria_cod y atributo_cod de un producto sin imagen, usando Claude Haiku."""
    import json
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
        return {"subcategoria_cod": "VAR", "atributo_cod": "GEN"}


def _agregar_producto_manual(
    nombre: str,
    precio_usd, piezas_x_caja, piezas_total, cajas_master,
    cbm_por_pieza, cbm_master_carton, cbm_total_sku,
    material: str, uso: str,
    imagen_bytes: bytes | None,
    imagen_ext: str | None,
) -> str:
    """Crea el producto manualmente, genera su SKU y lo agrega a la sesión. Devuelve el SKU."""
    n_manuales = sum(1 for p in st.session_state.get("productos", []) if p.get("_manual"))
    fila_0idx  = -(n_manuales + 1)   # índice negativo para no colisionar con filas del Excel

    prod: dict = {
        "nombre":            nombre,
        "precio_usd":        precio_usd        or None,
        "piezas_x_caja":     piezas_x_caja     or None,
        "piezas_total":      piezas_total       or None,
        "cajas_master":      cajas_master       or None,
        "cbm_por_pieza":     cbm_por_pieza      or None,
        "cbm_master_carton": cbm_master_carton  or None,
        "cbm_total_sku":     cbm_total_sku      or None,
        "material":          material           or "",
        "uso":               uso                or "",
        "fila_excel_0idx":   fila_0idx,
        "_manual":           True,
    }

    # Obtener metadata para el SKU
    if imagen_bytes and imagen_ext:
        datos = analizar_imagen_claude(imagen_bytes, imagen_ext, contexto=prod)
    else:
        datos = _inferir_categoria_manual(nombre, material, uso)

    sub_cod     = datos.get("subcategoria_cod", "VAR")
    att_cod     = datos.get("atributo_cod",     "GEN")
    sku_inicial = generar_sku(sub_cod, att_cod)
    skus_odoo   = st.session_state.get("odoo_skus", [])
    if skus_odoo:
        prod["sku"] = validar_sku_vs_odoo(sku_inicial, skus_odoo)["sku_ajustado"]
    else:
        prod["sku"] = sku_inicial

    if st.session_state.get("modo_prueba"):
        prod["sku"] += "_test"

    # Guardar imagen en carpeta temporal si existe
    if imagen_bytes and imagen_ext:
        try:
            stem = f"manual_{abs(fila_0idx)}"
            IMAGENES_TEMP_PATH.mkdir(parents=True, exist_ok=True)
            (IMAGENES_TEMP_PATH / f"{stem}.{imagen_ext}").write_bytes(imagen_bytes)
            prod["imagen_temp_stem"] = stem
        except Exception:
            pass

    st.session_state.productos.append(prod)
    return prod["sku"]


def _iniciar_chat(analisis, productos, advertencias, tipo_cambio, contenedor, n_imgs: int = 0):
    """Lanza el primer turno del chat después de resolver las dudas."""
    # ── Calcular CBM total y añadirlo a la tabla de resumen ya mostrada ────────
    total_cbm = sum(_cbm_total_fila(p) for p in productos)
    cbm_status = "⚠️ **Supera los 70 CBM** — verifica la capacidad del contenedor." if total_cbm > 70 else "✅ Dentro del límite de 70 CBM."
    cbm_line = f"\n\n---\n📦 **CBM Total del cargamento: {total_cbm:.4f}** — {cbm_status}"
    if n_imgs > 0:
        n_con_sku      = sum(1 for p in productos if p.get("sku"))
        sin_imagen     = [p for p in productos if p.get("sin_imagen")]
        errores_gemini = st.session_state.pop("gemini_errores", [])
        cbm_line      += f"\n🏷️ **SKUs generados: {n_con_sku} / {len(productos)} productos**"
        if errores_gemini:
            cbm_line += f"\n⚠️ **{len(errores_gemini)} imágenes no pudieron analizarse** — esos productos usaron valores por defecto. Error: {errores_gemini[0]}"
        if sin_imagen:
            filas = ", ".join(str(p.get("fila_sin_imagen", "?")) for p in sin_imagen)
            cbm_line += f"\n⚠️ **Sin imagen en filas:** {filas}"

    # ── Estado de validación ODOO ─────────────────────────────────────────────
    if st.session_state.odoo_conectado:
        n_skus_odoo  = len(st.session_state.odoo_skus)
        n_conflictos = len(st.session_state.resoluciones_conflictos)
        if n_conflictos > 0:
            cbm_line += f"\n🔀 **Validación ODOO:** {n_conflictos} conflicto(s) resueltos por el usuario"
        else:
            cbm_line += f"\n✅ **Validación ODOO:** sin conflictos detectados ({n_skus_odoo} SKUs comparados)"
    else:
        cbm_line += f"\n⚠️ **ODOO no conectado** — SKUs generados sin validar contra la base de datos"
    # Buscar el último mensaje del asistente y anexar el resumen al final
    for i in range(len(st.session_state.chat) - 1, -1, -1):
        if st.session_state.chat[i]["role"] == "assistant":
            c = st.session_state.chat[i]["content"]
            if isinstance(c, list):
                # Contenido rico (bloques_res de conflictos) → agregar como bloque de texto
                st.session_state.chat[i]["content"].append({"type": "text", "value": cbm_line})
            else:
                st.session_state.chat[i]["content"] += cbm_line
            break

    sku_info = f"{n_imgs} imágenes procesadas, {sum(1 for p in productos if p.get('sku'))} SKUs generados. " if n_imgs > 0 else ""
    primer_msg = (
        f"El usuario subió '{st.session_state.get('filename', '')}'. "
        f"El análisis y los {len(productos)} productos ya están listos. "
        f"{sku_info}"
        f"CBM total del cargamento: {total_cbm:.4f} ({'SUPERA los 70 CBM' if total_cbm > 70 else 'dentro de los 70 CBM'}). "
        f"Advertencias de datos: {advertencias if advertencias else 'ninguna'}. "
        "Preséntate brevemente, confirma cuántos productos encontraste "
        "y pregunta si puede proceder a generar el Excel."
    )
    system = build_system_prompt(
        analisis, productos, tipo_cambio, contenedor,
        respuestas_dudas=st.session_state.respuestas_dudas or None,
    )
    texto = llamar_agente(primer_msg, system, [])

    # Guardar en historial LangChain
    st.session_state.lc_messages.append(HumanMessage(content=primer_msg))
    st.session_state.lc_messages.append(AIMessage(content=texto))

    st.session_state.chat.append({"role": "assistant", "content": texto})


# ── Detectar nuevo archivo ─────────────────────────────────────────────────────
if archivo is not None:
    archivo_id = f"{archivo.name}_{archivo.size}"

    if archivo_id != st.session_state.archivo_id:
        # Preservar historial de chat y conexión ODOO al cambiar de archivo
        _chat_previo     = st.session_state.chat[:]
        _lc_previo       = st.session_state.lc_messages[:]
        _odoo_skus       = st.session_state.odoo_skus
        _odoo_productos  = st.session_state.odoo_productos
        _odoo_phashes    = st.session_state.odoo_phashes
        _odoo_conectado  = st.session_state.odoo_conectado
        for k in defaults:
            st.session_state[k] = defaults[k]
        st.session_state.chat            = _chat_previo
        st.session_state.lc_messages     = _lc_previo
        st.session_state.odoo_skus       = _odoo_skus
        st.session_state.odoo_productos  = _odoo_productos
        st.session_state.odoo_phashes    = _odoo_phashes
        st.session_state.odoo_conectado  = _odoo_conectado
        st.session_state.archivo_id  = archivo_id
        st.session_state["filename"] = archivo.name
        st.session_state.contenedor_val = extraer_contenedor(archivo.name)

        file_bytes = archivo.read()
        st.session_state.file_bytes = file_bytes

        with st.spinner("Analizando el formato del packing list..."):
            try:
                analisis = analizar_encabezados(file_bytes)
                st.session_state.analisis         = analisis
                _dr = analisis.get("dudas_relevantes", [])
                _dm = analisis.get("dudas_menores", [])
                # Normalizar: asegurar que son listas de dicts / listas de strings
                st.session_state.dudas_relevantes = [d for d in _dr if isinstance(d, dict)] if isinstance(_dr, list) else []
                st.session_state.dudas_menores    = [d for d in _dm if isinstance(d, str)]  if isinstance(_dm, list) else []

                if st.session_state.dudas_relevantes:
                    st.session_state.esperando_dudas = True

                    # ── Construir mensaje detallado de análisis para el historial ──
                    columnas  = analisis.get("columnas", {})
                    n_dudas   = len(st.session_state.dudas_relevantes)
                    CONFIANZA = {"alta": "✅", "media": "⚠️", "baja": "🔴", "no_encontrado": "❌"}
                    LABELS = {
                        "nombre_producto": "Nombre del producto",
                        "piezas_x_caja":  "Piezas por caja",
                        "largo_cm":       "Largo (cm)",
                        "ancho_cm":       "Ancho (cm)",
                        "alto_cm":        "Alto (cm)",
                        "cbm_por_pieza":     "CBM por pieza",
                        "cbm_master_carton": "CBM master carton",
                        "cbm_total_sku":     "CBM total SKU",
                        "precio_usd":     "Precio USD",
                    }

                    lineas_mapeo = []
                    for campo, label in LABELS.items():
                        _raw = columnas.get(campo, {})
                        info = _raw if isinstance(_raw, dict) else {}
                        conf = info.get("confianza", "no_encontrado")
                        enc  = info.get("encabezado_original") or "—"
                        mues = info.get("valor_muestra") or "—"
                        nota = f" *({info.get('nota')})*" if info.get("nota") else ""
                        lineas_mapeo.append(
                            f"| {CONFIANZA.get(conf,'❓')} **{label}** | `{enc}` | {mues}{nota} |"
                        )

                    lineas_dudas = []
                    for i, d in enumerate(st.session_state.dudas_relevantes):
                        lineas_dudas.append(f"**{i+1}.** {d['descripcion']}\n   → *{d['pregunta']}*")

                    lineas_menores = [f"- {dm}" for dm in st.session_state.dudas_menores]

                    msg_analisis = (
                        f"Analicé el archivo **{archivo.name}**.\n\n"
                        "**Mapeo de columnas detectado:**\n\n"
                        "| | Campo | Encabezado original | Valor muestra |\n"
                        "|---|---|---|---|\n"
                        + "\n".join(lineas_mapeo)
                        + (
                            f"\n\n**{n_dudas} duda{'s' if n_dudas != 1 else ''} relevante{'s' if n_dudas != 1 else ''} — respóndelas abajo antes de continuar:**\n\n"
                            + "\n\n".join(lineas_dudas)
                            if lineas_dudas else ""
                        )
                        + (
                            "\n\n**Avisos informativos:**\n" + "\n".join(lineas_menores)
                            if lineas_menores else ""
                        )
                    )
                    st.session_state.chat.append({"role": "assistant", "content": msg_analisis})
                else:
                    # Sin dudas relevantes → leer productos y arrancar chat directamente
                    # Auto-rellenar número de contenedor si el sidebar está vacío
                    _cont_excel = analisis.get("numero_contenedor", "")
                    if _cont_excel and not st.session_state.get("contenedor_val", ""):
                        st.session_state.contenedor_val = _cont_excel
                    if not st.session_state.odoo_conectado:
                        st.session_state.chat.append({
                            "role": "assistant",
                            "content": "⚠️ **ODOO no conectado** — los SKUs se generarán sin validar contra la base de datos. Conecta ODOO desde el panel izquierdo antes de subir el archivo para evitar duplicados."
                        })
                    productos, advertencias = leer_productos(
                        file_bytes, analisis["columnas"],
                        fila_encabezado=analisis.get("fila_encabezado", 1),
                    )
                    productos = corregir_cbm(productos, advertencias)
                    productos = corregir_cbm_inner(productos, advertencias, file_bytes, analisis["columnas"])
                    with st.spinner("Traduciendo y normalizando nombres..."):
                        productos = normalizar_nombres_productos(productos)
                    with st.spinner("Analizando clasificación de productos (padres, variantes, duplicados)..."):
                        propuestas = analizar_clasificacion_packing(file_bytes, productos)
                    st.session_state.clasificacion_propuestas  = propuestas
                    st.session_state.clasificacion_activa      = True
                    st.session_state.productos                 = productos
                    st.session_state.advertencias_productos    = advertencias

            except Exception as e:
                st.session_state.chat.append({"role": "assistant", "content": f"Error al analizar: {e}"})

        st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# MAIN CONTENT
# ══════════════════════════════════════════════════════════════════════════════

st.title("📦 FERRAFORME — Agente de Productos v2")

tab_pl, tab_agregar = st.tabs(["📦 Packing List", "➕ Agregar Productos"])

with tab_pl:
    # ── Historial de chat — siempre visible ───────────────────────────────────────
    for msg in st.session_state.chat:
        with st.chat_message(msg["role"]):
            content = msg["content"]
            if isinstance(content, str):
                st.markdown(content)
            else:
                # Lista de bloques: {"type": "text"|"image_bytes", ...}
                for bloque in content:
                    if not isinstance(bloque, dict):
                        st.markdown(str(bloque))
                        continue
                    if bloque["type"] == "text":
                        st.markdown(bloque["value"])
                    elif bloque["type"] == "image_bytes":
                        try:
                            st.image(
                                bloque["data"],
                                width=bloque.get("width", 160),
                                caption=bloque.get("caption", ""),
                            )
                        except Exception:
                            pass
                    elif bloque["type"] == "columns":
                        # Bloque de columnas: {"type": "columns", "items": [{"img": bytes, "caption": "...", "text": "..."}]}
                        items = bloque.get("items", [])
                        if items:
                            cols = st.columns(min(len(items), 3))
                            for ci, item in enumerate(items):
                                with cols[ci % 3]:
                                    if item.get("img"):
                                        try:
                                            st.image(item["img"], width=120, caption=item.get("caption", ""))
                                        except Exception:
                                            pass
                                    if item.get("text"):
                                        st.markdown(item["text"])

    # ── Botones de descarga en área principal ────────────────────────────────────
    _eb = st.session_state.get("excel_bytes")
    _mb = st.session_state.get("master_bytes")
    _pb = st.session_state.get("purchase_bytes")
    if _eb or _mb or _pb:
        _fn = st.session_state.get("filename", "packing_list.xlsx").replace(".xlsx", "")
        st.divider()
        _dcols = st.columns(2)
        with _dcols[0]:
            if _eb:
                st.download_button(
                    label="⬇️ Descargar Excel FERRAFORME",
                    data=_eb,
                    file_name=f"{_fn}_FERRAFORME.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary",
                )
        with _dcols[1]:
            if _mb:
                st.download_button(
                    label="⬇️ Descargar Excel Master Costos",
                    data=_mb,
                    file_name=f"{_fn}_MASTER.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

        # ── Reporte clasificación ─────────────────────────────────────────────
        _rb = st.session_state.get("clasificacion_reporte_bytes")
        if _rb:
            st.download_button(
                label="⬇️ Descargar Reporte Clasificación (bodega)",
                data=_rb,
                file_name=f"{_fn}_CLASIFICACION.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        # ── Purchase Order ────────────────────────────────────────────────────
        if _eb and _mb:
            _pcols = st.columns([1, 1])
            with _pcols[0]:
                if st.button("🛒 Crear Purchase", use_container_width=True,
                             help="Genera el Excel de Purchase Order listo para importar a Odoo"):
                    try:
                        _prods_po    = st.session_state.get("productos", [])
                        _tc_po       = st.session_state.get("tipo_cambio", 19.0)
                        _cc_po       = st.session_state.get("costo_contenedor", 525000.0)
                        _pb_new      = generar_excel_purchase(_prods_po, _tc_po, _cc_po)
                        st.session_state.purchase_bytes = _pb_new
                        st.rerun()
                    except Exception as _e:
                        st.error(f"Error al generar Purchase: {_e}")
            with _pcols[1]:
                if _pb:
                    st.download_button(
                        label="⬇️ Descargar Purchase Order",
                        data=_pb,
                        file_name=f"{_fn}_PURCHASE.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        type="primary",
                    )

    # ── Snapshot de duplicados/variantes (editable, visible tras confirmar) ──
    _snap = st.session_state.get("dup_snapshot")
    if _snap and not st.session_state.get("esperando_duplicados"):
        _snap_grupos = _snap.get("grupos", [])
        _snap_resp   = _snap.get("respuestas", {})
        _snap_imgs   = _snap.get("imagenes", {})
        _n_grupos    = len(_snap_grupos)
        _TIPO_LABEL_S = {
            "exacto":          "🔁 Duplicado exacto",
            "probable":        "❓ Muy similar",
            "similar":         "🈁 Posible variante",
            "nombre_similar":  "🔤 Nombre parecido",
            "variante_imagen": "🖼️ Imagen similar",
        }
        _DEC_OPTS = ["Mismo producto", "Variante", "Productos diferentes"]
        with st.expander(
            f"📋 Similares/variantes detectados: {_n_grupos} grupo(s) — {_snap.get('fecha', '')}",
            expanded=False,
        ):
            st.caption("Puedes corregir las decisiones si algo quedó mal y luego aplicar los cambios.")
            _nuevas_resp: dict = {}
            for _sg in _snap_grupos:
                _sgid   = str(_sg["id"])
                _sprods = _sg["productos"]
                _sp0    = _sprods[0]
                _sfilas = [
                    str(p.get("fila_excel_0idx", _sg["indices"][ci]) + 1)
                    for ci, p in enumerate(_sprods)
                ]
                _sdec_orig = _snap_resp.get(_sgid)
                if _sdec_orig == "diferente":
                    _def_idx = 2
                elif isinstance(_sdec_orig, dict) and _sdec_orig.get("tipo") == "mismo":
                    _def_idx = 0
                else:
                    _def_idx = 1

                _sc1, _sc2, _sc3 = st.columns([3, 5, 3])
                with _sc1:
                    _stcols = st.columns(min(len(_sprods), 4))
                    for _ci, (_sp, _stc) in enumerate(zip(_sprods[:4], _stcols)):
                        with _stc:
                            _sf0 = _sp.get("fila_excel_0idx")
                            _si  = _snap_imgs.get(_sf0) if _sf0 is not None else None
                            if _si:
                                st.image(_si["data"], width=55)
                            st.caption(f"F{_sfilas[_ci]}")
                with _sc2:
                    st.markdown(f"**{_sp0.get('nombre', '—')}**")
                    _scn = _sp0.get("nombre_chino_orig", "")
                    if _scn and _tiene_chino(_scn):
                        st.caption(f"🈁 {_scn}")
                    st.caption(
                        f"{_TIPO_LABEL_S.get(_sg['tipo'], _sg['tipo'])}  \n"
                        f"Filas: {', '.join(_sfilas)}"
                    )
                    if len(_sprods) > 1:
                        _sotros = [p.get("nombre", "—") for p in _sprods[1:4]]
                        st.caption("Con: " + " · ".join(f"`{n}`" for n in _sotros))
                with _sc3:
                    _nueva = st.selectbox(
                        "Decisión",
                        _DEC_OPTS,
                        index=_def_idx,
                        key=f"snap_dec_{_sgid}",
                        label_visibility="collapsed",
                    )
                    _nuevas_resp[_sgid] = _nueva
                st.divider()

            if st.button("🔄 Aplicar correcciones", key="snap_aplicar"):
                # Convertir opciones de texto a formato interno
                _resp_corr: dict = {}
                for _sgid2, _dec_txt in _nuevas_resp.items():
                    _gidx = next((i for i, g in enumerate(_snap_grupos) if str(g["id"]) == _sgid2), None)
                    if _gidx is None:
                        continue
                    _gx = _snap_grupos[_gidx]
                    if _dec_txt == "Productos diferentes":
                        _resp_corr[_sgid2] = "diferente"
                    elif _dec_txt == "Mismo producto":
                        _resp_corr[_sgid2] = {"tipo": "mismo", "sel": list(range(len(_gx["productos"])))}
                    else:
                        _resp_corr[_sgid2] = {"tipo": "variantes", "sel": list(range(len(_gx["productos"])))}
                # Re-entrar a paso 2 con las correcciones
                _prods_orig = _snap.get("productos_orig")
                if _prods_orig:
                    st.session_state.productos             = _prods_orig
                    st.session_state.duplicados_pendientes = _snap_grupos
                    st.session_state.imagenes_excel        = _snap_imgs
                    st.session_state.dup_respuestas        = _resp_corr
                    st.session_state.dup_paso              = 2
                    st.session_state.esperando_duplicados  = True
                    st.session_state.dup_snapshot          = None
                    st.rerun()

    # Scroll automático al último mensaje del chat
    if st.session_state.chat:
        st.components.v1.html(
            """<script>
            var main = window.parent.document.querySelector('section.main');
            if (main) main.scrollTo(0, main.scrollHeight);
            </script>""",
            height=0,
        )



# ── PASO: Resolver dudas relevantes ───────────────────────────────────────────
if st.session_state.esperando_dudas:
    with tab_pl:
        analisis         = st.session_state.analisis
        dudas_relevantes = st.session_state.dudas_relevantes
        dudas_menores    = st.session_state.dudas_menores

        st.subheader("El agente necesita aclarar algunas dudas antes de continuar")
        st.markdown(
            f"Se detectó el archivo en **{analisis.get('idioma_detectado', 'idioma desconocido')}**. "
            "Responde las siguientes preguntas para asegurar que el Excel se genere correctamente."
        )
        st.divider()

        respuestas_temp = {}

        for duda in dudas_relevantes:
            if not isinstance(duda, dict):
                continue
            duda_id = str(duda["id"])
            st.markdown(f"**{duda['id'] + 1}. {duda['descripcion']}**")

            tipo    = duda.get("tipo", "confirmar")
            opciones = duda.get("opciones", ["Sí, continuar", "No, revisar"])
            default  = duda.get("default", opciones[0])
            idx_default = opciones.index(default) if default in opciones else 0

            if tipo in ("eleccion", "confirmar"):
                respuesta = st.radio(
                    duda["pregunta"],
                    opciones,
                    index=idx_default,
                    key=f"duda_radio_{duda_id}",
                    horizontal=True,
                )
            elif tipo == "texto":
                respuesta = st.text_input(
                    duda["pregunta"],
                    value=duda.get("default", ""),
                    key=f"duda_texto_{duda_id}",
                )
            else:
                respuesta = default

            respuestas_temp[duda_id] = respuesta
            st.divider()

        # Dudas menores — solo informativas
        if dudas_menores:
            with st.expander(f"ℹ️ {len(dudas_menores)} avisos informativos (no bloquean la generación)"):
                for dm in dudas_menores:
                    st.markdown(f"- {dm}")

        col_a, col_b = st.columns(2)
        with col_a:
            if st.button("✅ Continuar con estas respuestas", type="primary", width="stretch"):
                st.session_state.respuestas_dudas = respuestas_temp

                # Mover aclaraciones al historial principal antes de continuar
                for m in st.session_state.chat_fase:
                    st.session_state.chat.append(m)
                st.session_state.chat_fase = []

                # Agregar resumen de respuestas del usuario al historial
                resumen_resp = "**Mis respuestas:**\n" + "\n".join(
                    f"- **{dudas_relevantes[int(k)].get('descripcion', '')[:60]}...** → {v}"
                    for k, v in respuestas_temp.items()
                    if int(k) < len(dudas_relevantes) and isinstance(dudas_relevantes[int(k)], dict)
                )
                st.session_state.chat.append({"role": "user", "content": resumen_resp})

                # Aplicar respuestas al mapeo de columnas
                columnas_actualizadas = aplicar_respuestas(
                    analisis["columnas"], dudas_relevantes, respuestas_temp
                )
                analisis_actualizado = {**analisis, "columnas": columnas_actualizadas}
                st.session_state.analisis        = analisis_actualizado
                st.session_state.esperando_dudas = False

                # Auto-rellenar número de contenedor si el sidebar está vacío
                _cont_excel = analisis_actualizado.get("numero_contenedor", "")
                if _cont_excel and not st.session_state.get("contenedor_val", ""):
                    st.session_state.contenedor_val = _cont_excel

                # Leer productos, enriquecer con imágenes y arrancar chat
                productos, advertencias = leer_productos(
                    st.session_state.file_bytes, columnas_actualizadas,
                    fila_encabezado=st.session_state.analisis.get("fila_encabezado", 1),
                )
                productos = corregir_cbm(productos, advertencias)
                productos = corregir_cbm_inner(productos, advertencias, st.session_state.file_bytes, columnas_actualizadas)
                with st.spinner("Traduciendo y normalizando nombres..."):
                    productos = normalizar_nombres_productos(productos)
                with st.spinner("Analizando clasificación de productos (padres, variantes, duplicados)..."):
                    propuestas = analizar_clasificacion_packing(st.session_state.file_bytes, productos)
                st.session_state.clasificacion_propuestas  = propuestas
                st.session_state.clasificacion_activa      = True
                st.session_state.productos                 = productos
                st.session_state.advertencias_productos    = advertencias

                st.rerun()

        with col_b:
            if st.button("↩️ Cancelar y revisar el archivo", width="stretch"):
                _chat_c  = st.session_state.chat[:]
                _lc_c    = st.session_state.lc_messages[:]
                for k in defaults:
                    st.session_state[k] = defaults[k]
                st.session_state.chat        = _chat_c
                st.session_state.lc_messages = _lc_c
                st.rerun()

        # ── Chat de aclaraciones (dentro de la fase de dudas) ─────────────────────
        if st.session_state.chat_fase:
            st.divider()
            for msg in st.session_state.chat_fase:
                with st.chat_message(msg["role"]):
                    st.markdown(msg["content"])


# ── PASO: Resolver filas duplicadas ──────────────────────────────────────────
if st.session_state.esperando_duplicados:
    with tab_pl:
        grupos_dup    = st.session_state.duplicados_pendientes
        productos_dup = st.session_state.productos
        imagenes_dup  = st.session_state.get("imagenes_excel", {})
        dup_paso      = st.session_state.get("dup_paso", 1)

        # ════════════════════════════════════════════════════════════════════════
        # PASO 1 — Auto-resolver exactos; pedir decisión solo para el resto
        # ════════════════════════════════════════════════════════════════════════
        if dup_paso == 1:
            # Todos los grupos son editables — los "exacto" tienen Mismo producto pre-seleccionado
            todos_grupos = grupos_dup

            st.subheader(f"🔍 {len(todos_grupos)} grupo(s) detectados — revisa y corrige si es necesario")
            st.caption("Puedes cambiar la decisión de cualquier grupo, incluso los que detectamos con alta confianza.")

            # Encabezados
            _hA, _hB, _hC, _hD = st.columns([3, 4, 3, 3])
            with _hA: st.markdown("**Imágenes**")
            with _hB: st.markdown("**Producto / Filas**")
            with _hC: st.markdown("**Detectado como**")
            with _hD: st.markdown("**Tu decisión**")
            st.divider()

            for grupo in todos_grupos:
                gid    = str(grupo["id"])
                prods  = grupo["productos"]
                p0     = prods[0]
                tipo_g = grupo["tipo"]
                n_p    = len(prods)
                filas  = [str(p.get("fila_excel_0idx", grupo["indices"][ci]) + 1)
                          for ci, p in enumerate(prods)]
                diffs  = grupo.get("diffs", {})

                tipo_badge = {
                    "exacto":          "✅ Mismo producto (alta confianza)",
                    "probable":        "❓ Posible mismo producto",
                    "similar":         "🈁 Posible variante (nombre chino similar)",
                    "nombre_similar":  "🔤 Nombres parecidos",
                    "variante_imagen": "🖼️ Imágenes similares",
                }.get(tipo_g, tipo_g)

                default_idx = {
                    "exacto":          0,  # Mismo producto
                    "probable":        0,  # Mismo producto
                    "similar":         1,  # Variante
                    "variante_imagen": 1,  # Variante
                    "nombre_similar":  2,  # Productos diferentes
                }.get(tipo_g, 2)

                col_imgs, col_info, col_tipo, col_dec = st.columns([3, 4, 3, 3])

                # ── Thumbnails ──────────────────────────────────────────────────
                with col_imgs:
                    n_thumb = min(n_p, 5)
                    tcols   = st.columns(n_thumb)
                    for ci in range(n_thumb):
                        with tcols[ci]:
                            fila_0 = prods[ci].get("fila_excel_0idx")
                            img_t  = imagenes_dup.get(fila_0) if fila_0 is not None else None
                            if img_t:
                                st.image(img_t["data"], width=65)
                            else:
                                st.caption("—")
                    if n_p > 5:
                        st.caption(f"+{n_p - 5} más")

                # ── Info ────────────────────────────────────────────────────────
                with col_info:
                    st.markdown(f"**{p0.get('nombre', '—')}**")
                    _cn_orig = p0.get("nombre_chino_orig", "")
                    if _cn_orig and _tiene_chino(_cn_orig):
                        st.caption(f"🈁 {_cn_orig}")
                    st.caption(f"Filas: {', '.join(filas)}")
                    if diffs:
                        diff_parts = []
                        for campo, vals in list(diffs.items())[:3]:
                            lbl = _LABELS_DIFF.get(campo, campo)
                            diff_parts.append(f"{lbl}: {' / '.join(str(v) for v in vals)}")
                        st.caption("↕ " + " · ".join(diff_parts))

                # ── Tipo detectado ───────────────────────────────────────────────
                with col_tipo:
                    st.caption(tipo_badge)

                # ── Decisión ────────────────────────────────────────────────────
                with col_dec:
                    dec = st.selectbox(
                        "Decisión",
                        ["Mismo producto", "Variante", "Productos diferentes"],
                        index=default_idx,
                        key=f"dup1_{gid}",
                        label_visibility="collapsed",
                    )

                # ── Sub-grupos ───────────────────────────────────────────────────
                # Mostrar siempre que haya más de 1 producto y la decisión no sea "diferentes"
                if dec != "Productos diferentes" and n_p > 1:
                    with st.expander(
                        f"✂️ Separar en sub-grupos ({n_p} productos) — haz clic si alguno no pertenece aquí",
                        expanded=False,
                    ):
                        st.caption(
                            "Mismo número → van juntos en el mismo SKU base  ·  "
                            "Número distinto → se tratan como grupos separados  ·  "
                            "Independiente → SKU completamente propio"
                        )
                        sub_cols = st.columns(min(n_p, 6))
                        for ci, prod_s in enumerate(prods):
                            with sub_cols[ci % 6]:
                                fila_0s = prod_s.get("fila_excel_0idx")
                                img_s   = imagenes_dup.get(fila_0s) if fila_0s is not None else None
                                if img_s:
                                    st.image(img_s["data"], width=90)
                                _nom_s = prod_s.get("nombre", f"Prod {ci+1}")
                                st.caption(f"F{filas[ci]}: {_nom_s[:25]}")
                                st.selectbox(
                                    "Sub-grupo",
                                    ["1", "2", "3", "4", "5", "Independiente"],
                                    index=0,
                                    key=f"dup_sub_{gid}_{ci}",
                                    label_visibility="collapsed",
                                )

                st.divider()

            # ── Botón de confirmación ──────────────────────────────────────────
            if st.button("✅ Confirmar — generar SKUs →", type="primary", width="stretch"):
                respuestas_paso1: dict = {}
                # Leer decisión de TODOS los grupos (incluyendo los antes "auto")
                for g in todos_grupos:
                    gid_b   = str(g["id"])
                    dec_val = st.session_state.get(f"dup1_{gid_b}", "Mismo producto")
                    if "Productos diferentes" in dec_val:
                        respuestas_paso1[gid_b] = "diferente"
                    else:
                        tipo = "mismo" if dec_val == "Mismo producto" else "variantes"
                        sub_map: dict[str, list[int]] = {}
                        for ci in range(len(g["productos"])):
                            sg = st.session_state.get(f"dup_sub_{gid_b}_{ci}", "1")
                            if sg != "Independiente":
                                sub_map.setdefault(sg, []).append(ci)
                        subgrupos = list(sub_map.values())
                        if len(subgrupos) <= 1:
                            respuestas_paso1[gid_b] = {"tipo": tipo, "sel": subgrupos[0] if subgrupos else list(range(len(g["productos"])))}
                        else:
                            respuestas_paso1[gid_b] = {"tipo": tipo, "subgrupos": subgrupos}
                st.session_state.dup_respuestas = respuestas_paso1
                st.session_state.dup_paso = 2
                st.rerun()

            # ── Mini-chat de aclaraciones (dentro del panel de duplicados) ────
            if st.session_state.chat_fase:
                st.divider()
                for _fcm in st.session_state.chat_fase:
                    with st.chat_message(_fcm["role"]):
                        st.markdown(_fcm["content"])

        # ════════════════════════════════════════════════════════════════════════
        # PASO 2 — Ejecutar resolución y proceder directamente con SKUs
        # ════════════════════════════════════════════════════════════════════════
        elif dup_paso == 2:
            respuestas_finales  = dict(st.session_state.get("dup_respuestas", {}))
            productos_resueltos = aplicar_resolucion_duplicados(
                productos_dup, grupos_dup, respuestas_finales
            )

            # Guardar historial
            st.session_state.historial_duplicados = (
                st.session_state.get("historial_duplicados", []) + [{
                    "fecha":      datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "archivo":    st.session_state.get("filename", "—"),
                    "grupos":     grupos_dup,
                    "respuestas": respuestas_finales,
                }]
            )

            # Proceder directamente a generar SKUs
            st.session_state.esperando_duplicados  = False
            st.session_state.dup_paso              = 1
            st.session_state.dup_respuestas        = {}
            st.session_state.duplicados_pendientes = []
            st.session_state.dup_grupos_segunda    = []
            st.session_state.chat_fase             = []

            analisis_res     = st.session_state.analisis
            advertencias_res = st.session_state.get("advertencias_productos", [])
            tipo_cambio_res  = st.session_state.get("tipo_cambio", 17.5)
            contenedor_res   = st.session_state.get("contenedor_val", "")

            with st.spinner("Analizando imágenes y generando SKUs..."):
                prods_enr, n_imgs_res, conflictos_res = enriquecer_con_imagenes(
                    st.session_state.file_bytes, productos_resueltos
                )
            st.session_state.productos              = prods_enr
            st.session_state.advertencias_productos = advertencias_res
            st.session_state.n_imgs_procesadas      = n_imgs_res

            # Guardar snapshot visual antes de limpiar imágenes
            import copy as _copy
            st.session_state.dup_snapshot = {
                "grupos":         grupos_dup,
                "respuestas":     respuestas_finales,
                "imagenes":       dict(st.session_state.get("imagenes_excel", {})),
                "productos_orig": _copy.deepcopy(productos_dup),  # pre-resolución
                "fecha":          datetime.now().strftime("%Y-%m-%d %H:%M"),
                "archivo":        st.session_state.get("filename", "—"),
            }
            st.session_state.imagenes_excel         = {}

            if conflictos_res:
                st.session_state.conflictos_pendientes = conflictos_res
                st.session_state.esperando_conflictos  = True
            else:
                _, errs_ren = renombrar_imagenes_con_sku(prods_enr)
                _reportar_errores_imagenes(errs_ren, "renombrado de imágenes")
                with st.spinner("Iniciando el agente..."):
                    _iniciar_chat(analisis_res, prods_enr, advertencias_res,
                                  tipo_cambio_res, contenedor_res, n_imgs_res)
            st.rerun()



# ── PASO: Resolver conflictos de SKU con ODOO ────────────────────────────────
if st.session_state.esperando_conflictos:
    with tab_pl:
        conflictos   = st.session_state.conflictos_pendientes
        analisis_cf  = st.session_state.analisis
        productos_cf = st.session_state.productos
        advertencias_cf = st.session_state.get("advertencias_productos", [])
        n_imgs_cf       = st.session_state.get("n_imgs_procesadas", 0)

        st.subheader(f"⚠️ Posibles duplicados detectados ({len(conflictos)})")
        st.markdown(
            "Los siguientes productos coinciden con productos ya existentes en ODOO "
            "por **prefijo SKU**, **imagen similar** o **nombre similar**. "
            "Indica si es el mismo producto o uno diferente."
        )
        st.divider()

        resoluciones_temp = dict(st.session_state.resoluciones_conflictos)

        for conflicto in conflictos:
            c_idx        = conflicto["idx"]
            c_nombre     = conflicto["nombre"]
            c_sku_prop   = conflicto["sku_propuesto"]
            c_sku_aj     = conflicto["sku_ajustado"]
            c_datos_gem  = conflicto.get("datos_gemini", {})
            c_prods_odoo = conflicto.get("productos_odoo", [])
            c_razon      = conflicto.get("razon", "sku")
            c_similares  = conflicto.get("similares_odoo", [])

            # Badge de razón
            razon_badge = {
                "sku":               "🔑 Mismo prefijo SKU",
                "imagen similar":    "🖼️ Imagen similar",
                "nombre similar":    "🔤 Nombre similar",
                "semántica similar": "🧠 Semántica similar (RAG)",
            }.get(c_razon, f"⚠️ {c_razon}")

            st.markdown(f"### Producto: **{c_nombre}**")
            st.markdown(f"**Razón:** {razon_badge}")
            st.markdown(f"- SKU generado: `{c_sku_prop}`")
            if c_sku_aj != c_sku_prop:
                st.markdown(f"- SKU alternativo (max+1): `{c_sku_aj}`")

            # Imagen del producto extraída del Excel
            prod_excel = productos_cf[c_idx] if c_idx < len(productos_cf) else {}
            stem_excel = prod_excel.get("imagen_temp_stem")
            if stem_excel and IMAGENES_TEMP_PATH.exists():
                archivo_img = next(
                    (f for f in IMAGENES_TEMP_PATH.iterdir() if f.stem == stem_excel),
                    None,
                )
                if archivo_img:
                    st.markdown("**Tu producto (del Excel):**")
                    st.image(str(archivo_img), width=160)

            # Scores de similitud si vienen de búsqueda visual/nombre
            if c_similares:
                for sim in c_similares[:3]:
                    badges = []
                    if sim.get("por_imagen") and sim.get("similitud_imagen") is not None:
                        pct = int(sim["similitud_imagen"] * 100)
                        badges.append(f"🖼️ imagen {pct}%")
                    if sim.get("por_nombre"):
                        pct = int(sim["similitud_nombre"] * 100)
                        badges.append(f"🔤 nombre {pct}%")
                    if sim.get("por_rag"):
                        pct = int(sim["similitud"] * 100)
                        badges.append(f"🧠 RAG {pct}%")
                    st.caption(f"`{sim['sku']}` — {sim['nombre']} — {' · '.join(badges)}")

            if c_prods_odoo:
                st.markdown("**Productos existentes en ODOO con ese prefijo:**")
                n_cols = min(len(c_prods_odoo), 3)
                cols_odoo = st.columns(n_cols)
                for ci, prod_odoo in enumerate(c_prods_odoo):
                    with cols_odoo[ci % n_cols]:
                        if prod_odoo.get("image_128"):
                            try:
                                img_bytes_odoo = base64.b64decode(prod_odoo["image_128"])
                                st.image(img_bytes_odoo, width=120)
                            except Exception:
                                pass
                        categ_raw  = prod_odoo.get("categ_id", [None, "—"])
                        categ_name = categ_raw[1] if isinstance(categ_raw, (list, tuple)) and len(categ_raw) > 1 else str(categ_raw)
                        st.markdown(
                            f"**SKU:** `{prod_odoo.get('default_code', '—')}`  \n"
                            f"**Nombre:** {prod_odoo.get('name', '—')}  \n"
                            f"**Categoría:** {categ_name}  \n"
                            f"**Descripción:** {prod_odoo.get('description_sale') or '—'}  \n"
                            f"**Precio:** ${prod_odoo.get('list_price', 0):.2f}"
                        )

            decision = st.radio(
                "¿Este producto nuevo es el mismo que alguno de arriba?",
                ["Diferente — asignar SKU nuevo", "Mismo producto — reutilizar SKU de ODOO"],
                key=f"conflicto_tipo_{c_idx}",
                horizontal=True,
            )

            res = resoluciones_temp.setdefault(str(c_idx), {})

            if "Mismo" in decision and c_prods_odoo:
                opciones_skus_odoo = [p.get("default_code", "") for p in c_prods_odoo]
                sku_sel = st.selectbox(
                    "SKU de ODOO a reutilizar:",
                    opciones_skus_odoo,
                    key=f"conflicto_sku_sel_{c_idx}",
                )
                prod_sel = next((p for p in c_prods_odoo if p.get("default_code") == sku_sel), {})
                categ_sel_raw  = prod_sel.get("categ_id", [None, ""])
                categ_sel_name = categ_sel_raw[1] if isinstance(categ_sel_raw, (list, tuple)) and len(categ_sel_raw) > 1 else str(categ_sel_raw)

                res["tipo"]        = "mismo"
                res["sku_final"]   = sku_sel
                res["nombre"]      = st.text_input(
                    "Nombre del producto:", value=prod_sel.get("name", c_nombre),
                    key=f"conflicto_nom_{c_idx}"
                )
                res["descripcion"] = st.text_input(
                    "Descripción:", value=prod_sel.get("description_sale") or c_datos_gem.get("descripcion", ""),
                    key=f"conflicto_desc_{c_idx}"
                )
                res["categoria"]   = st.text_input(
                    "Categoría:", value=categ_sel_name or c_datos_gem.get("categoria", ""),
                    key=f"conflicto_cat_{c_idx}"
                )
                res["atributo"]    = st.text_input(
                    "Atributo:", value=c_datos_gem.get("atributo_desc", ""),
                    key=f"conflicto_attr_{c_idx}"
                )
            else:
                res["tipo"]      = "diferente"
                res["sku_final"] = c_sku_aj

            st.divider()

        # ── Mini-chat de aclaraciones (dentro del panel de conflictos) ───────
        if st.session_state.chat_fase:
            st.divider()
            for _fcm in st.session_state.chat_fase:
                with st.chat_message(_fcm["role"]):
                    st.markdown(_fcm["content"])

        if st.button("✅ Confirmar todas las resoluciones", type="primary", width="stretch"):
            st.session_state.resoluciones_conflictos = resoluciones_temp
            productos_finales = aplicar_resoluciones_conflictos(
                productos_cf, conflictos, resoluciones_temp
            )
            st.session_state.productos             = productos_finales
            st.session_state.esperando_conflictos  = False
            st.session_state.conflictos_pendientes = []
            st.session_state.chat_fase             = []

            # SKUs definitivos confirmados → renombrar imágenes temp
            _, errs_ren = renombrar_imagenes_con_sku(productos_finales)
            _reportar_errores_imagenes(errs_ren, "renombrado de imágenes")


            # Guardar resumen rico de resoluciones en el historial del chat
            bloques_res = [{"type": "text", "value": f"**Resolución de {len(conflictos)} conflicto(s) SKU:**"}]
            for c in conflictos:
                c_idx_r    = c["idx"]
                res        = resoluciones_temp.get(str(c_idx_r), {})
                nombre_r   = c.get("nombre", f"Producto {c_idx_r + 1}")
                razon_r    = {"sku": "🔑 Mismo prefijo SKU", "imagen similar": "🖼️ Imagen similar",
                              "nombre similar": "🔤 Nombre similar", "semántica similar": "🧠 Semántica similar"
                              }.get(c.get("razon", "sku"), f"⚠️ {c.get('razon', '')}")
                if res.get("tipo") == "mismo":
                    decision_txt = f"✅ Reutilizado SKU `{res.get('sku_final', '')}` de ODOO"
                else:
                    decision_txt = f"🆕 SKU nuevo asignado: `{c['sku_ajustado']}`"
                bloques_res.append({"type": "text", "value": f"---\n### {nombre_r}\n**Razón:** {razon_r}  \n{decision_txt}"})

                # Imagen del producto del Excel (aún en carpeta temp antes de borrar)
                prod_r    = productos_cf[c_idx_r] if c_idx_r < len(productos_cf) else {}
                stem_r    = prod_r.get("imagen_temp_stem")
                img_excel = None
                if stem_r and IMAGENES_TEMP_PATH.exists():
                    arch_r = next((f for f in IMAGENES_TEMP_PATH.iterdir() if f.stem == stem_r), None)
                    if arch_r:
                        try:
                            img_excel = arch_r.read_bytes()
                        except Exception:
                            pass

                # Imágenes de ODOO similares
                odoo_items = []
                for prod_odoo in c.get("productos_odoo", [])[:3]:
                    item = {
                        "text": (
                            f"**SKU:** `{prod_odoo.get('default_code', '—')}`  \n"
                            f"**Nombre:** {prod_odoo.get('name', '—')}  \n"
                            f"**Precio:** ${prod_odoo.get('list_price', 0):.2f}"
                        ),
                        "img": None,
                        "caption": prod_odoo.get("default_code", ""),
                    }
                    if prod_odoo.get("image_128"):
                        try:
                            item["img"] = base64.b64decode(prod_odoo["image_128"])
                        except Exception:
                            pass
                    odoo_items.append(item)

                # Añadir imagen del Excel y columnas de ODOO al bloque
                if img_excel:
                    bloques_res.append({"type": "text", "value": "**Tu producto (del Excel):**"})
                    bloques_res.append({"type": "image_bytes", "data": img_excel, "width": 160, "caption": nombre_r})
                if odoo_items:
                    bloques_res.append({"type": "text", "value": "**Productos similares en ODOO:**"})
                    bloques_res.append({"type": "columns", "items": odoo_items})

            st.session_state.chat.append({"role": "assistant", "content": bloques_res})

            with st.spinner("Iniciando el agente..."):
                _iniciar_chat(analisis_cf, productos_finales, advertencias_cf, tipo_cambio, contenedor, n_imgs_cf)
            st.rerun()


# ── PASO: Revisar clasificación padre/variante ────────────────────────────────
if st.session_state.get("clasificacion_activa"):
    with tab_pl:
        _props      = st.session_state.clasificacion_propuestas
        _n_props    = len(_props)
        _analisis_c = st.session_state.get("analisis", {})
        _advertencias_c = st.session_state.get("advertencias_productos", [])
        _productos_c    = st.session_state.get("productos", [])

        st.subheader(f"🗂️ Clasificación de {_n_props} producto(s) — revisa y confirma")
        st.caption(
            "El agente analizó cada producto y propuso cómo clasificarlos en Odoo. "
            "Marca **Requiere revisión** en los que quieras que bodega revise manualmente."
        )

        # ── Estado de Odoo y modo prueba ──────────────────────────────────────
        _cls_odoo_ok = st.session_state.get("odoo_conectado", False)
        _cls_prueba  = st.session_state.get("modo_prueba", False)
        if _cls_odoo_ok:
            if _cls_prueba:
                st.info("🧪 **Modo prueba activado** — al confirmar se crearán productos con sufijo `_test` en categoría `PRUEBAS_AGENTE`")
            else:
                st.success("✅ **Odoo conectado** — al confirmar se crearán los productos en Odoo")
        else:
            st.warning("⚠️ **Odoo no conectado** — al confirmar solo se generará el reporte, no se crearán productos en Odoo")

        # ── Resumen rápido ────────────────────────────────────────────────────
        _acc_counts: dict[str, int] = {}
        for _p in _props:
            _acc_counts[_p["accion"]] = _acc_counts.get(_p["accion"], 0) + 1
        _sum_cols = st.columns(5)
        for _ci, (_acc, _lbl) in enumerate([
            ("crear_padre_y_variante", "🆕 Padre+variante"),
            ("crear_padre_solo",       "🆕 Solo padre"),
            ("crear_variante",         "➕ Variante"),
            ("reutilizar",             "♻️ Reutilizar"),
            ("duplicado",              "🔁 Duplicado"),
        ]):
            with _sum_cols[_ci]:
                st.metric(_lbl, _acc_counts.get(_acc, 0))

        st.divider()

        # ── Encabezados de columnas ───────────────────────────────────────────
        _odoo_urls = st.session_state.get("odoo_imagen_urls", {})
        _hc = st.columns([1, 1, 4, 2, 1, 2])
        for _ht, _hcol in zip(
            ["🖼️ Producto", "🏷️ Padre Odoo", "Nombre / SKU / Acción", "Atributo", "⚠️ Rev.", "Nota bodega"],
            _hc,
        ):
            _hcol.markdown(f"**{_ht}**")
        st.divider()

        # ── Fila por producto ─────────────────────────────────────────────────
        _ACCION_COLOR = {
            "crear_padre_y_variante": "🟦",
            "crear_padre_solo":       "🟦",
            "crear_variante":         "🟩",
            "reutilizar":             "🟨",
            "duplicado":              "🟥",
        }
        for _i, _prop in enumerate(_props):
            _prod_p  = _prop["producto"]
            _fila    = str((_prod_p.get("fila_excel_0idx") or 0) + 1)
            _accion  = _prop["accion"]
            _badge   = _ACCION_COLOR.get(_accion, "⬜")
            _c_img, _c_padre, _c_info, _c_att, _c_rev, _c_nota = st.columns([1, 1, 4, 2, 1, 2])

            # ── Imagen del producto (del Excel) ───────────────────────────────
            with _c_img:
                _img_path = _prop.get("imagen_temp_path")
                if _img_path and Path(_img_path).exists():
                    try:
                        st.image(_img_path, width=80, caption=f"F{_fila}")
                    except Exception:
                        st.caption(f"F{_fila}")
                else:
                    st.caption(f"F{_fila}\n_(sin img)_")

            # ── Imagen del padre en Odoo ──────────────────────────────────────
            with _c_padre:
                _padre_url = None
                if _prop.get("padre_fuente") == "odoo" and _prop.get("padre_sku"):
                    _padre_url = _odoo_urls.get(_prop["padre_sku"])
                    # También intentar con el SKU base (sin atributo) en caso de variante
                    if not _padre_url:
                        _partes_sku = _prop["padre_sku"].split("-")
                        if len(_partes_sku) >= 2:
                            _sku_base = "-".join(_partes_sku[:2])
                            _padre_url = _odoo_urls.get(_sku_base)
                if _padre_url:
                    try:
                        st.image(_padre_url, width=80, caption=_prop.get("padre_odoo_nombre", "")[:20])
                    except Exception:
                        st.caption(_prop.get("padre_odoo_nombre", "—"))
                elif _prop.get("padre_fuente") == "local":
                    st.caption("📋 Local\n(mismo PL)")
                else:
                    st.caption("—")

            # ── Info del producto ─────────────────────────────────────────────
            with _c_info:
                st.markdown(
                    f"**{_prop['nombre']}**  \n"
                    f"Base: *{_prop['nombre_base']}*  \n"
                    f"`{_prop['sku']}` {_badge} {_prop['accion_display']}"
                )
                if _prop.get("padre_sku") and _prop["padre_sku"] != _prop["sku"]:
                    st.caption(f"Padre: `{_prop['padre_sku']}`")

            # ── Atributo ──────────────────────────────────────────────────────
            with _c_att:
                if _prop.get("atributo_tipo") and _prop.get("atributo_valor"):
                    st.caption(f"{_prop['atributo_tipo']}  \n**{_prop['atributo_valor']}**")
                else:
                    st.caption("—")

            # ── Checkbox requiere revisión ────────────────────────────────────
            with _c_rev:
                st.checkbox(
                    "",
                    value=_prop.get("requiere_revision", False),
                    key=f"_cls_rev_{_i}",
                    label_visibility="collapsed",
                )

            # ── Nota ──────────────────────────────────────────────────────────
            with _c_nota:
                st.text_input(
                    "",
                    value=_prop.get("nota_revision", ""),
                    key=f"_cls_nota_{_i}",
                    placeholder="Observación...",
                    label_visibility="collapsed",
                )

            st.divider()

        # ── Botones de acción ─────────────────────────────────────────────────
        _btn_c1, _btn_c2 = st.columns([3, 1])
        with _btn_c1:
            _confirmar = st.button(
                "✅ Confirmar clasificación y continuar",
                type="primary",
                use_container_width=True,
            )
        with _btn_c2:
            if st.session_state.get("clasificacion_reporte_bytes"):
                st.download_button(
                    "⬇️ Reporte bodega",
                    data=st.session_state.clasificacion_reporte_bytes,
                    file_name=f"{st.session_state.get('filename', 'packing').replace('.xlsx', '')}_CLASIFICACION.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

        if _confirmar:
            # Leer valores de los widgets individuales
            for _i, _prop in enumerate(_props):
                _prop["requiere_revision"] = bool(st.session_state.get(f"_cls_rev_{_i}", False))
                _prop["nota_revision"]     = str(st.session_state.get(f"_cls_nota_{_i}", "") or "")

            # ── Crear padres y variantes en Odoo ─────────────────────────────
            _odoo_ok = st.session_state.get("odoo_conectado", False)
            _res_odoo: list[str] = []
            _err_odoo: list[str] = []
            if _odoo_ok:
                with st.spinner("Creando productos en Odoo..."):
                    _res_odoo, _err_odoo = crear_clasificacion_en_odoo(
                        _props, tipo_cambio,
                        st.session_state.get("costo_contenedor", 525000.0),
                    )

            # ── Generar reporte de clasificación (para bodega) ────────────────
            _nombre_pl = st.session_state.get("filename", "")
            _reporte_bytes = generar_reporte_clasificacion(_props, _nombre_pl)
            st.session_state.clasificacion_reporte_bytes = _reporte_bytes

            # ── Mensaje resumen en el chat ────────────────────────────────────
            _n_rev = sum(1 for p in _props if p["requiere_revision"])
            _res_lines = [f"**Clasificación confirmada — {_n_props} producto(s)**"]
            _acc_display = {
                "crear_padre_y_variante": "🆕 Padre nuevo + variante nueva",
                "crear_padre_solo":       "🆕 Padre nuevo (sin variante)",
                "crear_variante":         "➕ Variante nueva en padre existente",
                "reutilizar":             "♻️ Reutilizado desde Odoo",
                "duplicado":              "🔁 Duplicado (se ignorará)",
            }
            for _acc, _cnt in sorted(_acc_counts.items()):
                if _cnt:
                    _res_lines.append(f"- {_acc_display.get(_acc, _acc)}: **{_cnt}**")
            if _n_rev:
                _res_lines.append(
                    f"\n⚠️ **{_n_rev} producto(s) marcados para revisión** — creados en Odoo con tag 'Requiere Revisión' e incluidos en el reporte.")
            if _res_odoo:
                _res_lines.append(f"\n**Odoo ({len(_res_odoo)} creados):**")
                _res_lines.extend(_res_odoo[:10])  # máx 10 líneas para no saturar el chat
                if len(_res_odoo) > 10:
                    _res_lines.append(f"  … y {len(_res_odoo) - 10} más.")
            if _err_odoo:
                _res_lines.append(f"\n**Errores Odoo ({len(_err_odoo)}):**")
                _res_lines.extend(_err_odoo[:5])
            elif not _odoo_ok:
                _res_lines.append("\n⚠️ Odoo no conectado — productos no creados en el sistema.")

            st.session_state.chat.append({"role": "assistant", "content": "\n".join(_res_lines)})

            # Limpiar estado de clasificación y arrancar chat
            st.session_state.clasificacion_activa = False

            # Renombrar imágenes temp con los SKUs definitivos
            _, _errs_ren_c = renombrar_imagenes_con_sku(_productos_c)
            _reportar_errores_imagenes(_errs_ren_c, "renombrado de imágenes")

            with st.spinner("Iniciando el agente..."):
                _n_imgs_c = st.session_state.get("n_imgs_procesadas", 0)
                _iniciar_chat(_analisis_c, _productos_c, _advertencias_c,
                              tipo_cambio, contenedor, _n_imgs_c)
            st.rerun()


# ── PASO: Agregar productos en lote ──────────────────────────────────────────
if st.session_state.get("agregar_lote_activo"):
    with tab_pl:
        lote_paso = st.session_state.get("agregar_lote_paso", 1)

        # ════════ PASO 1 — Subir imágenes y nombres ═══════════════════════════════
        if lote_paso == 1:
            st.subheader("➕ Agregar productos nuevos")
            st.markdown("Fotografías de productos que no venían en el packing list original.")

            n_lote = st.number_input(
                "¿Cuántos productos deseas agregar?",
                min_value=1, max_value=30, value=st.session_state.get("agregar_lote_n", 1), step=1,
                key="lote_n_input",
            )

            imgs_subidas = st.file_uploader(
                f"Sube hasta {int(n_lote)} imagen(es) de los productos",
                type=["png", "jpg", "jpeg", "webp"],
                accept_multiple_files=True,
                key="lote_imgs_uploader",
            )

            # Leer bytes una sola vez antes de cualquier rerun
            imgs_data_preload = []
            if imgs_subidas:
                for f in imgs_subidas[: int(n_lote)]:
                    ext = Path(f.name).suffix.lstrip(".").lower() or "jpg"
                    imgs_data_preload.append({"data": f.read(), "ext": ext})

            # Previsualizar + nombres opcionales
            nombres_lote: list[str] = []
            if imgs_data_preload:
                st.markdown(f"**{len(imgs_data_preload)} imagen(es) cargadas:**")
                n_cols = min(len(imgs_data_preload), 4)
                preview_cols = st.columns(n_cols)
                for i, img_info in enumerate(imgs_data_preload):
                    with preview_cols[i % n_cols]:
                        st.image(img_info["data"], width=130)
                        nom = st.text_input(
                            f"Nombre {i + 1}",
                            placeholder=f"Producto {i + 1}",
                            key=f"lote_nombre_{i}",
                        )
                        nombres_lote.append(nom)

            col_proc, col_cancel = st.columns(2)
            with col_proc:
                if st.button(
                    "▶️ Procesar imágenes",
                    type="primary",
                    width="stretch",
                    disabled=not imgs_data_preload,
                ):
                    imgs_dict: dict = {}
                    for i, img_info in enumerate(imgs_data_preload):
                        nom = nombres_lote[i].strip() if i < len(nombres_lote) else ""
                        imgs_dict[i] = {"data": img_info["data"], "ext": img_info["ext"], "nombre": nom}
                    st.session_state.agregar_lote_imgs = imgs_dict
                    st.session_state.agregar_lote_n    = len(imgs_dict)
                    st.session_state.agregar_lote_paso = 2
                    st.rerun()
            with col_cancel:
                if st.button("✖️ Cancelar", width="stretch"):
                    st.session_state.agregar_lote_activo = False
                    st.rerun()

        # ════════ PASO 2 — Procesar con Vision + ODOO (automático) ════════════════
        elif lote_paso == 2:
            imgs_dict    = st.session_state.get("agregar_lote_imgs", {})
            skus_odoo    = st.session_state.get("odoo_skus", [])
            prods_odoo   = st.session_state.get("odoo_productos", [])
            phashes_odoo = st.session_state.get("odoo_phashes", {})

            with st.spinner(f"Analizando {len(imgs_dict)} imagen(es) con Vision y comparando con ODOO..."):
                prods_nuevos: list[dict]  = []
                conflictos_nuevos: list[dict] = []
                n_manuales_base = sum(1 for p in st.session_state.get("productos", []) if p.get("_manual"))

                for i, img_info in imgs_dict.items():
                    fila_0idx     = -(n_manuales_base + i + 1)
                    nombre_manual = img_info.get("nombre", "").strip() or f"Producto {i + 1}"

                    prod: dict = {
                        "nombre":          nombre_manual,
                        "fila_excel_0idx": fila_0idx,
                        "_manual":         True,
                    }

                    datos = analizar_imagen_claude(img_info["data"], img_info["ext"], contexto=prod)
                    sub_cod = datos.get("subcategoria_cod", "VAR")
                    att_cod = datos.get("atributo_cod", "GEN")
                    sku_inicial = generar_sku(sub_cod, att_cod)

                    prod.update({
                        "categoria":     datos.get("categoria", ""),
                        "subcategoria":  datos.get("subcategoria", ""),
                        "atributo_desc": datos.get("atributo_desc", ""),
                        "titulo":        datos.get("titulo", nombre_manual),
                        "descripcion":   datos.get("descripcion", ""),
                    })

                    conflicto_entry = None

                    if skus_odoo:
                        validacion = validar_sku_vs_odoo(sku_inicial, skus_odoo)
                        sku_final  = validacion["sku_ajustado"]
                        if validacion["conflicto"]:
                            conflicto_entry = {
                                "idx":           i,
                                "nombre":        nombre_manual,
                                "sku_propuesto": sku_inicial,
                                "sku_ajustado":  sku_final,
                                "datos_gemini":  datos,
                                "productos_odoo": [],
                                "razon":         "sku",
                                "similares_odoo": [],
                            }
                    else:
                        sku_final = sku_inicial

                    if prods_odoo:
                        similares = buscar_similares_odoo(
                            img_info["data"], nombre_manual, prods_odoo, phashes_odoo
                        )
                        if similares and conflicto_entry is None:
                            razones = []
                            if any(s.get("por_imagen") for s in similares):
                                razones.append("imagen similar")
                            if any(s.get("por_nombre") for s in similares):
                                razones.append("nombre similar")
                            conflicto_entry = {
                                "idx":           i,
                                "nombre":        nombre_manual,
                                "sku_propuesto": sku_inicial,
                                "sku_ajustado":  sku_inicial,
                                "datos_gemini":  datos,
                                "razon":         ", ".join(razones) or "imagen similar",
                                "similares_odoo": similares,
                                "productos_odoo": [s["producto_odoo"] for s in similares if s.get("producto_odoo")],
                            }
                        elif similares and conflicto_entry is not None:
                            # Enriquecer conflicto por SKU con similares encontrados
                            conflicto_entry["similares_odoo"] = similares
                            conflicto_entry["productos_odoo"] = [s["producto_odoo"] for s in similares if s.get("producto_odoo")]

                    prod["sku"] = sku_final
                    if st.session_state.get("modo_prueba"):
                        prod["sku"] += "_test"
                    prod["_img_data"] = img_info["data"]
                    prod["_img_ext"]  = img_info["ext"]

                    prods_nuevos.append(prod)
                    if conflicto_entry:
                        conflictos_nuevos.append(conflicto_entry)

            st.session_state.agregar_lote_prods = prods_nuevos
            st.session_state.agregar_lote_conf  = conflictos_nuevos
            st.session_state.agregar_lote_paso  = 3
            st.rerun()

        # ════════ PASO 3 — Revisar resultados y confirmar ═════════════════════════
        elif lote_paso == 3:
            prods_nuevos    = st.session_state.get("agregar_lote_prods", [])
            conflictos_lote = st.session_state.get("agregar_lote_conf", [])

            st.subheader(f"✅ {len(prods_nuevos)} producto(s) listos para agregar")

            # Preview de productos
            n_cols_prev = min(len(prods_nuevos), 4) or 1
            prev_cols   = st.columns(n_cols_prev)
            for i, prod in enumerate(prods_nuevos):
                with prev_cols[i % n_cols_prev]:
                    img_d = prod.get("_img_data")
                    if img_d:
                        st.image(img_d, width=130)
                    sku_badge = f"`{prod.get('sku', '—')}`"
                    st.markdown(f"**{prod.get('nombre', '—')}**  \nSKU: {sku_badge}")
                    if prod.get("categoria"):
                        st.caption(prod["categoria"])

            # Resolución de conflictos ODOO (si los hay)
            resoluciones_lote: dict = {}
            if conflictos_lote:
                st.divider()
                st.warning(f"⚠️ {len(conflictos_lote)} posible(s) coincidencia(s) con productos en ODOO — revisa antes de confirmar.")
                for conf in conflictos_lote:
                    c_idx      = conf["idx"]
                    c_nom      = conf["nombre"]
                    c_razon    = conf.get("razon", "")
                    c_similares = conf.get("similares_odoo", [])
                    c_sku_aj   = conf.get("sku_ajustado", conf.get("sku_propuesto", ""))

                    razon_badge = {
                        "sku":               "🔑 Mismo prefijo SKU",
                        "imagen similar":    "🖼️ Imagen similar",
                        "nombre similar":    "🔤 Nombre similar",
                        "semántica similar": "🧠 Semántica similar",
                    }.get(c_razon, f"⚠️ {c_razon}")

                    st.markdown(f"**{c_nom}** — {razon_badge}")

                    # Mostrar productos similares de ODOO
                    if c_similares:
                        sim_cols = st.columns(min(len(c_similares), 3))
                        for si, sim in enumerate(c_similares[:3]):
                            with sim_cols[si]:
                                _img_url = st.session_state.get("odoo_imagen_urls", {}).get(sim["sku"])
                                if _img_url:
                                    st.image(_img_url, width=100)
                                st.caption(
                                    f"SKU: `{sim['sku']}`  \n"
                                    f"{sim.get('nombre', '—')}"
                                )

                    dec_conf = st.radio(
                        "¿Qué hacemos con este producto?",
                        [
                            f"Asignar SKU nuevo `{c_sku_aj}` (es un producto diferente)",
                            "Reutilizar SKU de ODOO (es el mismo producto)",
                        ],
                        key=f"lote_conf_dec_{c_idx}",
                    )
                    if "Reutilizar" in dec_conf and c_similares:
                        sim_options = [s["sku"] for s in c_similares[:5]]
                        sku_sel = st.selectbox(
                            "SKU de ODOO a reutilizar:",
                            sim_options,
                            key=f"lote_conf_sku_{c_idx}",
                        )
                        resoluciones_lote[c_idx] = {"tipo": "mismo", "sku": sku_sel}
                    else:
                        resoluciones_lote[c_idx] = {"tipo": "diferente", "sku": c_sku_aj}

                    st.divider()

            col_conf, col_back = st.columns(2)
            with col_conf:
                if st.button("✅ Confirmar y agregar al Excel", type="primary", width="stretch"):
                    # Aplicar resoluciones de conflicto
                    for conf in conflictos_lote:
                        c_idx = conf["idx"]
                        res   = resoluciones_lote.get(c_idx, {"tipo": "diferente", "sku": conf.get("sku_ajustado", "")})
                        if res["tipo"] == "mismo":
                            prods_nuevos[c_idx]["sku"] = res["sku"]

                    # Agregar productos a la sesión + guardar imágenes temp
                    prods_agregados = []
                    for prod in prods_nuevos:
                        img_d = prod.pop("_img_data", None)
                        img_e = prod.pop("_img_ext", None)
                        if img_d and img_e:
                            try:
                                stem = f"manual_{abs(prod['fila_excel_0idx'])}"
                                IMAGENES_TEMP_PATH.mkdir(parents=True, exist_ok=True)
                                (IMAGENES_TEMP_PATH / f"{stem}.{img_e}").write_bytes(img_d)
                                prod["imagen_temp_stem"] = stem
                            except Exception:
                                pass
                        st.session_state.productos.append(prod)
                        prods_agregados.append(prod)

                    # Limpiar estado del lote
                    st.session_state.agregar_lote_activo = False
                    st.session_state.agregar_lote_paso   = 1
                    st.session_state.agregar_lote_imgs   = {}
                    st.session_state.agregar_lote_prods  = []
                    st.session_state.agregar_lote_conf   = []

                    # Notificar al agente
                    _msg_lote = (
                        f"Agregué {len(prods_agregados)} producto(s) que no venían en el packing list: "
                        + "; ".join(
                            f"'{p.get('nombre', '?')}' (SKU: `{p.get('sku', '?')}`)"
                            for p in prods_agregados
                        )
                        + ". Inclúyelos en el conteo y en el Excel cuando lo generes."
                    )
                    _sys_lote = build_system_prompt(
                        st.session_state.analisis,
                        st.session_state.productos,
                        st.session_state.get("tipo_cambio", 17.5),
                        st.session_state.get("contenedor_val", ""),
                        respuestas_dudas=st.session_state.respuestas_dudas or None,
                    )
                    st.session_state.chat.append({"role": "user", "content": _msg_lote})
                    with st.spinner("Notificando al agente..."):
                        _resp_lote = llamar_agente(_msg_lote, _sys_lote, st.session_state.lc_messages)
                    st.session_state.lc_messages.append(HumanMessage(content=_msg_lote))
                    st.session_state.lc_messages.append(AIMessage(content=_resp_lote))
                    st.session_state.chat.append({"role": "assistant", "content": _resp_lote})
                    st.rerun()

            with col_back:
                if st.button("↩️ Volver y cambiar imágenes", width="stretch"):
                    st.session_state.agregar_lote_paso  = 1
                    st.session_state.agregar_lote_prods = []
                    st.session_state.agregar_lote_conf  = []
                    st.rerun()



with tab_pl:
    # ── Chat input — siempre accesible cuando hay un archivo cargado ───────────────
    if st.session_state.analisis is not None:
            if st.session_state.esperando_dudas:
                placeholder = "Pregúntale al agente sobre estas dudas..."
            elif st.session_state.esperando_duplicados:
                placeholder = "Haz una aclaración o escribe tus cambios..."
            elif st.session_state.esperando_conflictos:
                placeholder = "Haz una aclaración o escribe tus cambios..."
            elif st.session_state.get("clasificacion_activa"):
                placeholder = "Haz una aclaración sobre la clasificación..."
            else:
                placeholder = "Escríbele al agente..."

            # ── Botones de respuesta rápida (solo en chat principal, no en pasos intermedios) ──
            _en_chat_principal = (
                not st.session_state.esperando_dudas
                and not st.session_state.esperando_duplicados
                and not st.session_state.esperando_conflictos
                and not st.session_state.get("clasificacion_activa")
                and not st.session_state.get("agregar_lote_activo")
            )
            if _en_chat_principal and st.session_state.chat:
                _ultimo_asist = next(
                    (m for m in reversed(st.session_state.chat) if m["role"] == "assistant"),
                    None,
                )
                _txt_ultimo = _ultimo_asist.get("content", "") if _ultimo_asist else ""
                _pide_confirm = isinstance(_txt_ultimo, str) and any(
                    s in _txt_ultimo.lower()
                    for s in ("proceder", "generar el excel", "generar los excel",
                              "¿puedo", "puedo proceder", "deseas que", "quieres que",
                              "procedo", "confirma", "podemos continuar")
                )
                if _pide_confirm:
                    _c1, _c2 = st.columns(2)
                    with _c1:
                        if st.button("✅ Sí, generar Excel", type="primary", width="stretch"):
                            st.session_state._quick_reply = "Sí, procede a generar el Excel."
                            st.rerun()
                    with _c2:
                        if st.button("✏️ Tengo cambios / preguntas", width="stretch"):
                            pass  # el usuario escribe en el chat

            # Consumir respuesta rápida de botón o esperar input del usuario
            _quick = st.session_state.get("_quick_reply")
            if _quick:
                st.session_state._quick_reply = None
            _chat_input = st.chat_input(placeholder)
            user_input = _quick or _chat_input

            if user_input:

                # ── Aclaración durante fase de dudas ──────────────────────────────────
                if st.session_state.esperando_dudas:
                    with st.chat_message("user"):
                        st.markdown(user_input)
                    st.session_state.chat_fase.append({"role": "user", "content": user_input})

                    with st.chat_message("assistant"):
                        with st.spinner("..."):
                            try:
                                _extra_dudas = f"DUDAS MOSTRADAS:\n{json.dumps(st.session_state.dudas_relevantes, ensure_ascii=False, indent=2)}"
                                system_aclaracion = _build_system_fase(
                                    "revisando las dudas encontradas en el packing list",
                                    st.session_state.productos,
                                    st.session_state.analisis,
                                    extra=_extra_dudas,
                                )
                                lc_aclaracion = [
                                    HumanMessage(content=m["content"]) if m["role"] == "user"
                                    else AIMessage(content=m["content"])
                                    for m in st.session_state.chat_fase
                                ]
                                llm_ac = ChatAnthropic(model="claude-sonnet-4-6", max_tokens=1000)
                                resp   = llm_ac.invoke([SystemMessage(content=system_aclaracion)] + lc_aclaracion)
                                respuesta = resp.content.strip()
                                st.markdown(respuesta)
                                st.session_state.chat_fase.append({"role": "assistant", "content": respuesta})
                            except Exception as e:
                                st.error(f"Error: {e}")
                    st.rerun()

                # ── Chat durante clasificación padre/variante ─────────────────────────
                elif st.session_state.get("clasificacion_activa"):
                    with st.chat_message("user"):
                        st.markdown(user_input)
                    st.session_state.chat_fase.append({"role": "user", "content": user_input})
                    with st.chat_message("assistant"):
                        with st.spinner("..."):
                            try:
                                system_clas = _build_system_fase(
                                    "revisando la clasificación de productos (padres, variantes, duplicados)",
                                    st.session_state.productos,
                                    st.session_state.analisis,
                                )
                                lc_clas = [
                                    HumanMessage(content=m["content"]) if m["role"] == "user"
                                    else AIMessage(content=m["content"])
                                    for m in st.session_state.chat_fase
                                    if isinstance(m.get("content"), str)
                                ]
                                llm_clas = ChatAnthropic(model="claude-sonnet-4-6", max_tokens=800)
                                resp_clas = llm_clas.invoke([SystemMessage(content=system_clas)] + lc_clas)
                                respuesta_clas = resp_clas.content.strip()
                                st.markdown(respuesta_clas)
                                st.session_state.chat_fase.append({"role": "assistant", "content": respuesta_clas})
                            except Exception as e:
                                st.error(f"Error: {e}")
                    st.rerun()

                # ── Chat durante duplicados o conflictos: aclaración contextual ────────
                elif st.session_state.esperando_duplicados or st.session_state.esperando_conflictos:
                    with st.chat_message("user"):
                        st.markdown(user_input)
                    st.session_state.chat_fase.append({"role": "user", "content": user_input})

                    with st.chat_message("assistant"):
                        with st.spinner("..."):
                            try:
                                _nombre_fase = (
                                    "revisando posibles productos duplicados o variantes"
                                    if st.session_state.esperando_duplicados
                                    else "revisando conflictos de SKU con ODOO"
                                )
                                system_fase = _build_system_fase(
                                    _nombre_fase,
                                    st.session_state.productos,
                                    st.session_state.analisis,
                                )
                                lc_fase = [
                                    HumanMessage(content=m["content"]) if m["role"] == "user"
                                    else AIMessage(content=m["content"])
                                    for m in st.session_state.chat_fase
                                    if isinstance(m.get("content"), str)
                                ]
                                llm_fase = ChatAnthropic(model="claude-sonnet-4-6", max_tokens=800)
                                resp_fase = llm_fase.invoke([SystemMessage(content=system_fase)] + lc_fase)
                                respuesta_fase = resp_fase.content.strip()
                                st.markdown(respuesta_fase)
                                st.session_state.chat_fase.append({"role": "assistant", "content": respuesta_fase})
                            except Exception as e:
                                st.error(f"Error: {e}")
                    st.rerun()

                # ── Chat principal ─────────────────────────────────────────────────────
                else:
                    if not st.session_state.chat:
                        st.info("Sube un packing list en el panel izquierdo para comenzar.")

                    with st.chat_message("user"):
                        st.markdown(user_input)
                    st.session_state.chat.append({"role": "user", "content": user_input})

                    with st.chat_message("assistant"):
                        with st.spinner("El agente está procesando..."):
                            try:
                                system = build_system_prompt(
                                    st.session_state.analisis,
                                    st.session_state.productos,
                                    tipo_cambio,
                                    contenedor,
                                    respuestas_dudas=st.session_state.respuestas_dudas or None,
                                )
                                excel_antes = st.session_state.excel_bytes
                                texto = llamar_agente(
                                    user_input,
                                    system,
                                    st.session_state.lc_messages,
                                )
                                # Guardar en historial LangChain
                                st.session_state.lc_messages.append(HumanMessage(content=user_input))
                                st.session_state.lc_messages.append(AIMessage(content=texto))
                                st.session_state.chat.append({"role": "assistant", "content": texto})

                                st.markdown(texto)
                                if st.session_state.excel_bytes != excel_antes:
                                    st.success("Excel actualizado — usa los botones de descarga de abajo.")

                            except Exception as e:
                                err = f"Error: {e}"
                                st.session_state.chat.append({"role": "assistant", "content": err})
                                st.error(err)

                    st.rerun()

    # ── Pantalla inicial ───────────────────────────────────────────────────────────
    elif not st.session_state.chat:
            st.info("Sube un packing list en el panel izquierdo para comenzar.")




# ═══════════════════════════════════════════════════════════════════════════════
# TAB 2 — AGREGAR PRODUCTOS A EXCEL FERRAFORME
# ═══════════════════════════════════════════════════════════════════════════════

def _agregar_productos_a_excel_ferraforme(excel_bytes: bytes,
                                           prods_nuevos: list[dict]) -> bytes:
    """
    Carga el Excel FERRAFORME existente desde un archivo temporal (para preservar
    las imágenes originales exactamente igual) y agrega los nuevos productos al
    final — sin tocar las filas ni imágenes existentes.

    Columnas FERRAFORME (después del delete_cols que hace generar_excel):
      1:nombre  2:sku  3:imagen  4:cbm_por_pieza  5:precio_usd  6:None
      7:piezas_x_caja  8:cbm_caja  9:tipo_producto  10:EMPRESA  11:None
      12:contenedor  13:costo_landed_usd  14:descripcion  15:categoria
      16:atributo  17-20:None
    """
    IMG_W_PX = 120
    IMG_H_PX = 120
    ROW_H_PT = 90

    # ── 1. Abrir el Excel existente ───────────────────────────────────────────
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
    ws = wb.active

    # ── 2. Encontrar la última fila con datos (scan de arriba hacia abajo) ──────
    #    max_row no es confiable cuando el Excel viene del template (reporta 2075
    #    por celdas fantasma aunque los datos reales terminen mucho antes).
    #    Scan top-down: avanzar fila a fila y parar en la primera fila donde
    #    tanto col 1 (nombre) como col 2 (sku) estén vacías.
    _last_data_row = 2  # mínimo: fila de encabezados
    for _r in range(3, ws.max_row + 2):
        if ws.cell(_r, 1).value or ws.cell(_r, 2).value:
            _last_data_row = _r
        else:
            break  # primera fila vacía → aquí terminan los datos

    # Inferir contenedor de los existentes
    _contenedor_ref = ""
    for _r in range(3, _last_data_row + 1):
        _cv = ws.cell(_r, 12).value
        if _cv:
            _contenedor_ref = str(_cv)
            break

    # ── 3. Agregar filas de los productos nuevos al final ─────────────────────
    for _i, _np in enumerate(prods_nuevos):
        _r   = _last_data_row + 1 + _i
        _sku = _np.get("sku", "")
        _nom = _np.get("titulo") or _np.get("nombre", "")

        ws.cell(_r, 1).value = _nom
        ws.cell(_r, 2).value = _sku

        # Imagen del producto nuevo
        _img_incrustada = False
        _img_d = _np.get("_img_data")
        if _img_d and _PILLOW_OK:
            try:
                _pil = PILImage.open(io.BytesIO(_img_d)).convert("RGB")
                _pil.thumbnail((IMG_W_PX, IMG_H_PX), PILImage.LANCZOS)
                _buf_i = io.BytesIO()
                _pil.save(_buf_i, format="PNG")
                _buf_i.seek(0)
                _xl_img = XLImage(_buf_i)
                _xl_img.width  = IMG_W_PX
                _xl_img.height = IMG_H_PX
                _xl_img.anchor = f"C{_r}"
                ws.add_image(_xl_img)
                ws.row_dimensions[_r].height = ROW_H_PT
                _img_incrustada = True
            except Exception:
                pass
        if not _img_incrustada:
            ws.cell(_r, 3).value = ""

        _cu = _safe_float(_np.get("cbm_por_pieza"))
        _cbm_cell = ws.cell(_r, 4)
        _cbm_cell.value = _cu if _cu > 0 else None
        if _cbm_cell.value is not None:
            _cbm_cell.number_format = "0.000000"

        ws.cell(_r, 5).value = _np.get("precio_usd")
        ws.cell(_r, 6).value = _np.get("piezas_total")    # col "Unidades"

        _px = _np.get("piezas_x_caja")
        ws.cell(_r, 7).value = _px
        _cbm_caja = ws.cell(_r, 8)
        _cbm_caja.value = round(float(_px) * _cu, 6) if _px and _cu > 0 else None
        if _cbm_caja.value is not None:
            _cbm_caja.number_format = "0.000000"

        ws.cell(_r, 9).value  = "Producto almacenable"
        ws.cell(_r, 10).value = EMPRESA
        ws.cell(_r, 11).value = None
        ws.cell(_r, 12).value = _contenedor_ref
        ws.cell(_r, 13).value = float(_np["precio_usd"]) if _np.get("precio_usd") else None
        ws.cell(_r, 14).value = _np.get("descripcion")
        ws.cell(_r, 15).value = _np.get("categoria")
        ws.cell(_r, 16).value = _np.get("atributo_desc") or _np.get("atributo")
        for _col in range(17, 21):
            ws.cell(_r, _col).value = None

    # ── 4. Guardar a BytesIO y devolver ───────────────────────────────────────
    _buf_out = io.BytesIO()
    wb.save(_buf_out)
    _buf_out.seek(0)
    return _buf_out.read(), len(prods_nuevos)


_solo_defaults = {
    "agregar_solo_paso":         1,
    "agregar_solo_imgs":         {},
    "agregar_solo_prods":        [],
    "agregar_solo_conf":         [],
    "agregar_solo_n":            1,
    "agregar_solo_excel_bytes":  None,
    "agregar_solo_excel_nombre": "",
    "agregar_solo_excel_result": None,   # bytes del Excel ya actualizado
    "agregar_solo_ultimo_error": None,   # error persistente para mostrar tras rerun
}
for _k, _v in _solo_defaults.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v

with tab_agregar:
    solo_paso = st.session_state.get("agregar_solo_paso", 1)

    # ════════ PASO 1 — Subir Excel base + imágenes ════════════════════════════
    if solo_paso == 1:
        st.subheader("Agregar productos a un Excel FERRAFORME")
        st.markdown(
            "Carga el Excel que ya tienes y las fotografías de los productos nuevos. "
            "El sistema generará el SKU, nombre, descripción y atributos a partir de "
            "la imagen, los comparará con Odoo y los agregará al Excel."
        )
        st.divider()

        # ── Excel FERRAFORME ──────────────────────────────────────────────────
        st.markdown("**1. Excel FERRAFORME al que quieres agregar productos**")
        solo_excel_subido = st.file_uploader(
            "Sube el Excel (.xlsx)",
            type=["xlsx", "xls"],
            key="solo_excel_uploader",
        )
        if solo_excel_subido:
            st.success(f"✅ **{solo_excel_subido.name}** cargado")

        st.divider()

        # ── Imágenes de los nuevos productos ─────────────────────────────────
        st.markdown("**2. Fotografías de los productos a agregar**")
        solo_n = st.number_input(
            "Cantidad de productos",
            min_value=1, max_value=30,
            value=st.session_state.get("agregar_solo_n", 1), step=1,
            key="solo_n_input",
        )
        solo_imgs_subidas = st.file_uploader(
            f"Sube hasta {int(solo_n)} imagen(es)",
            type=["png", "jpg", "jpeg", "webp"],
            accept_multiple_files=True,
            key="solo_imgs_uploader",
        )

        solo_imgs_data: list[dict] = []
        if solo_imgs_subidas:
            for _f in solo_imgs_subidas[: int(solo_n)]:
                _ext = Path(_f.name).suffix.lstrip(".").lower() or "jpg"
                solo_imgs_data.append({"data": _f.read(), "ext": _ext})

        solo_nombres: list[str] = []
        if solo_imgs_data:
            st.markdown(f"**{len(solo_imgs_data)} imagen(es) cargadas:**")
            _nc = min(len(solo_imgs_data), 4)
            _pc = st.columns(_nc)
            for _i, _img in enumerate(solo_imgs_data):
                with _pc[_i % _nc]:
                    st.image(_img["data"], width=130)
                    _nom = st.text_input(
                        f"Nombre {_i + 1} (opcional)",
                        placeholder=f"Producto {_i + 1}",
                        key=f"solo_nombre_{_i}",
                    )
                    solo_nombres.append(_nom)

        _puede_procesar = solo_imgs_data and solo_excel_subido
        if not solo_excel_subido and solo_imgs_data:
            st.warning("Sube el Excel FERRAFORME para continuar.")

        if st.button(
            "▶️ Analizar y generar SKUs",
            type="primary",
            width="stretch",
            disabled=not _puede_procesar,
            key="solo_procesar_btn",
        ):
            _imgs_dict: dict = {}
            for _i, _img in enumerate(solo_imgs_data):
                _nom = solo_nombres[_i].strip() if _i < len(solo_nombres) else ""
                _imgs_dict[_i] = {"data": _img["data"], "ext": _img["ext"], "nombre": _nom}
            st.session_state.agregar_solo_imgs         = _imgs_dict
            st.session_state.agregar_solo_n            = len(_imgs_dict)
            st.session_state.agregar_solo_excel_bytes  = solo_excel_subido.read()
            st.session_state.agregar_solo_excel_nombre = solo_excel_subido.name
            st.session_state.agregar_solo_excel_result = None
            st.session_state.agregar_solo_paso         = 2
            st.rerun()

    # ════════ PASO 2 — Vision + ODOO (automático) ════════════════════════════
    elif solo_paso == 2:
        _solo_imgs_dict    = st.session_state.get("agregar_solo_imgs", {})
        _solo_skus_odoo    = st.session_state.get("odoo_skus", [])
        _solo_prods_odoo   = st.session_state.get("odoo_productos", [])
        _solo_phashes_odoo = st.session_state.get("odoo_phashes", {})

        with st.spinner(f"Analizando {len(_solo_imgs_dict)} imagen(es) con Vision..."):
            _solo_prods_nuevos: list[dict] = []
            _solo_conflictos:   list[dict] = []

            for _i, _img_info in _solo_imgs_dict.items():
                _nombre_manual = _img_info.get("nombre", "").strip()
                # Solo pasar nombre como contexto si el usuario lo escribió;
                # si no, dejar que Vision lo genere libre desde la imagen.
                _ctx = {"fila_excel_0idx": -(_i + 1), "_manual": True}
                if _nombre_manual:
                    _ctx["nombre"] = _nombre_manual

                _datos   = analizar_imagen_claude(_img_info["data"], _img_info["ext"], contexto=_ctx)
                _sub_cod = _datos.get("subcategoria_cod", "VAR")
                _att_cod = _datos.get("atributo_cod", "GEN")
                _sku_ini = generar_sku(_sub_cod, _att_cod)

                # El nombre final es: lo que Vision generó (titulo), o lo que escribió el usuario
                _titulo_final = _datos.get("titulo") or _nombre_manual or f"Producto {_i + 1}"
                _prod: dict = {
                    "nombre":        _titulo_final,
                    "titulo":        _titulo_final,
                    "categoria":     _datos.get("categoria", ""),
                    "subcategoria":  _datos.get("subcategoria", ""),
                    "atributo_desc": _datos.get("atributo_desc", ""),
                    "atributo":      _datos.get("atributo_desc", ""),
                    "descripcion":   _datos.get("descripcion", ""),
                    "fila_excel_0idx": -(_i + 1),
                    "_manual": True,
                }

                _conf_entry = None
                if _solo_skus_odoo:
                    _val = validar_sku_vs_odoo(_sku_ini, _solo_skus_odoo)
                    _sku_final = _val["sku_ajustado"]
                    if _val["conflicto"]:
                        _conf_entry = {
                            "idx": _i, "nombre": _nombre_manual,
                            "sku_propuesto": _sku_ini, "sku_ajustado": _sku_final,
                            "datos_gemini": _datos, "productos_odoo": [],
                            "razon": "sku", "similares_odoo": [],
                        }
                else:
                    _sku_final = _sku_ini

                if _solo_prods_odoo:
                    _sims = buscar_similares_odoo(
                        _img_info["data"], _nombre_manual, _solo_prods_odoo, _solo_phashes_odoo
                    )
                else:
                    _sims = []

                if _sims and _conf_entry is None:
                    _razones = []
                    if any(s.get("por_imagen") for s in _sims): _razones.append("imagen similar")
                    if any(s.get("por_nombre") for s in _sims): _razones.append("nombre similar")
                    _conf_entry = {
                        "idx": _i, "nombre": _nombre_manual,
                        "sku_propuesto": _sku_ini, "sku_ajustado": _sku_ini,
                        "datos_gemini": _datos,
                        "razon": ", ".join(_razones) or "similitud detectada",
                        "similares_odoo": _sims,
                        "productos_odoo": [s["producto_odoo"] for s in _sims if s.get("producto_odoo")],
                    }
                elif _sims and _conf_entry is not None:
                    _conf_entry["similares_odoo"] = _sims
                    _conf_entry["productos_odoo"] = [s["producto_odoo"] for s in _sims if s.get("producto_odoo")]

                _prod["sku"]       = _sku_final
                _prod["_img_data"] = _img_info["data"]
                _prod["_img_ext"]  = _img_info["ext"]

                _solo_prods_nuevos.append(_prod)
                if _conf_entry:
                    _solo_conflictos.append(_conf_entry)

        st.session_state.agregar_solo_prods = _solo_prods_nuevos
        st.session_state.agregar_solo_conf  = _solo_conflictos
        st.session_state.agregar_solo_paso  = 3
        st.rerun()

    # ════════ PASO 3 — Revisar, resolver conflictos y descargar Excel ════════
    elif solo_paso == 3:
        _solo_prods_nuevos    = st.session_state.get("agregar_solo_prods", [])
        _solo_conflictos_lote = st.session_state.get("agregar_solo_conf", [])
        _solo_excel_nombre    = st.session_state.get("agregar_solo_excel_nombre", "archivo.xlsx")
        _solo_excel_result    = st.session_state.get("agregar_solo_excel_result")
        _solo_excel_bytes_ok  = bool(st.session_state.get("agregar_solo_excel_bytes"))

        # Mostrar error persistente si lo hubo en el intento anterior
        if st.session_state.get("agregar_solo_ultimo_error"):
            st.error(st.session_state.agregar_solo_ultimo_error)

        _n_esperados = st.session_state.get("agregar_solo_n", len(_solo_prods_nuevos))
        _n_reales    = len(_solo_prods_nuevos)

        st.subheader(f"✅ {_n_reales} producto(s) listos para agregar")
        if not _solo_excel_bytes_ok:
            st.error("⚠️ No se encontraron los bytes del Excel. Vuelve al Paso 1 y sube el archivo de nuevo.")
        else:
            st.caption(f"Se agregarán al Excel: **{_solo_excel_nombre}**")

        # Aviso si la cantidad analizada no cuadra con lo que el usuario subió
        if _n_reales != _n_esperados:
            st.warning(
                f"⚠️ Indicaste **{_n_esperados}** producto(s) pero solo se analizaron **{_n_reales}**. "
                "Verifica que todas las imágenes se subieron correctamente o vuelve al Paso 1."
            )

        # ── Vista previa de productos ─────────────────────────────────────────
        _nc2 = min(len(_solo_prods_nuevos), 4) or 1
        _pc2 = st.columns(_nc2)
        for _i, _prod in enumerate(_solo_prods_nuevos):
            with _pc2[_i % _nc2]:
                if _prod.get("_img_data"):
                    st.image(_prod["_img_data"], width=130)
                st.markdown(
                    f"**{_prod.get('titulo') or _prod.get('nombre', '—')}**  \n"
                    f"SKU: `{_prod.get('sku', '—')}`"
                )
                if _prod.get("categoria"):
                    st.caption(f"📂 {_prod['categoria']}")
                if _prod.get("atributo_desc"):
                    st.caption(f"🏷️ {_prod['atributo_desc']}")
                if _prod.get("descripcion"):
                    st.caption(f"📝 {_prod['descripcion']}")

        # ── Resolución de conflictos ODOO ─────────────────────────────────────
        _solo_resoluciones: dict = {}
        if _solo_conflictos_lote:
            st.divider()
            st.warning(
                f"⚠️ {len(_solo_conflictos_lote)} posible(s) coincidencia(s) con productos en Odoo — "
                "revisa antes de continuar."
            )
            for _conf in _solo_conflictos_lote:
                _c_idx      = _conf["idx"]
                _c_nom      = _conf["nombre"]
                _c_razon    = _conf.get("razon", "")
                _c_similares = _conf.get("similares_odoo", [])
                _c_sku_aj   = _conf.get("sku_ajustado", _conf.get("sku_propuesto", ""))

                _razon_badge = {
                    "sku":            "🔑 Mismo prefijo SKU",
                    "imagen similar": "🖼️ Imagen similar",
                    "nombre similar": "🔤 Nombre similar",
                }.get(_c_razon, f"⚠️ {_c_razon}")

                st.markdown(f"**{_c_nom}** — {_razon_badge}")
                if _c_similares:
                    _sc = st.columns(min(len(_c_similares), 3))
                    for _si, _sim in enumerate(_c_similares[:3]):
                        with _sc[_si]:
                            _img_url = st.session_state.get("odoo_imagen_urls", {}).get(_sim["sku"])
                            if _img_url:
                                st.image(_img_url, width=100)
                            st.caption(f"SKU: `{_sim['sku']}`  \n{_sim.get('nombre', '—')}")

                _dec_conf = st.radio(
                    "¿Qué hacemos?",
                    [
                        f"SKU nuevo `{_c_sku_aj}` (es un producto diferente)",
                        "Reutilizar SKU de Odoo (ya existe)",
                    ],
                    key=f"solo_conf_dec_{_c_idx}",
                )
                if "Reutilizar" in _dec_conf and _c_similares:
                    _sim_opts = [s["sku"] for s in _c_similares[:5]]
                    _sku_sel = st.selectbox(
                        "SKU de Odoo a reutilizar:",
                        _sim_opts,
                        key=f"solo_conf_sku_{_c_idx}",
                    )
                    _solo_resoluciones[_c_idx] = {"tipo": "mismo", "sku": _sku_sel}
                else:
                    _solo_resoluciones[_c_idx] = {"tipo": "diferente", "sku": _c_sku_aj}
                st.divider()

        # ── Botones ───────────────────────────────────────────────────────────
        _col_gen, _col_back2 = st.columns(2)

        with _col_gen:
            if st.button("📥 Generar Excel actualizado", type="primary",
                         width="stretch", key="solo_generar_btn",
                         disabled=not _solo_excel_bytes_ok):
                # Aplicar resoluciones de SKU
                for _conf in _solo_conflictos_lote:
                    _c_idx = _conf["idx"]
                    _res   = _solo_resoluciones.get(
                        _c_idx, {"tipo": "diferente", "sku": _conf.get("sku_ajustado", "")}
                    )
                    if _res["tipo"] == "mismo":
                        _solo_prods_nuevos[_c_idx]["sku"] = _res["sku"]

                st.session_state.agregar_solo_ultimo_error = None
                with st.spinner("Agregando productos al Excel..."):
                    try:
                        _result_bytes, _n_agregados = _agregar_productos_a_excel_ferraforme(
                            st.session_state.agregar_solo_excel_bytes,
                            _solo_prods_nuevos,
                        )
                        # Avisar si se agregaron menos filas de las esperadas
                        if _n_agregados != _n_esperados:
                            st.session_state.agregar_solo_ultimo_error = (
                                f"⚠️ Se agregaron {_n_agregados} fila(s) pero se esperaban "
                                f"{_n_esperados}. Verifica el resultado antes de usar el Excel."
                            )
                        st.session_state.agregar_solo_excel_result  = _result_bytes
                        st.session_state.agregar_solo_prods         = _solo_prods_nuevos
                        st.session_state.agregar_solo_n_agregados   = _n_agregados
                        st.rerun()
                    except Exception as _ex:
                        import traceback
                        _tb = traceback.format_exc()
                        st.session_state.agregar_solo_ultimo_error = (
                            f"Error al actualizar el Excel: {_ex}\n\nDetalle:\n{_tb}"
                        )
                        st.rerun()

        with _col_back2:
            if st.button("↩️ Volver y cambiar imágenes", width="stretch", key="solo_back_btn"):
                st.session_state.agregar_solo_paso  = 1
                st.session_state.agregar_solo_prods = []
                st.session_state.agregar_solo_conf  = []
                st.session_state.agregar_solo_excel_result = None
                st.rerun()

        # ── Botón de descarga (aparece tras generar) ──────────────────────────
        if _solo_excel_result:
            _base_nom, _ = (lambda p: (p.rsplit(".", 1)[0], p.rsplit(".", 1)[1]) if "." in p else (p, ""))(_solo_excel_nombre)
            _nombre_out = _base_nom + "_actualizado.xlsx"
            _n_agr = st.session_state.get("agregar_solo_n_agregados", len(_solo_prods_nuevos))
            if _n_agr == _n_esperados:
                st.success(f"✅ Excel listo — se agregaron **{_n_agr}** fila(s) correctamente.")
            else:
                st.warning(f"⚠️ Excel generado, pero se agregaron **{_n_agr}** de **{_n_esperados}** fila(s) esperadas.")
            st.download_button(
                label="⬇️ Descargar Excel actualizado",
                data=_solo_excel_result,
                file_name=_nombre_out,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="solo_download_btn",
                use_container_width=True,
            )
            if st.button("➕ Agregar más productos", width="stretch", key="solo_reset_btn"):
                # Mantener el Excel resultado como base para la siguiente ronda
                st.session_state.agregar_solo_excel_bytes  = _solo_excel_result
                st.session_state.agregar_solo_paso         = 1
                st.session_state.agregar_solo_imgs         = {}
                st.session_state.agregar_solo_prods        = []
                st.session_state.agregar_solo_conf         = []
                st.session_state.agregar_solo_excel_result = None
                st.rerun()

# ── Historial de conversación (visible desde ambas tabs) ──────────────────────
with tab_agregar:
    if st.session_state.chat:
        with st.expander(
            f"🕓 Historial del agente ({len(st.session_state.chat)} mensajes)",
            expanded=False,
        ):
            for _hmsg in st.session_state.chat:
                with st.chat_message(_hmsg["role"]):
                    _hcontent = _hmsg["content"]
                    if isinstance(_hcontent, str):
                        st.markdown(_hcontent)
                    else:
                        for _hbloque in _hcontent:
                            if not isinstance(_hbloque, dict):
                                st.markdown(str(_hbloque))
                            elif _hbloque.get("type") == "text":
                                st.markdown(_hbloque["value"])
                            elif _hbloque.get("type") == "image_bytes":
                                try:
                                    st.image(_hbloque["data"], width=_hbloque.get("width", 160),
                                             caption=_hbloque.get("caption", ""))
                                except Exception:
                                    pass
