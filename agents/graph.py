"""
Graph — Fase F de la arquitectura multi-agente FERRAFORME.

StateGraph LangGraph que paraleliza Vision y Normalización de nombres,
los dos pasos más costosos del análisis del packing list.

Topología:
    START ──→ nodo_vision   ──→ END
          └─→ nodo_nombres  ──→ END

LangGraph 1.1 ejecuta nodos con múltiples aristas desde el mismo origen
de forma concurrente en un thread pool → Vision (I/O pesado con API) y
Nombres (un solo call a Claude Haiku batch) corren en paralelo.

Uso:
    from agents.graph import ejecutar_vision_y_nombres

    datos_vision, clasificaciones = ejecutar_vision_y_nombres(
        productos, imagenes, forzar_batch=True
    )
"""
from __future__ import annotations

import concurrent.futures
from typing import TypedDict

from langgraph.graph import StateGraph, START, END

from agents.vision_agent import agente_vision
from agents.nombres_agent import _extraer_nombre_base_atributo_batch, _inferir_categorias_batch


# ── Estado del grafo ──────────────────────────────────────────────────────────

class ClasificacionState(TypedDict):
    productos:       list[dict]
    imagenes:        dict[int, dict]
    forzar_batch:    bool
    on_progress:     object          # callable | None
    datos_vision:    dict[int, dict]
    clasificaciones: list[dict]


# ── Nodos ─────────────────────────────────────────────────────────────────────

def _nodo_vision(state: ClasificacionState) -> dict:
    datos = agente_vision(
        state["productos"],
        state["imagenes"],
        forzar_batch=state.get("forzar_batch", False),
        on_progress=state.get("on_progress"),
    )
    return {"datos_vision": datos}


def _nodo_nombres(state: ClasificacionState) -> dict:
    clas = _extraer_nombre_base_atributo_batch(state["productos"])
    return {"clasificaciones": clas}


# ── Construcción del grafo (se compila una vez al importar el módulo) ─────────

def _build_graph() -> "CompiledStateGraph":
    g = StateGraph(ClasificacionState)

    g.add_node("vision",  _nodo_vision)
    g.add_node("nombres", _nodo_nombres)

    # Fan-out: ambos nodos arrancan desde START de forma concurrente
    g.add_edge(START, "vision")
    g.add_edge(START, "nombres")

    # Fan-in: cada nodo escribe a campos distintos → sin conflicto de reducer
    g.add_edge("vision",  END)
    g.add_edge("nombres", END)

    return g.compile()


_grafo = _build_graph()


# ── API pública ───────────────────────────────────────────────────────────────

def ejecutar_vision_y_nombres(
    productos:    list[dict],
    imagenes:     dict[int, dict],
    forzar_batch: bool = False,
    on_progress:  "callable | None" = None,
) -> tuple[dict[int, dict], list[dict]]:
    """
    Ejecuta Vision y Normalización de nombres en paralelo (LangGraph fan-out).

    • Vision   → datos_vision    {idx: {subcategoria_cod, atributo_cod, titulo, …}}
    • Nombres  → clasificaciones [{nombre_base, atributo_tipo, atributo_valor}]

    Devuelve (datos_vision, clasificaciones).
    """
    initial: ClasificacionState = {
        "productos":       productos,
        "imagenes":        imagenes,
        "forzar_batch":    forzar_batch,
        "on_progress":     on_progress,
        "datos_vision":    {},
        "clasificaciones": [],
    }
    result = _grafo.invoke(initial)
    return result["datos_vision"], result["clasificaciones"]


def ejecutar_solo_nombres(
    productos: list[dict],
) -> tuple[dict[int, dict], list[dict]]:
    """
    Fase 1 (solo texto): extrae nombres y categorías sin Vision API.
    Corre _extraer_nombre_base_atributo_batch + _inferir_categorias_batch en paralelo.

    Devuelve (datos_vision_fase1, clasificaciones) donde datos_vision_fase1 tiene
    subcategoria_cod/atributo_cod inferidos por texto y _vision_pendiente=True.
    """
    with concurrent.futures.ThreadPoolExecutor(max_workers=2) as ex:
        fut_nombres    = ex.submit(_extraer_nombre_base_atributo_batch, productos)
        fut_categorias = ex.submit(_inferir_categorias_batch, productos)
        clasificaciones  = fut_nombres.result()
        datos_categorias = fut_categorias.result()

    datos_vision: dict[int, dict] = {}
    for i in range(len(productos)):
        cat = datos_categorias[i] if i < len(datos_categorias) else {}
        datos_vision[i] = {
            "subcategoria_cod":  cat.get("subcategoria_cod") or "VAR",
            "atributo_cod":      cat.get("atributo_cod")     or "EST",
            "titulo":            "",
            "descripcion":       "",
            "categoria":         "",
            "atributo_desc":     "",
            "_vision_pendiente": True,
            "_error":            None,
        }

    return datos_vision, clasificaciones


def ejecutar_vision_bloque_desde_disco(
    productos:  list[dict],
    offset:     int = 0,
    block_size: int = 50,
) -> dict[int, dict]:
    """
    Fase 2: corre Vision sobre productos[offset:offset+block_size].
    Lee las imágenes desde imagen_temp_path (guardadas en disco por guardar_imagenes_temp).
    Devuelve {global_idx: datos_vision} — las claves están en el espacio global (offset incluido).
    """
    from pathlib import Path as _Path

    block_prods: list[dict] = productos[offset: offset + block_size]

    imagenes_block: dict[int, dict] = {}
    for prod in block_prods:
        img_path = prod.get("imagen_temp_path")
        row_num  = prod.get("fila_excel_0idx")
        if img_path and row_num is not None:
            p = _Path(img_path)
            if p.exists():
                imagenes_block[row_num] = {"data": p.read_bytes(), "ext": p.suffix.lstrip(".")}

    datos_bloque = agente_vision(block_prods, imagenes_block)
    return {offset + local_i: datos for local_i, datos in datos_bloque.items()}
