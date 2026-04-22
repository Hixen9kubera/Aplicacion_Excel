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

from typing import TypedDict

from langgraph.graph import StateGraph, START, END

from agents.vision_agent import agente_vision
from agents.nombres_agent import _extraer_nombre_base_atributo_batch


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
