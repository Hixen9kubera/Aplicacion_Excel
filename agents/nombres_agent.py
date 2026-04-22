"""
Agente Nombres — Fase C de la arquitectura multi-agente FERRAFORME.

Responsabilidades:
- Normalizar nombres de productos (traducción chino → español)
- Detectar duplicados y variantes entre productos del mismo packing list
- Aplicar resoluciones del usuario a grupos de duplicados
- Extraer nombre_base y atributo en batch (para clasificación de variantes)

Corre en paralelo con agente_vision (no depende de imágenes).

Entrada : productos (list[dict])
Salida  : {productos (normalizados), clasificaciones (nombre_base/atributo por producto)}
"""
from __future__ import annotations

import difflib
import json
import re

from langchain_anthropic import ChatAnthropic
from langchain_core.messages import HumanMessage

from agents.utils import IMAGENES_TEMP_PATH, _phash_imagen, _similitud_nombres


# ── Detección de chino ────────────────────────────────────────────────────────

def _tiene_chino(texto: str) -> bool:
    """Devuelve True si el texto contiene caracteres chinos."""
    return bool(re.search(r'[\u4e00-\u9fff\u3400-\u4dbf]', str(texto)))


# ── Normalización de nombres ──────────────────────────────────────────────────

def normalizar_nombres_productos(productos: list[dict]) -> list[dict]:
    """
    Para cada producto, elige el nombre más descriptivo entre nombre y nombre_alt
    y lo traduce al español si está en chino. Una sola llamada batch a Claude.
    """
    for prod in productos:
        if "nombre_chino_orig" not in prod:
            _n_raw = str(prod.get("nombre")     or "").strip()
            _n_alt = str(prod.get("nombre_alt") or "").strip()
            if _tiene_chino(_n_raw):
                prod["nombre_chino_orig"] = _n_raw
            elif _tiene_chino(_n_alt):
                prod["nombre_chino_orig"] = _n_alt
            else:
                prod["nombre_chino_orig"] = _n_raw

    local_a_real: dict[int, int] = {}
    items_para_claude: list[dict] = []

    for real_i, prod in enumerate(productos):
        nombre     = str(prod.get("nombre")     or "").strip()
        nombre_alt = str(prod.get("nombre_alt") or "").strip()
        if _tiene_chino(nombre) or _tiene_chino(nombre_alt):
            local_idx = len(items_para_claude)
            local_a_real[local_idx] = real_i
            items_para_claude.append({
                "idx":        local_idx,
                "nombre":     nombre,
                "nombre_alt": nombre_alt or None,
            })

    if not items_para_claude:
        return productos

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
            if local_idx is not None and nombre_final and local_idx in local_a_real:
                productos[local_a_real[local_idx]]["nombre"] = nombre_final
    except Exception:
        pass

    return productos


# ── Campos para detección de duplicados ──────────────────────────────────────

_CAMPOS_DETERMINANTES = [
    "precio_usd", "cbm_por_pieza", "cbm_master_carton", "piezas_x_caja", "cbm_total_sku",
]
_CAMPOS_DIFF_DISPLAY = [
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
    tiene_datos = any(
        a.get(c) not in (None, "", 0) and b.get(c) not in (None, "", 0)
        for c in _CAMPOS_DETERMINANTES
    )
    if not tiene_datos:
        return False
    return all(str(a.get(c, "")).strip() == str(b.get(c, "")).strip() for c in _CAMPOS_DETERMINANTES)


def _nombre_norm(p: dict) -> str:
    return str(p.get("nombre", "") or "").strip().lower()


def _sim_chino(a: dict, b: dict) -> float:
    """Similitud [0–1] entre los nombres chinos originales de dos productos."""
    ca = str(a.get("nombre_chino_orig") or a.get("nombre", "")).strip()
    cb = str(b.get("nombre_chino_orig") or b.get("nombre", "")).strip()
    if not ca or not cb:
        return 0.0
    return difflib.SequenceMatcher(None, ca, cb).ratio()


# ── Clasificación de palabras de variante en 3 niveles ───────────────────────

_PALABRAS_VERDE: set[str] = {
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
    "red","blue","yellow","black","white","gray","grey",
    "orange","purple","pink","brown","gold","silver",
    "xs","sm","md","lg","xl","xxl","xxxl",
}

_PALABRAS_AMARILLO: set[str] = {
    "mini","maxi","micro",
    "pequeño","pequeña","chico","chica",
    "mediano","mediana","grande","grandes",
    "largo","larga","corto","corta",
    "angosto","angosta","compacto","compacta",
    "madera","mdf","bambu","bambú","roble","pino","nogal","haya",
    "metal","metalico","metálico","metalica","metálica",
    "acero","inoxidable","aluminio","hierro","fierro",
    "plastico","plástico","plastica","plástica","pvc",
    "tela","lona","cuero","piel","nylon","poliester","poliéster","oxford",
    "vidrio","cristal","ceramica","cerámica","porcelana",
    "silicona","goma","hule","latex","látex",
    "pack","set","kit","pieza","piezas","unidad","unidades",
    "simple","doble","triple",
}

_PALABRAS_ROJO: set[str] = {
    "electrico","eléctrico","electrica","eléctrica",
    "manual","automatico","automático","automatica","automática",
    "inalambrico","inalámbrico","inalambrica","inalámbrica",
    "recargable","a","pilas",
    "neumatico","neumático","hidraulico","hidráulico",
    "pro","lite","plus","max","ultra","prime","elite","sport","deluxe",
    "plegable","portatil","portátil","fijo","fija","desmontable",
}

_PALABRAS_VARIANTE: set[str] = _PALABRAS_VERDE | _PALABRAS_AMARILLO


def _es_tecnico(token: str) -> bool:
    """True si el token es especificación técnica: 21v, 500w, 2l, 12kg, etc."""
    return bool(re.match(r'^\d+(\.\d+)?[a-z]+$', token))


def _nombre_base_verde(nom: str) -> str:
    """Quita solo atributos cosméticos (colores, tallas). Base verde coincide = variante segura."""
    tokens = []
    for t in nom.split():
        if _es_tecnico(t):
            tokens.append(t)
        elif t in _PALABRAS_VERDE:
            pass
        elif len(t) > 2:
            tokens.append(t)
    return " ".join(tokens)


def _nombre_base_amarillo(nom: str) -> str:
    """Quita cosméticos + materiales/presentación. Base amarilla coincide = posible variante."""
    tokens = []
    for t in nom.split():
        if _es_tecnico(t):
            tokens.append(t)
        elif t in _PALABRAS_VERDE or t in _PALABRAS_AMARILLO:
            pass
        elif len(t) > 2:
            tokens.append(t)
    return " ".join(tokens)


def _tiene_diferencia_roja(nom_a: str, nom_b: str) -> bool:
    """True si los nombres difieren en palabras ROJAS (tipo/specs) → productos distintos."""
    tok_a = set(nom_a.split())
    tok_b = set(nom_b.split())
    if (tok_a & _PALABRAS_ROJO).symmetric_difference(tok_b & _PALABRAS_ROJO):
        return True
    specs_a = {t for t in tok_a if _es_tecnico(t)}
    specs_b = {t for t in tok_b if _es_tecnico(t)}
    if specs_a and specs_b and specs_a != specs_b:
        return True
    return False


def _nombres_son_similares(nom_a: str, nom_b: str, ratio_min: float = 0.78) -> bool:
    """True si los nombres normalizados son suficientemente similares."""
    if not nom_a or not nom_b or nom_a == nom_b:
        return False
    if difflib.SequenceMatcher(None, nom_a, nom_b).ratio() >= ratio_min:
        return True
    tok_a = set(nom_a.split())
    tok_b = set(nom_b.split())
    if not tok_a or not tok_b:
        return False
    shorter = tok_a if len(tok_a) <= len(tok_b) else tok_b
    longer  = tok_b if len(tok_a) <= len(tok_b) else tok_a
    shorter_sig = {t for t in shorter if len(t) > 2}
    if not shorter_sig:
        return False
    return len(shorter_sig & longer) / len(shorter_sig) >= 0.75


def _diffs_display(prods: list[dict]) -> dict[str, list[str]]:
    """Devuelve {campo: [val_prod0, val_prod1, ...]} solo para campos que varían."""
    result = {}
    for campo in _CAMPOS_DIFF_DISPLAY:
        vals = [str(p.get(campo, "") or "—").strip() for p in prods]
        if len(set(vals)) > 1:
            result[campo] = vals
    return result


# ── Detección de duplicados ───────────────────────────────────────────────────

def detectar_productos_duplicados(
    productos: list[dict],
    imagenes: dict | None = None,
) -> list[dict]:
    """
    Detecta grupos de productos duplicados o variantes en el packing list.

    FASE 0 — Nombre chino original (señal más fuerte):
      A0. Nombre chino idéntico (≥98%) → "exacto"
      A1. Nombre chino ≥80% + phash imagen ≤ 6 → "exacto"
      A2. Nombre chino ≥82% (sin imagen) → "probable"
      A3. Nombre chino ≥62% → "similar"

    FASE 1 — Mismo producto por nombre español/determinantes:
      B. Imagen idéntica + mismo nombre → "exacto"
      C. Imagen idéntica + nombre distinto → "variante_imagen"
      D. Mismo nombre español → "probable"

    FASE 2 — Posibles variantes:
      D. Base verde igual + sin dif. roja → "nombre_similar"
      E. Base amarilla igual + imagen similar → "similar"
      F. Determinantes iguales + sin dif. roja → "similar"
      G. Imagen similar + nombre parcial + sin dif. roja → "variante_imagen"
      H. Nombre fuzzy similar + sin dif. roja → "nombre_similar"
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

    phashes: dict[int, object] = {}
    if imagenes:
        for idx, prod in enumerate(productos):
            fila = prod.get("fila_excel_0idx")
            img  = imagenes.get(fila + 1) if fila is not None else None
            if img:
                ph = _phash_imagen(img["data"])
                if ph is not None:
                    phashes[idx] = ph

    # ── FASE 0 — NOMBRE CHINO ─────────────────────────────────────────────────
    _CHINO_EXACTO   = 0.98
    _CHINO_PROBABLE = 0.82
    _CHINO_VARIANTE = 0.62
    _hay_chino = any(_tiene_chino(str(p.get("nombre_chino_orig", ""))) for p in productos)

    if _hay_chino:
        for i in range(len(productos)):
            if i in usados or not _tiene_chino(str(productos[i].get("nombre_chino_orig", ""))):
                continue
            grp = [i]
            for j in range(i + 1, len(productos)):
                if j not in usados and _sim_chino(productos[i], productos[j]) >= _CHINO_EXACTO:
                    grp.append(j)
            if len(grp) > 1:
                for k in grp: usados.add(k)
                _agregar(grp, "exacto")

        for i in range(len(productos)):
            if i in usados or i not in phashes or not _tiene_chino(str(productos[i].get("nombre_chino_orig", ""))):
                continue
            grp = [i]
            for j in range(i + 1, len(productos)):
                if j in usados or j not in phashes:
                    continue
                if _sim_chino(productos[i], productos[j]) >= 0.80 and phashes[i] - phashes[j] <= 6:
                    grp.append(j)
            if len(grp) > 1:
                for k in grp: usados.add(k)
                _agregar(grp, "exacto")

        for i in range(len(productos)):
            if i in usados or not _tiene_chino(str(productos[i].get("nombre_chino_orig", ""))):
                continue
            grp = [i]
            for j in range(i + 1, len(productos)):
                if j not in usados:
                    sim = _sim_chino(productos[i], productos[j])
                    if _CHINO_PROBABLE <= sim < _CHINO_EXACTO:
                        grp.append(j)
            if len(grp) > 1:
                for k in grp: usados.add(k)
                _agregar(grp, "probable")

        for i in range(len(productos)):
            if i in usados or not _tiene_chino(str(productos[i].get("nombre_chino_orig", ""))):
                continue
            grp = [i]
            for j in range(i + 1, len(productos)):
                if j not in usados:
                    sim = _sim_chino(productos[i], productos[j])
                    if _CHINO_VARIANTE <= sim < _CHINO_PROBABLE:
                        grp.append(j)
            if len(grp) > 1:
                for k in grp: usados.add(k)
                _agregar(grp, "similar")

    # ── FASE 1 — MISMO PRODUCTO ───────────────────────────────────────────────
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

    for i in range(len(productos)):
        if i in usados or i not in phashes:
            continue
        grp = [i]
        for j in range(i + 1, len(productos)):
            if j in usados or j not in phashes:
                continue
            if any(phashes[j] - phashes[k] <= 2 for k in grp if k in phashes):
                grp.append(j)
        if len(grp) > 1:
            for k in grp: usados.add(k)
            _agregar(grp, "variante_imagen")

    for i in range(len(productos)):
        if i in usados:
            continue
        nom_i = _nombre_norm(productos[i])
        grp = [i]
        for j in range(i + 1, len(productos)):
            if j not in usados and _nombre_norm(productos[j]) == nom_i:
                grp.append(j)
        if len(grp) > 1:
            for k in grp: usados.add(k)
            _agregar(grp, "probable")

    # ── FASE 2 — POSIBLES VARIANTES ───────────────────────────────────────────
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
            nom_j = _nombre_norm(productos[j])
            if (_nombre_base_verde(nom_j) == base_vi
                    and nom_i != nom_j
                    and not _tiene_diferencia_roja(nom_i, nom_j)):
                grp.append(j)
        if len(grp) > 1:
            for k in grp: usados.add(k)
            _agregar(grp, "nombre_similar")

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
            nom_j = _nombre_norm(productos[j])
            img_ok = any(phashes[j] - phashes[k] <= 10 for k in grp if k in phashes)
            if (_nombre_base_amarillo(nom_j) == base_ai
                    and nom_i != nom_j
                    and not _tiene_diferencia_roja(nom_i, nom_j)
                    and img_ok):
                grp.append(j)
        if len(grp) > 1:
            for k in grp: usados.add(k)
            _agregar(grp, "similar")

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
            for k in grp: usados.add(k)
            _agregar(grp, "similar")

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

    # ── Post-procesamiento: fusionar grupos relacionados ──────────────────────
    _PRIO = {"exacto": 4, "probable": 3, "similar": 3, "variante_imagen": 2, "nombre_similar": 1}

    def _grupos_relacionados(ga: dict, gb: dict) -> bool:
        _noms_a  = [_nombre_norm(p) for p in ga["productos"]]
        _noms_b  = [_nombre_norm(p) for p in gb["productos"]]
        _bases_va = [_nombre_base_verde(n)    for n in _noms_a if len(_nombre_base_verde(n).split())    >= 2]
        _bases_vb = [_nombre_base_verde(n)    for n in _noms_b if len(_nombre_base_verde(n).split())    >= 2]
        _bases_aa = [_nombre_base_amarillo(n) for n in _noms_a if len(_nombre_base_amarillo(n).split()) >= 2]
        _bases_ab = [_nombre_base_amarillo(n) for n in _noms_b if len(_nombre_base_amarillo(n).split()) >= 2]
        if any(_nombres_son_similares(na, nb, _FUZZY_RATIO) for na in _noms_a for nb in _noms_b):
            return True
        if any(bva == bvb for bva in _bases_va for bvb in _bases_vb):
            return True
        if any(baa == bab for baa in _bases_aa for bab in _bases_ab):
            return True
        _chin_a = [p for p in ga["productos"] if _tiene_chino(str(p.get("nombre_chino_orig", "")))]
        _chin_b = [p for p in gb["productos"] if _tiene_chino(str(p.get("nombre_chino_orig", "")))]
        if _chin_a and _chin_b and any(_sim_chino(pa, pb) >= _CHINO_VARIANTE for pa in _chin_a for pb in _chin_b):
            return True
        _ph_a = {idx: phashes[idx] for idx in ga["indices"] if idx in phashes}
        _ph_b = {idx: phashes[idx] for idx in gb["indices"] if idx in phashes}
        if _ph_a and _ph_b and any(pha - phb <= 10 for pha in _ph_a.values() for phb in _ph_b.values()):
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


# ── Aplicar resoluciones del usuario ─────────────────────────────────────────

def aplicar_resolucion_duplicados(
    productos: list[dict],
    grupos: list[dict],
    respuestas: dict,
) -> list[dict]:
    """
    respuestas: {str(gid): "diferente" | dict}
    - "diferente"                             → todos independientes
    - {"tipo": "mismo",    "sel": [ci,...]}   → fusiona los seleccionados
    - {"tipo": "variantes","sel": [ci,...]}   → comparten base de SKU
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
            pass

        elif isinstance(decision, dict):
            tipo = decision.get("tipo", "variantes")
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

    return [p for i, p in enumerate(productos) if i not in indices_a_eliminar]


# ── Extracción de nombre_base y atributo en batch ────────────────────────────

def _extraer_nombre_base_atributo_batch(productos: list[dict]) -> list[dict]:
    """
    Una sola llamada a Claude Haiku para extraer nombre_base, atributo_tipo y
    atributo_valor de todos los productos. Devuelve lista en el mismo orden.
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


# ── Nodo LangGraph ─────────────────────────────────────────────────────────────

def agente_nombres(productos: list[dict]) -> dict:
    """
    Nodo LangGraph — pipeline de normalización y extracción de atributos.
    Corre en paralelo con agente_vision (no depende de imágenes).

    Pasos:
    1. normalizar_nombres_productos — traduce nombres chinos al español
    2. _extraer_nombre_base_atributo_batch — extrae nombre_base y atributo

    No incluye detectar_productos_duplicados ni aplicar_resolucion_duplicados
    ya que son pasos interactivos (requieren respuesta del usuario en el UI).

    Devuelve:
      {
        "productos":        list[dict],  # con nombres normalizados
        "clasificaciones":  list[dict],  # [{nombre_base, atributo_tipo, atributo_valor}]
      }
    """
    productos = normalizar_nombres_productos(productos)
    clasificaciones = _extraer_nombre_base_atributo_batch(productos)
    return {
        "productos":       productos,
        "clasificaciones": clasificaciones,
    }
