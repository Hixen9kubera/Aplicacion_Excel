# FERRAFORME — Agente de Productos

Agente conversacional para procesar packing lists de importación y generar Excels de productos para Odoo. Construido con Streamlit + Claude (Anthropic) + LangChain.

---

## Qué hace

### Tab 1 — Packing List
1. **Sube un packing list** (Excel `.xlsx`) con los productos de un contenedor.
2. El agente analiza encabezados automáticamente (soporta formatos variados, cabeceras en chino/inglés/español).
3. Si hay dudas sobre columnas ambiguas, las pregunta antes de continuar.
4. Traduce nombres en chino al español usando Claude Vision.
5. Detecta **productos duplicados o variantes** (misma base, diferentes atributos) y pide confirmación sobre cómo tratarlos.
6. Detecta **conflictos de SKU** contra Odoo y permite resolverlos.
7. Genera un **Excel FERRAFORME** listo para importar a Odoo.
8. Permite agregar los productos nuevos directamente al chat del agente.

### Tab 2 — Agregar Productos
Flujo independiente para agregar productos sueltos (con foto) a un Excel FERRAFORME existente:
1. Sube el Excel de Ferraforme ya existente.
2. Sube 1 a N imágenes de productos nuevos.
3. Claude Vision analiza cada imagen y extrae: nombre, descripción, categoría, precio estimado, CBM.
4. Valida y ajusta SKUs contra Odoo (detecta conflictos, sugiere reutilizar SKU existente).
5. Descarga el Excel actualizado con los productos nuevos al final.

---

## Tecnologías

| Componente | Uso |
|---|---|
| [Streamlit](https://streamlit.io) | Interfaz web |
| [Claude / Anthropic](https://anthropic.com) | Análisis de imágenes (Vision), traducción, agente conversacional |
| [LangChain](https://langchain.com) | Orquestación del agente y herramientas |
| [LangSmith](https://smith.langchain.com) | Trazabilidad de llamadas al agente (opcional) |
| [Odoo XML-RPC](https://www.odoo.com/documentation/17.0/developer/reference/external_api.html) | Validación y búsqueda de SKUs existentes |
| [ChromaDB](https://trychroma.com) + sentence-transformers | Búsqueda semántica de productos Odoo |
| [openpyxl](https://openpyxl.readthedocs.io) | Lectura y escritura de archivos Excel |
| [Pillow](https://python-pillow.org) + [ImageHash](https://github.com/JohannesBuchner/imagehash) | Comparación perceptual de imágenes |

---

## Instalación

**Requisitos:** Python 3.12+

```bash
pip install -r requirements.txt
```

---

## Configuración

Crea un archivo `.env` en la raíz del proyecto con las siguientes variables:

```env
# Anthropic (Claude)
ANTHROPIC_API_KEY=

# LangSmith — trazabilidad (opcional)
LANGCHAIN_API_KEY=
LANGCHAIN_PROJECT=

# Odoo
ODOO_URL=https://tu-instancia.odoo.com
ODOO_DB=nombre_db
ODOO_USER=usuario@empresa.com
ODOO_PASSWORD=contraseña

# Google Drive (opcional, para guardar imágenes)
GOOGLE_DRIVE_CREDENTIALS=ruta/al/credentials.json
GOOGLE_DRIVE_FOLDER_ID=id_carpeta_drive
```

---

## Archivos necesarios

| Archivo | Descripción |
|---|---|
| `app_v3.py` | Aplicación principal |
| `FORMULA FERRAFORME PRODUCTOS .xlsx` | Template base para generar los Excels de salida — **no borrar** |
| `.env` | Variables de entorno y credenciales |

---

## Ejecución

```bash
streamlit run app_v3.py
```

---

## Estructura de SKUs

Los SKUs siguen el formato `SUBCAT-####-ATRIBUTO`, por ejemplo:

- `MUE-0001-AZL` — Mueble Hogar, número 0001, color Azul
- `JUG-0042-RJO` — Juguete, número 0042, color Rojo
- `BEB-0010` — Artículo Bebé sin atributo específico

El agente valida que el número no esté ya ocupado en Odoo y ajusta automáticamente en caso de conflicto.

---

## Notas

- El cache de Odoo (`odoo_cache.pkl`) y la base vectorial (`chroma_db/`) se generan automáticamente en la primera carga y se reutilizan para evitar llamadas repetidas a Odoo.
- Las imágenes temporales procesadas se guardan en `imagenes_temp/` y se limpian automáticamente.
- Para resetear el cache de Odoo, eliminar `odoo_cache.pkl` y `chroma_db/`.
