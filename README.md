# Plantilla institucional UCEVA 55 años — lista para IA

Este repo contiene la plantilla de presentaciones de la UCEVA (55 años) **desmenuzada y documentada**,
para que cualquier persona o IA (Claude, ChatGPT, Gemini, Copilot…) pueda:

1. **Crear** una presentación nueva con la plantilla a partir de contenido (títulos, textos, tablas, código, imágenes).
2. **Convertir** una presentación existente de otro estilo a esta plantilla, diapositiva por diapositiva.

…sin descomprimir el `.pptx` ni adivinar posiciones, colores o fuentes.

![Portada de la plantilla](assets/previews/slide02.png)

## En 1 minuto

- **Si eres una IA**: lee [`AGENTS.md`](AGENTS.md).
- **Si eres una persona**: instala, ejecuta el ejemplo y mira el resultado en `salida/`.

```bash
pip install -r requirements.txt
python ejemplos/ejemplo_crear.py          # crea salida/ejemplo.pptx
python scripts/validar.py salida/ejemplo.pptx
```

## Ejemplo mínimo

```python
import sys; sys.path.insert(0, "lib")
from uceva_pptx import *

prs = abrir_plantilla()
portada(prs, "Bases de datos", "Sesión 1 · Curso de ejemplo", autor="Equipo docente")
s, area = diapositiva_contenido(prs, "Facultad de Ingeniería", "¿Qué es SQL?")
texto(s, area, ["Un lenguaje para consultar datos", "Lo entienden casi todos los motores"], vinetas=True)
cierre(prs, "¡Gracias!")
guardar(prs, "salida/bases_de_datos.pptx")
```

## Qué hay en el repo

| Carpeta / archivo | Contenido |
|---|---|
| `plantilla/` | El `.pptx` original. **No se modifica nunca.** |
| `spec/tokens.json` | Colores, fuentes, tamaños, márgenes, zonas protegidas y medidas de cada elemento. |
| `spec/layouts.json` | Los 11 diseños (layouts) del archivo y para qué sirve cada uno. |
| `spec/slides.json` | Las 11 diapositivas de muestra: cada forma con posición, relleno, línea, texto, fuente y sombra. |
| `assets/img/` | Imágenes de la plantilla con nombres claros (fondos, logo blanco, logo a color, foto). |
| `assets/previews/` | Un PNG por diapositiva de muestra. |
| `lib/uceva_pptx.py` | Librería sobre python-pptx con funciones listas ([documentación](lib/README.md)). |
| `ejemplos/` | Crear una presentación, convertir otra, ver todos los tipos, contenido de ejemplo. |
| `scripts/` | Extraer la plantilla, validar un `.pptx`, renderizar vistas previas. |
| `docs/` | [Guía para IA](docs/GUIA_PARA_IA.md), [catálogo](docs/CATALOGO_DE_DIAPOSITIVAS.md), [reglas de diseño](docs/REGLAS_DE_DISENO.md), [errores comunes](docs/ERRORES_COMUNES.md). |

## Convertir una presentación existente

```bash
python ejemplos/ejemplo_convertir.py mi_presentacion.pptx --etiqueta "Facultad de Ingeniería"
# → salida/mi_presentacion_uceva.pptx
```

Es una heurística (título → contenido, tablas → tabla, imágenes → imagen, notas → notas). Revisa el
resultado; la guía explica cómo decidir cada diapositiva.

## Vistas previas

```bash
bash scripts/render_previews.sh salida/ejemplo.pptx salida/previews
```

Necesita **LibreOffice** (`soffice`) u **OnlyOffice Desktop Editors** (usa su conversor `x2t`), y
`pdftoppm` (paquete `poppler-utils`). Sin Arial Black instalada, el título de la portada se ve con
otra letra en la vista previa (en PowerPoint se ve bien).

## Regenerar todo desde la plantilla

Solo hace falta si cambia el `.pptx`:

```bash
python scripts/extraer_plantilla.py     # regenera spec/, assets/img/ y assets/previews/
```

El script lee el archivo sin modificarlo, vuelve a leer cada forma con python-pptx para comprobar que
posiciones y colores coinciden, y es determinista (dos ejecuciones dan los mismos archivos).

## Contribuir

- Abre un issue o pull request con un ejemplo concreto (qué contenido, qué salió mal, captura).
- Si cambias la librería: ejecuta `ejemplos/ejemplo_crear.py`, `ejemplos/ejemplo_todos_los_tipos.py`
  y la conversión de prueba (`ejemplos/generar_pptx_de_prueba.py` + `ejemplo_convertir.py`), valida
  los tres y mira las vistas previas.
- Si encuentras un error nuevo, agrégalo a `docs/ERRORES_COMUNES.md`.
- Nada de datos personales ni de estudiantes en ejemplos: usa datos inventados.
- Documentación en español sencillo. Commits pequeños con mensajes claros en español.
