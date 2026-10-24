# `lib/uceva_pptx.py`

Librería pequeña sobre `python-pptx` para crear diapositivas con la plantilla UCEVA 55 años
sin tocar XML. Lee las medidas y colores de `spec/tokens.json`.

## Importar

```python
import sys
sys.path.insert(0, "lib")          # ruta a esta carpeta, relativa a la raíz del repo
from uceva_pptx import *
```

## Ejemplo de 10 líneas

```python
from uceva_pptx import *

prs = abrir_plantilla()
portada(prs, "Bases de datos", "Sesión 1 · Curso de ejemplo", autor="Equipo docente")
s, area = diapositiva_contenido(prs, "Facultad de Ingeniería", "¿Qué es SQL?", "Un lenguaje para consultar datos")
texto(s, area, ["Se escribe como frases en inglés", "Lo entienden casi todos los motores"], vinetas=True)
s, area = diapositiva_contenido(prs, "Facultad de Ingeniería", "Tipos de datos", color="rosado")
tabla(s, area, [["Tipo", "Ejemplo"], ["INTEGER", "42"], ["VARCHAR", "'Ana'"]])
cierre(prs, "¡Gracias!", "Preguntas")
guardar(prs, "salida/bases_de_datos.pptx")
```

## Cómo se trabaja con el espacio: `Area`

Las funciones de diapositiva devuelven `(diapositiva, area)`. `area` es el espacio libre para el
contenido (en EMU: 914400 por pulgada) y ya respeta márgenes, logo, etiqueta y título.
Las funciones de contenido **devuelven el `Area` que ocuparon**, para seguir debajo:

```python
from uceva_pptx import *

prs = abrir_plantilla()
s, area = diapositiva_contenido(prs, "Facultad de Ingeniería", "Resumen")
usada = texto(s, area, "Una frase de introducción.")
resto = area.debajo(usada)                 # lo que queda debajo, con separación estándar (0,2")
fila_de_tarjetas(s, resto, [("Uno", "Primera idea."), ("Dos", "Segunda idea.")])

izq, der = area.columnas(2)                # dos columnas iguales
a, b = area.columnas(2, proporciones=[2, 1])
arriba, abajo = area.filas(2)
area.pulgadas()                            # (x, y, ancho, alto) en pulgadas, para depurar
```

Si algo no cabe, la función lanza **`ContenidoDemasiadoLargo`** con un mensaje que dice qué hacer.
La respuesta correcta es dividir el contenido en más diapositivas, no achicar la letra.

## Abrir y guardar

| Función | Qué hace |
|---|---|
| `abrir_plantilla(ruta=None)` | Abre `plantilla/…VF.pptx` (no lo modifica). Recuerda las 11 muestras, las imágenes y los diseños. |
| `guardar(prs, ruta, conservar_muestras=False)` | Borra las 11 muestras (soltando su relación con `drop_rel`, sin partes huérfanas), quita placeholders vacíos, pone los cierres al final, renumera y guarda. Crea la carpeta si no existe. |

## Tipos de diapositiva

Todas aceptan `notas="…"` (notas del orador).

| Función | Muestra | Devuelve |
|---|---|---|
| `bienvenida(prs)` | 1 | diapositiva |
| `portada(prs, titulo, subtitulo=None, autor=None, fecha=None)` | 2 | diapositiva |
| `diapositiva_contenido(prs, etiqueta, titulo, subtitulo=None, color="morado")` | 4, 5, 6 | `(diapositiva, area)` |
| `diapositiva_barra_lateral(prs, titulo, subtitulo=None, color="morado")` | 7, 8 | `(diapositiva, area)` |
| `diapositiva_en_blanco(prs)` | 3 | `(diapositiva, area)` |
| `diapositiva_foto(prs, ruta_imagen, pie, texto=None)` | 9 | diapositiva |
| `diapositiva_lista_de_bloques(prs, bloques, titulo=None, imagen_derecha=None)` | 10 | diapositiva |
| `cierre(prs, texto_principal="¡Gracias!", texto_secundario=None)` | 11 | diapositiva |
| `copiar_muestra(prs, numero)` + `reemplazar_texto(s, viejo, nuevo)` | cualquiera | diapositiva |

Detalles:

- **`portada`**: título en la caja morada en Arial Black 32 pt. Si es largo baja hasta 24 pt y la caja
  crece hacia abajo (hasta 3 líneas). El subtítulo (24 pt negrita morado, baja a 18 si no cabe en
  una línea), el autor y la fecha (16 pt gris) se apilan **debajo** de la caja; el marco punteado
  cian crece para envolverlos. Título máximo recomendado: ~60 caracteres.
- **`diapositiva_contenido`**: `etiqueta` es el texto de la etiqueta de facultad (o `None` para no
  ponerla). `color`: `"morado"` (muestra 6), `"rosado"` (5), `"amarillo"` (4). El título va en morado
  (en rosado si `color="rosado"`). Título máx. 2 líneas, subtítulo máx. 2 líneas.
- **`diapositiva_barra_lateral`**: `color` `"morado"` o `"rosado"`. El área empieza a la derecha de la barra.
- **`diapositiva_foto`**: la foto se encaja sin deformarse (máx. 6,97" × 4,65"), con borde blanco de 6 pt y
  sombra. `pie` va en la etiqueta morada (una línea); `texto` es una frase debajo (14 pt).
- **`diapositiva_lista_de_bloques`**: lista de `(etiqueta, texto)`. 1-4 bloques en una columna
  (con `imagen_derecha` opcional en la mitad derecha); 5-8 en dos columnas. Colores en orden
  morado, rosado, amarillo, cian. Etiqueta de bloque máx. ~28 caracteres.
- **`cierre`**: el texto va sobre una capa morada al 85 % de opacidad (transparencia con `<a:alpha>`),
  centrada sobre la foto y lejos de la banda morada. Queda al final al guardar.
- **`copiar_muestra`**: copia exacta (fondo, formas, imágenes) de la muestra `numero` (1-11). Úsala solo
  si necesitas la forma original; luego cambia **todos** los textos de ejemplo con `reemplazar_texto`
  (el validador marca error si queda alguno).

## Elementos de contenido

Todos reciben la diapositiva `s` y un `area` (o x, y, ancho) y devuelven el `Area` ocupada.

| Función | Para qué |
|---|---|
| `texto(s, area, contenido, tam=16, color="gris_texto", negrita=False, alineacion="izquierda", vinetas=False)` | Párrafo o lista. `contenido` puede ser `str` o lista; con `vinetas=True` cada elemento es una viñeta y `("texto", 1)` es una subviñeta. `**así**` va en negrita. Si no cabe baja hasta 14 pt; si aun así no cabe, error. |
| `tarjeta(s, x, y, ancho, titulo, cuerpo=None, color="morado", alto=None, tam_cuerpo=16)` | Tarjeta blanca con borde gris claro, cabecera de color y texto. Alto ajustado al contenido. |
| `fila_de_tarjetas(s, area, [(titulo, cuerpo), …], colores=None)` | 2-4 tarjetas iguales en fila (agenda, comparaciones). |
| `cuadricula(s, area, [(titulo, cuerpo), …], columnas=2)` | Tarjetas en cuadrícula (repasos, 4-6 ideas). |
| `tabla(s, area, filas, anchos=None, tam=16, alineaciones=None)` | `filas[0]` es el encabezado (morado, texto blanco); filas alternas morado suave / blanco. |
| `dividir_tabla(filas, max_filas)` | Parte una tabla larga repitiendo el encabezado. |
| `bloque_codigo(s, area, codigo, tam=14, lenguaje=None)` | Código en Courier New, texto blanco sobre `morado_oscuro`. No parte líneas; baja hasta 12 pt; si no cabe, error. |
| `imagen(s, ruta, area, alinear="centro", borde=False)` | Imagen encajada sin deformar. `borde=True`: borde blanco y sombra como la muestra 9. |
| `numero_grande(s, x, y, ancho, numero, descripcion, color="morado")` | Cifra de 54 pt con descripción de 14 pt. |
| `fila_de_numeros(s, area, [(numero, descripcion), …])` | 2-4 cifras en fila. |

Utilidades: `medir_alto_texto(parrafos, pt, ancho_emu)`, `color_texto_sobre(fondo)` (devuelve
`"blanco"` o `"gris_texto"` según el contraste), `COLORES` (nombre → hex), `TOKENS` (todo `tokens.json`).

## Colores por nombre

`morado` 6000AF · `rosado` FF0066 · `amarillo` FFC403 · `cian` 00E8FC · `gris_texto` 3B3838 ·
`gris_claro` E7E6E6 · `blanco` FFFFFF · `negro` 000000 (solo sombras) ·
derivados: `morado_oscuro` 1D0035 (fondo de código) · `morado_suave` F2EBF9 (filas alternas).

## Qué resuelve por ti

- Borra los placeholders vacíos que `python-pptx` copia del diseño al crear cada diapositiva.
- Al borrar diapositivas suelta la relación (`drop_rel`) para no dejar partes huérfanas.
- Declara Arial en cada texto (si no, PowerPoint usaría Calibri, la fuente del tema).
- Reutiliza las imágenes de la plantilla (logo, fondos) sin duplicarlas en el archivo.
- Elige texto blanco o gris oscuro sobre cada color para que haya contraste (en amarillo y cian, gris oscuro).
- Mide el texto (con Arial o Liberation Sans, que tiene las mismas medidas) para ajustar altos y evitar desbordes.
