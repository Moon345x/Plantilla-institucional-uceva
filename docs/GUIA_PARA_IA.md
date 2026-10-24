# Guía para IA: crear y convertir presentaciones con la plantilla UCEVA 55 años

Esta guía es la referencia completa. Si tienes poco contexto, lee primero `AGENTS.md`.
Todas las medidas salen de `spec/tokens.json` y `spec/slides.json`, que se generan desde el `.pptx`
con `scripts/extraer_plantilla.py`. No hace falta abrir ni descomprimir la plantilla.

Unidades: **EMU**. 914400 EMU = 1 pulgada; 12700 EMU = 1 punto. La hoja mide 12192000 × 6858000 EMU
(13,333" × 7,5", 16:9).

---

## 1. Qué tipo de diapositiva usar

| Si tienes… | Usa | Función |
|---|---|---|
| El inicio de la presentación (título, curso, autor) | Portada (muestra 2) | `portada` |
| Un saludo de bienvenida a un evento o clase, sin texto propio | Bienvenida (1) | `bienvenida` |
| Un tema con texto, lista, tabla, código o imagen | Contenido con etiqueta (6; 5 rosada; 4 amarilla) | `diapositiva_contenido` + elementos |
| El comienzo de una sección o un mensaje que quieres destacar | Barra lateral (7 morada, 8 rosada) | `diapositiva_barra_lateral` |
| Una foto protagonista con un pie corto | Foto con pie (9) | `diapositiva_foto` |
| 2 a 8 ideas cortas con un nombre cada una (objetivos, pasos, características) | Lista de bloques (10) | `diapositiva_lista_de_bloques` |
| Contenido libre sin título (un diagrama, una imagen grande) | En blanco (3) | `diapositiva_en_blanco` |
| El final | Cierre (11) | `cierre` |
| Una lista corta (≤ 6 viñetas) | Contenido | `texto(..., vinetas=True)` |
| 2-4 conceptos paralelos (agenda, comparación) | Contenido | `fila_de_tarjetas` |
| 4-6 ideas para repasar | Contenido | `cuadricula` |
| Datos en filas y columnas | Contenido | `tabla` (+ `dividir_tabla` si es larga) |
| Código fuente | Contenido | `bloque_codigo` |
| 1-4 cifras clave | Contenido | `fila_de_numeros` / `numero_grande` |
| Texto + imagen | Contenido | `area.columnas(2)` + `texto` + `imagen` |

Colores de la etiqueta: morado por defecto; rosado y amarillo para variar o para distinguir secciones.
Mantén el mismo color dentro de una sección.

## 2. Flujo para crear una presentación nueva

1. **Plan**. Escribe la lista de diapositivas: título, tipo (tabla anterior), contenido y notas.
   Una idea por diapositiva. Portada al principio y cierre al final.
2. **Mapear a funciones**. Cada diapositiva = una función de tipo + uno o más elementos.
3. **Generar** (ver `ejemplos/ejemplo_crear.py`):

   ```python
   import sys; sys.path.insert(0, "lib")
   from uceva_pptx import *
   prs = abrir_plantilla()
   portada(prs, "Título", "Subtítulo", autor="Autor", notas="…")
   s, area = diapositiva_contenido(prs, "Facultad de …", "Título de la diapositiva", notas="…")
   texto(s, area, ["Idea 1", "Idea 2"], vinetas=True)
   cierre(prs)
   guardar(prs, "salida/mi_presentacion.pptx")
   ```

   Si una función lanza `ContenidoDemasiadoLargo`, divide ese contenido en dos diapositivas.
4. **Validar**: `python scripts/validar.py salida/mi_presentacion.pptx`. Corrige todos los `ERROR`.
   Revisa cada `AVISO` y deja solo los que tengan una razón.
5. **Render**: `bash scripts/render_previews.sh salida/mi_presentacion.pptx salida/previews`.
6. **Revisar las imágenes**: abre cada PNG. Busca texto cortado, cosas pisadas, huecos grandes,
   contraste pobre, alineaciones raras. El validador estima; la imagen es la prueba.
7. **Corregir** y repetir 4-6.

> Las vistas previas se hacen con LibreOffice u OnlyOffice. Si la máquina no tiene Arial usan
> Liberation Sans (mismas medidas, se ve igual). Si no tiene **Arial Black**, el título de la portada
> se verá con una letra más delgada: es solo la vista previa; en PowerPoint se ve en Arial Black.

## 3. Flujo para convertir un `.pptx` de otro estilo

Punto de partida: `python ejemplos/ejemplo_convertir.py ORIGEN.pptx --etiqueta "Facultad de …"`.
Ese script aplica una heurística simple; tú decides lo que la heurística no sabe.

1. **Leer el original**. Opciones:
   - `markitdown ORIGEN.pptx` → texto en Markdown, rápido para entender el contenido.
   - `python-pptx` → título, párrafos con nivel de viñeta, tablas, imágenes (`shape.image.blob`) y
     notas (`slide.notes_slide.notes_text_frame.text`). Ver la función `leer` de `ejemplo_convertir.py`.
2. **Decidir, diapositiva por diapositiva**, a qué tipo corresponde (tabla de la sección 1).
   No copies el diseño viejo: si el original tenía 3 columnas de texto, piensa si son tarjetas;
   si tenía una cifra enorme, `numero_grande`; si era una portada de sección, barra lateral.
3. **Reconstruir** con la librería. **Nunca copies formas del archivo viejo** (traen fuentes,
   colores y posiciones de otra plantilla). Copia solo el texto, los datos y las imágenes.
4. **Conservar** los textos (sin resumir salvo que el usuario lo pida) y las **notas del orador**
   (`notas=`). Si una diapositiva original no cabe, repártela en varias y pon "(cont.)" en el título.
5. **Recomponer imágenes**: extráelas a un archivo y colócalas con `imagen(...)` dentro de un `Area`
   (se encajan sin deformarse). Las fotos protagonistas, con `diapositiva_foto`. Las imágenes con
   texto pequeño (capturas, diagramas) necesitan espacio: dales media diapositiva o una completa.
6. Ojo con el tamaño: los originales 4:3 (10" × 7,5") se rehacen en 16:9; no estires nada.
7. Valida, renderiza y revisa igual que al crear.

## 4. Cómo adaptar contenido denso

- **Más de 6 viñetas** → divide en dos diapositivas o agrupa en tarjetas con título.
- **Párrafos largos** → quédate con la idea; lo demás va a las notas del orador.
- **Listas de "concepto: explicación"** → `fila_de_tarjetas` (2-4) o `cuadricula` (4-6) o `diapositiva_lista_de_bloques`.
- **Comparaciones** → tabla de 2-3 columnas o tarjetas lado a lado.
- **Cifras sueltas en el texto** → sácalas a `fila_de_numeros`.
- **Tablas de más de ~8 filas** → `dividir_tabla(filas, 6)` y una diapositiva por trozo.
- La letra no baja de 14 pt en el cuerpo. Si no cabe a 14, sobra contenido.

## 5. Casos especiales

- **Código**: `bloque_codigo(s, area, codigo, lenguaje="Python")`. Máximo ~14 líneas de ~70
  caracteres. Muestra solo el fragmento importante; el resto, en un repositorio o en las notas.
  Las líneas no se parten (sería engañoso en código).
- **Tablas**: encabezado corto; números alineados a la derecha o al centro con `alineaciones=[…]`;
  `anchos=[1, 3, 2]` para columnas proporcionales.
- **Ecuaciones**: python-pptx no crea ecuaciones de Office. Opciones: escribirla en texto plano si es
  simple (`a² + b² = c²`, usa caracteres Unicode), o generar una imagen PNG (p. ej. con matplotlib
  `mathtext` o LaTeX) con fondo transparente y ponerla con `imagen(...)`.
- **Gráficos**: genera el gráfico como PNG (matplotlib u otra) usando los colores de `tokens.json`
  (morado 6000AF primero, luego rosado FF0066; amarillo y cian solo como relleno, nunca para texto),
  fuente Arial y fondo transparente o blanco; ponlo con `imagen(...)`. Si necesitas un gráfico
  editable, usa `slide.shapes.add_chart(...)` de python-pptx dentro del `Area` y aplica esos colores.
- **Imágenes**: siempre dentro de un `Area`; `borde=True` para fotos (como la muestra 9); sin borde
  para capturas y diagramas. No pongas texto encima de fotos salvo en el cierre (que ya trae la capa).

## 6. Checklist de calidad final

- [ ] `validar.py` sin errores; avisos revisados.
- [ ] Revisé cada PNG: nada cortado, nada pisado, nada fuera de la hoja.
- [ ] Contraste: texto oscuro sobre claro y blanco sobre morado/rosado; nada de texto amarillo o cian sobre blanco.
- [ ] Alineación: todo el contenido empieza en el mismo margen izquierdo que el título.
- [ ] Tamaños: cuerpo ≥ 14 pt, pies ≥ 12 pt, títulos 24 pt.
- [ ] Sin restos de la plantilla ("TÍTULO DE EJEMPLO", "Subtítulo de ejemplo", "Este es un texto corto de ejemplo", "Facultad de Ciencias Jurídicas y Humanísticas") ni placeholders vacíos.
- [ ] Una idea por diapositiva; sin líneas decorativas bajo títulos.
- [ ] Ortografía y tildes revisadas (títulos incluidos); mayúsculas solo al inicio y en nombres propios.
- [ ] Orden: portada primero, cierre último, secciones en orden lógico.
- [ ] Notas del orador presentes donde había notas en el original o donde el usuario las pidió.
- [ ] Datos inventados o del usuario; nada personal de terceros.

## 7. Si no tienes Python

Usa los datos de `spec/tokens.json` y `spec/slides.json` como fuente de verdad y las imágenes de
`assets/img/`. Resumen de lo esencial (en pulgadas; EMU entre paréntesis):

**Hoja**: 13,333" × 7,5" (16:9). En pptxgenjs: `pptx.layout = "LAYOUT_WIDE"`.

**Fondos** (imagen a hoja completa):

| Tipo | Imagen |
|---|---|
| Contenido, en blanco, foto, bloques | `assets/img/fondo_contenido.jpg` (ya trae el logo pequeño y las figuras grises) |
| Portada | `assets/img/fondo_portada.png` (logo grande + "Bienvenidos" pequeño abajo a la izquierda) |
| Bienvenida | `assets/img/fondo_bienvenidos.png` |
| Barra lateral | `assets/img/fondo_barra_lateral.jpg` (sin logo) + rectángulo de color + `logo_55_uceva_blanco.png` |
| Cierre | `assets/img/fondo_cierre.png` (foto aérea + banda morada con logo, redes y web) |

**Diapositiva de contenido**:

| Elemento | x | y | ancho | alto | Estilo |
|---|---|---|---|---|---|
| Etiqueta de facultad | 7,805 (7136946) | 0,306 (279399) | 5,286 (4833258) | 0,419 (382685) | Rectángulo redondeado relleno 6000AF (o FF0066 / FFC403), sin línea. Texto Arial 16 negrita centrado, blanco con sombra (gris oscuro 3B3838 sobre amarillo). Su borde derecho está en 13,091". |
| Título | 0,51 (466607) | 1,205 (1101988) | 12,31 (11258786) | 0,4 por línea | Arial 24 negrita 6000AF, margen interno 0. Máx. 2 líneas. |
| Subtítulo (opcional) | 0,51 | debajo del título + 0,04" | 12,31 | 0,3 por línea | Arial 18 3B3838. |
| Área de contenido | 0,51 | 1,81 sin subtítulo · 2,15 con subtítulo | 12,31 | hasta y = 7,1 (6492240) | Cuerpo Arial 16 (mín. 14) 3B3838. |

Zona protegida del logo: x 0–2,53", y 0–0,86" (`zonas_protegidas.logo_contenido`).

**Portada**: caja del título 5,786, 1,622, 7,235 × 0,837 (rectángulo redondeado 6000AF, texto Arial Black
32 blanco centrado con sombra; si no cabe en una línea, baja a 24-30 pt y haz la caja más alta);
marco 6,184, 1,276, 7,347 × 1,949 (rectángulo redondeado sin relleno, línea 00E8FC 2,25 pt discontinua;
se sale 0,2" por la derecha a propósito; dibújalo antes que la caja); subtítulo en x 6,906, justo
debajo de la caja + 0,1", ancho hasta x = 12,82", Arial 24 negrita 6000AF. Si el título crece, baja el
subtítulo y agranda el marco: **nunca dejes el subtítulo en su posición fija**.

**Barra lateral**: rectángulo 0, 0, 2,31 × 7,5 (6000AF o FF0066); logo blanco 0,155, 0,307, 1,994 × 0,473;
título en x 2,49, y 0,357, Arial 24 negrita del color de la barra.

**Foto con pie**: foto máx. 6,974 × 4,649 centrada desde y 1,425, borde blanco 6 pt, sombra exterior
(desenfoque 4 pt, distancia 3 pt, 45°, negro 40 %); etiqueta 5,286 × 0,419 centrada, que empieza 0,15" por encima
del borde inferior de la foto (sobresale 0,27" por debajo), 6000AF, Arial 18 negrita blanco; texto debajo Arial 14 3B3838 centrado.

**Bloques**: etiquetas 3,945 × 0,419 en x 0,66 desde y 1,392 (6000AF, FF0066, FFC403, 00E8FC);
texto Arial 14 3B3838 debajo de cada etiqueta (+0,07").

**Cierre**: capa rectángulo redondeado 7" de ancho, centrada en x (3,167), centrada en vertical por
encima de la banda (la banda empieza en y = 5,243"), relleno 6000AF con **85 % de opacidad**; texto
blanco Arial 44 negrita + 20 regular.

### Advertencias por herramienta

- **pptxgenjs**: colores **sin `#`** y **de 6 dígitos** (`"6000AF"`); un hex de 8 dígitos rompe el
  archivo. La transparencia va aparte: `fill: { color: "6000AF", transparency: 15 }`. Fondo:
  `slide.background = { path: "assets/img/fondo_contenido.jpg" }`. Rectángulo redondeado:
  `pptx.ShapeType.roundRect` con `rectRadius`. Declara `fontFace: "Arial"` en cada texto. Crea un
  objeto de opciones nuevo para cada forma (algunas versiones modifican el objeto que reciben).
  Las posiciones van en pulgadas.
- **XML a mano**: posiciones en EMU (`<a:off x= y=/>`, `<a:ext cx= cy=/>`); colores
  `<a:srgbClr val="6000AF"/>`; transparencia `<a:srgbClr val="6000AF"><a:alpha val="85000"/></a:srgbClr>`;
  fuente en cada run: `<a:latin typeface="Arial"/>` (si falta, hereda Calibri del tema); tamaño
  `sz="1600"` = 16 pt. Si borras una diapositiva, quita su `<p:sldId>` **y** su relación en
  `ppt/_rels/presentation.xml.rels` y su parte, o quedará huérfana. La sombra del texto de las
  etiquetas está copiada literal en `tokens.json → elementos.etiqueta_facultad.sombra_texto_xml`.
- **Google Slides / Canva**: importa el `.pptx` generado; no reconstruyas a mano. Si no hay otra opción,
  usa las medidas de esta sección y las imágenes de `assets/img/`.
- **Cualquier herramienta**: no uses la fuente del tema (Calibri); las muestras no tienen
  placeholders con estilo propio, todo son cuadros de texto y formas.
