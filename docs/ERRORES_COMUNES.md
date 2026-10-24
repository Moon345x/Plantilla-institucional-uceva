# Errores comunes y cómo se resuelven

Problemas que ya aparecieron al usar la plantilla (a mano o con IA). Para cada uno: qué pasa, por qué
y cómo lo resuelve la librería o qué hacer si trabajas sin ella. El validador detecta los marcados con ✔.

## 1. El título largo de la portada se pisa con el subtítulo ✔

- **Qué pasa**: el título ocupa dos líneas, la caja morada crece (o el texto se sale) y tapa el subtítulo,
  que en la muestra está en una posición fija justo debajo.
- **Por qué**: Arial Black 32 pt es muy ancha: en la caja de 7,2" caben unos 25 caracteres por línea.
- **Solución**: `portada()` baja la letra (32 → 24 pt), agranda la caja hacia abajo y coloca el subtítulo
  **después** de la caja; el marco cian crece para envolver todo. Sin la librería: mide el título,
  agranda la caja y mueve el subtítulo. Mejor aún: título corto (≤ 60 caracteres) y el resto al subtítulo.

## 2. El subtítulo en dos líneas choca con el texto de abajo ✔

- **Qué pasa**: el subtítulo de la portada (24 pt negrita) en la caja original de 3,9" de ancho parte en
  dos líneas y se monta sobre el autor o la fecha. En diapositivas de contenido, un título de dos líneas
  tapa el primer párrafo.
- **Solución**: la librería ensancha el subtítulo de la portada hasta el margen derecho, baja a 18 pt si
  hace falta y apila cada texto según su alto real. En contenido, `diapositiva_contenido` mide título y
  subtítulo y devuelve el `Area` libre debajo: pon el contenido **en ese `Area`**, nunca en una y fija.

## 3. Placeholders vacíos sueltos ✔

- **Qué pasa**: aparece "Haga clic para agregar título" o una caja vacía encima del contenido.
- **Por qué**: `python-pptx` copia los marcadores del diseño al crear una diapositiva
  (`add_slide(layout)`), y el diseño "Diapositiva de título" trae título y subtítulo.
- **Solución**: la librería los borra al crear cada diapositiva y otra vez al guardar. Sin la librería:
  `for ph in list(s.placeholders): ph._element.getparent().remove(ph._element)`.

## 4. Texto blanco ilegible sobre la foto del cierre

- **Qué pasa**: "¡Gracias!" en blanco sobre la foto aérea (árboles, techos claros) no se lee.
- **Solución**: `cierre()` pone el texto sobre una capa morada con 85 % de opacidad. La transparencia va en
  XML con `<a:alpha val="85000"/>` dentro del color; **no** con un hex de 8 dígitos (`6000AFD9`), que
  PowerPoint no entiende y en pptxgenjs corrompe el archivo.

## 5. Tarjetas con mucho espacio vacío abajo

- **Qué pasa**: tarjetas altas (p. ej. todo el alto disponible) con dos líneas de texto arriba y un
  hueco grande debajo.
- **Solución**: `tarjeta`, `fila_de_tarjetas` y `cuadricula` calculan el alto **ajustado al contenido**;
  en una fila, todas toman el alto de la más alta. Si sobra espacio en la diapositiva, está bien:
  no estires las tarjetas. Si quieres llenar, sube el tamaño de letra del cuerpo (`tam_cuerpo=18`).

## 6. Partes huérfanas al borrar diapositivas ✔

- **Qué pasa**: el archivo pesa casi 10 MB aunque tenga 3 diapositivas, o PowerPoint pide "reparar".
- **Por qué**: se quitó la diapositiva de la lista (`<p:sldId>`) pero no su relación; la parte (y sus
  fotos) siguen dentro del archivo.
- **Solución**: `guardar()` hace `lista.remove(sldId)` **y** `prs.part.drop_rel(rId)`, y renumera las partes.

## 7. Las 11 diapositivas de muestra quedan en el archivo final ✔

- **Solución**: guarda con `guardar(prs, ruta)` (no con `prs.save`). El validador avisa si una diapositiva
  es idéntica a una muestra y da error si siguen todas.

## 8. Textos de ejemplo que quedan ✔

- "TÍTULO DE EJEMPLO", "Subtítulo de ejemplo", "Este es un texto corto de ejemplo",
  "Este es un texto de ejemplo", "Facultad de Ciencias Jurídicas y Humanísticas".
- Pasa sobre todo al usar `copiar_muestra`. Cambia todos con `reemplazar_texto` o, mejor, usa las funciones
  que crean la diapositiva desde cero.

## 9. El texto sale en Calibri ✔ (aviso)

- **Por qué**: la fuente del tema del archivo es Calibri. Un texto sin `<a:latin typeface="Arial"/>` la hereda.
- **Solución**: la librería declara la fuente en cada run. Sin la librería, decláralo siempre.

## 10. Amarillo o cian como color de texto ✔ (aviso)

- La muestra 4 usa el subtítulo en amarillo y las etiquetas usan blanco sobre amarillo o cian: se leen mal.
- La librería usa amarillo y cian solo como relleno, con texto gris oscuro encima; `texto()` rechaza
  `color="amarillo"` o `"cian"`.

## 11. Una línea vacía en el código se ve más alta

- **Por qué**: un párrafo vacío toma el tamaño de su marca de fin de párrafo, que por defecto es 18 pt.
- **Solución**: la librería pone en cada párrafo `endParaRPr` con el mismo tamaño y fuente que el texto.

## 12. La vista previa no se parece a PowerPoint en el título de la portada

- **Por qué**: si la máquina no tiene Arial Black, LibreOffice/OnlyOffice usan otra fuente más delgada.
  Arial sí se ve igual porque se sustituye por Liberation Sans, que tiene exactamente las mismas medidas.
- **Qué hacer**: nada en el archivo. La librería mide el título como si fuera Arial Black (Arial negrita
  ensanchada un 15 %), así que en PowerPoint cabe. Instalar las fuentes de Microsoft en la máquina de
  render hace las vistas previas exactas.

## 13. Tablas largas que se salen por abajo ✔

- **Por qué**: las filas de una tabla crecen solas con el texto; la tabla termina más abajo de lo planeado.
- **Solución**: `tabla()` mide cada fila antes de crearla y lanza `ContenidoDemasiadoLargo` si no cabe.
  Usa `dividir_tabla(filas, 6)` y una diapositiva por trozo (título con "(cont.)").

## 14. Imágenes deformadas

- **Solución**: `imagen()` y `diapositiva_foto()` encajan la imagen en el área conservando su proporción.
  Nunca fijes ancho **y** alto a mano sin calcular la proporción.

## 15. Coordenadas inventadas

- **Qué pasa**: la IA pone el contenido en y = 1" (encima del título) o en x = 0,2" (sobre el logo).
- **Solución**: usa siempre el `Area` que devuelve la función de la diapositiva y sus métodos
  (`debajo`, `columnas`, `filas`). Sin librería, usa las medidas de la sección 7 de la guía.
