# Reglas de diseño

Qué sí y qué no hacer con la marca UCEVA 55 años. La librería ya aplica casi todo; estas reglas sirven
para revisar el resultado y para quien trabaje sin la librería.

## Logo

- El logo pequeño ya viene en el fondo de las diapositivas de contenido, y el grande en la portada,
  la bienvenida y el cierre. **No agregues otro logo** encima.
- En las diapositivas con barra lateral va el logo **blanco** (`assets/img/logo_55_uceva_blanco.png`),
  siempre sobre morado o rosado, en la posición de la muestra 7.
- No estires, recortes, recolores ni rotes el logo. Si necesitas el logo a color suelto, usa
  `assets/img/logo_55_uceva_color_recorte.png` (es un recorte del fondo de la portada) manteniendo la proporción.

## Zonas protegidas

Medidas en `spec/tokens.json → zonas_protegidas` (el validador las revisa):

| Zona | Dónde | Regla |
|---|---|---|
| Logo (contenido) | x 0–2,53", y 0–0,86" | Nada encima. |
| Logo (portada y bienvenida) | x 0–3,49", y 0–1,12" | Nada encima. |
| Arte "Bienvenidos" de la portada | x 0,64–4,89", y 4,3–6,83" | Nada encima. |
| Etiqueta de facultad | arriba a la derecha, 0,306" a 0,725" de alto | Nada encima ni pegado. |
| Barra lateral | x 0–2,31" | Solo la barra y el logo blanco. |
| Banda del cierre | y 5,243"–6,72" | Nada encima. |

Márgenes: al menos **0,4"** a los lados y abajo. El contenido se alinea con el título en x = 0,51".

## Paleta y cuándo usar cada color

| Nombre | Hex | Para |
|---|---|---|
| morado | 6000AF | Color principal: títulos, etiquetas, caja de la portada, barra, encabezados de tabla. |
| rosado | FF0066 | Acento: segunda etiqueta, barra alternativa, títulos de una sección, cifras. |
| amarillo | FFC403 | Solo relleno (etiquetas, bloques, cabeceras de tarjeta) con texto gris oscuro encima. |
| cian | 00E8FC | Marco punteado de la portada y relleno de bloques; en código, el nombre del lenguaje sobre fondo oscuro. |
| gris_texto | 3B3838 | Todo el texto corriente. |
| gris_claro | E7E6E6 | Bordes finos de tarjetas y líneas de tabla. |
| blanco | FFFFFF | Texto sobre morado/rosado; borde de fotos. |
| negro | 000000 | Solo sombras con transparencia. |
| morado_oscuro *(derivado)* | 1D0035 | Fondo de bloques de código. |
| morado_suave *(derivado)* | F2EBF9 | Filas alternas de tablas. |

Contraste: texto blanco solo sobre morado, rosado o morado oscuro; texto gris oscuro sobre blanco,
amarillo, cian, gris claro y morado suave. **Nunca** texto amarillo o cian sobre blanco.

## Tipografía

| Uso | Fuente | Tamaño | Peso | Color |
|---|---|---|---|---|
| Título de portada | Arial Black | 32 pt (baja hasta 24 si es largo) | normal | blanco con sombra |
| Subtítulo de portada | Arial | 24 pt (hasta 18) | negrita | morado |
| Título de diapositiva | Arial | 24 pt | negrita | morado (o rosado) |
| Subtítulo de diapositiva | Arial | 18 pt | normal | gris |
| Cuerpo | Arial | 16 pt (mín. 14) | normal | gris |
| Etiqueta de facultad | Arial | 16 pt | negrita | blanco (gris sobre amarillo) |
| Etiqueta de bloque / pie de foto | Arial | 18 pt | negrita | blanco |
| Texto corto de bloques y pies | Arial | 14 pt | normal | gris |
| Código | Courier New | 14 pt (mín. 12) | normal | blanco sobre morado oscuro |
| Cifras grandes | Arial | 54 pt | negrita | morado o rosado |
| Mínimo absoluto (pies, etiquetas) | Arial | 12 pt | | |

Arial Black solo en el título de la portada. Siempre declara la fuente: si no, el texto hereda Calibri
(la fuente del tema del archivo), que no es de la marca.

## Qué no hacer

- Estirar o deformar el logo, las fotos o los fondos.
- Cambiar los colores de marca o inventar colores nuevos (usa los de la tabla).
- Usar fuentes ajenas (Calibri, Times, Comic Sans, fuentes "modernas" de otras plantillas).
- Poner texto sobre zonas de la foto con mucho detalle sin una capa de apoyo (como la del cierre).
- Líneas decorativas debajo de los títulos, sombras pesadas, degradados, WordArt.
- Más de una idea por diapositiva; párrafos largos; más de 6 viñetas.
- Letra por debajo de 14 pt en el cuerpo para "hacer caber" el contenido: divide.
- Dejar textos de ejemplo de la plantilla o placeholders vacíos.
