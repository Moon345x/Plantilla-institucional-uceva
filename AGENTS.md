# Instrucciones para agentes (IA)

Este repo sirve para crear o convertir presentaciones a la **plantilla institucional UCEVA 55 años**
sin descomprimir el `.pptx` ni adivinar medidas. Todo lo necesario ya está extraído y documentado.

## Lee primero (en este orden)

1. `spec/tokens.json` — colores, fuentes, tamaños, márgenes, zonas protegidas y medidas de cada elemento.
2. `docs/GUIA_PARA_IA.md` — qué tipo de diapositiva usar, flujos de crear y convertir, checklist.
3. `lib/README.md` — funciones de `lib/uceva_pptx.py` con ejemplos cortos.

Si necesitas más detalle: `docs/CATALOGO_DE_DIAPOSITIVAS.md` (las 11 muestras con vista previa),
`spec/slides.json` (cada forma de cada muestra), `docs/ERRORES_COMUNES.md`.

## Flujo en 5 pasos

1. **Plan**: una lista de diapositivas, una idea por diapositiva, y el tipo de cada una (tabla de la guía).
2. **Generar** con `lib/uceva_pptx.py` (`abrir_plantilla` → funciones → `guardar`). No edites XML.
3. **Validar**: `python scripts/validar.py salida/archivo.pptx` → debe terminar en `OK` o solo avisos justificados.
4. **Render**: `bash scripts/render_previews.sh salida/archivo.pptx salida/previews` y **mira cada PNG**.
5. **Corregir** lo que se vea mal (desbordes, textos pisados, huecos) y repetir 3-4.

## Las 10 reglas más importantes

1. Abre siempre con `abrir_plantilla()` y guarda con `guardar()` (quita las 11 muestras, deja el cierre al final).
2. No modifiques `plantilla/*.pptx`. Trabaja sobre copias (la librería ya lo hace).
3. Colores solo de `spec/tokens.json` (`morado` 6000AF es el principal). Nada de amarillo o cian como texto sobre blanco.
4. Fuentes: Arial en todo; Arial Black solo en el título de la portada; Courier New solo en código.
5. Tamaños mínimos: cuerpo 14 pt, pies y etiquetas 12 pt; títulos de diapositiva 24 pt.
6. No pongas nada en la zona del logo (arriba a la izquierda) ni encima de la etiqueta (arriba a la derecha).
7. Usa el `Area` que devuelven las funciones; no inventes coordenadas. Márgenes mínimos de 0,4".
8. Si la librería lanza `ContenidoDemasiadoLargo`, **divide** el contenido en más diapositivas; no bajes la letra a mano.
9. Sin líneas decorativas bajo los títulos, sin textos de ejemplo de la plantilla, sin placeholders vacíos.
10. Conserva las notas del orador (`notas=` en cada función) y usa datos reales del usuario o inventados, nunca datos personales de terceros.

## Comandos

```bash
pip install -r requirements.txt
python ejemplos/ejemplo_crear.py                         # ejemplo completo → salida/ejemplo.pptx
python ejemplos/ejemplo_convertir.py ORIGEN.pptx         # convierte otro .pptx → salida/ORIGEN_uceva.pptx
python scripts/validar.py salida/ejemplo.pptx
bash scripts/render_previews.sh salida/ejemplo.pptx salida/previews
python scripts/extraer_plantilla.py                      # solo si cambia la plantilla: regenera spec/ y assets/
```

## Sin Python

Usa `spec/tokens.json` y `spec/slides.json` como fuente de verdad y las imágenes de `assets/img/`.
La sección 7 de `docs/GUIA_PARA_IA.md` explica cómo hacerlo con pptxgenjs o editando XML.

## Git

No hagas commits salvo que el usuario lo pida. Mensajes en español, cortos y claros.
