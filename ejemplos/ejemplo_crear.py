#!/usr/bin/env python3
"""Crea salida/ejemplo.pptx: 8 diapositivas variadas a partir de contenido_ejemplo.json.

    python ejemplos/ejemplo_crear.py
    python scripts/validar.py salida/ejemplo.pptx
    bash scripts/render_previews.sh salida/ejemplo.pptx salida/previews
"""
import json
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ / "lib"))

from uceva_pptx import (abrir_plantilla, bloque_codigo, cierre, cuadricula, diapositiva_contenido,  # noqa: E402
                        fila_de_numeros, fila_de_tarjetas, guardar, imagen, portada, tabla, texto)

c = json.loads((RAIZ / "ejemplos" / "contenido_ejemplo.json").read_text(encoding="utf-8"))
fac = c["facultad"]
prs = abrir_plantilla()

# 1. Portada
p = c["portada"]
portada(prs, p["titulo"], p["subtitulo"], autor=p["autor"], notas=p["notas"])

# 2. Agenda en tarjetas
a = c["agenda"]
s, area = diapositiva_contenido(prs, fac, a["titulo"], a["subtitulo"], notas=a["notas"])
fila_de_tarjetas(s, area, a["tarjetas"])

# 3. Tabla (etiqueta rosada)
t = c["tabla"]
s, area = diapositiva_contenido(prs, fac, t["titulo"], color="rosado", notas=t["notas"])
tabla(s, area, t["filas"], anchos=[1, 3, 3], alineaciones=["centro", "izquierda", "izquierda"])

# 4. Código
k = c["codigo"]
s, area = diapositiva_contenido(prs, fac, k["titulo"], k["subtitulo"], notas=k["notas"])
bloque_codigo(s, area, k["codigo"], lenguaje=k["lenguaje"])

# 5. Imagen + texto en dos columnas
im = c["imagen"]
s, area = diapositiva_contenido(prs, fac, im["titulo"], color="amarillo", notas=im["notas"])
izq, der = area.columnas(2, proporciones=[1, 1])
texto(s, izq, im["texto"], vinetas=True, tam=18)
imagen(s, RAIZ / im["imagen"], der, borde=True)

# 6. Cifras grandes
n = c["cifras"]
s, area = diapositiva_contenido(prs, fac, n["titulo"], n["subtitulo"], notas=n["notas"])
fila_de_numeros(s, area, n["cifras"])

# 7. Repaso en cuadrícula 2×2
r = c["repaso"]
s, area = diapositiva_contenido(prs, fac, r["titulo"], notas=r["notas"])
cuadricula(s, area, r["tarjetas"], columnas=2)

# 8. Cierre (queda al final aunque se cree antes)
z = c["cierre"]
cierre(prs, z["texto_principal"], z["texto_secundario"], notas=z["notas"])

ruta = guardar(prs, RAIZ / "salida" / "ejemplo.pptx")
print(f"Creado: {ruta.relative_to(RAIZ)} ({len(prs.slides)} diapositivas)")
