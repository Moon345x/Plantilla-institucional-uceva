#!/usr/bin/env python3
"""Crea salida/todos_los_tipos.pptx con una diapositiva de cada tipo de la librería
(incluye casos difíciles: título de portada largo, subtítulo en dos líneas, 6 bloques).
Sirve como catálogo vivo y como prueba de humo. Datos inventados."""
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ / "lib"))

from uceva_pptx import (abrir_plantilla, bienvenida, cierre, copiar_muestra, diapositiva_barra_lateral,  # noqa: E402
                        diapositiva_contenido, diapositiva_en_blanco, diapositiva_foto,
                        diapositiva_lista_de_bloques, guardar, numero_grande, portada, reemplazar_texto, texto)

FOTO = RAIZ / "assets" / "img" / "foto_ejemplo_campus.jpeg"
prs = abrir_plantilla()

bienvenida(prs, notas="Diapositiva de bienvenida (muestra 1).")
portada(prs, "Semillero de investigación en análisis de datos educativos abiertos",
        "Informe de avance del primer semestre con resultados preliminares",
        autor="Grupo de ejemplo", fecha="Marzo de 2026")

s, area = diapositiva_contenido(prs, "Facultad de Ciencias de la Salud", "Etiqueta rosada", "El título toma el color rosado",
                                color="rosado")
texto(s, area, "Texto de cuerpo en gris oscuro, 16 pt. **Negrita** con asteriscos dobles.")

s, area = diapositiva_barra_lateral(prs, "Sección 2: Metodología", "Barra lateral morada con logo blanco")
texto(s, area, ["Paso 1: recoger datos", "Paso 2: limpiarlos", ("Quitar duplicados", 1), "Paso 3: analizarlos"], vinetas=True, tam=20)

s, area = diapositiva_barra_lateral(prs, "Sección 3: Resultados", color="rosado")
izq, der = area.columnas(2)
numero_grande(s, izq.x, izq.y, izq.ancho, "3,8", "promedio inventado", color="rosado")
numero_grande(s, der.x, der.y, der.ancho, "+12 %", "mejora inventada")

diapositiva_foto(prs, FOTO, "Bloque de laboratorios", "Foto de ejemplo incluida en la plantilla",
                 notas="Muestra 9: foto con pie.")

diapositiva_lista_de_bloques(prs, [
    ("Objetivo", "Medir el uso de la biblioteca digital."),
    ("Población", "Estudiantes de primer semestre (ejemplo)."),
    ("Instrumento", "Encuesta corta de diez preguntas."),
    ("Resultado", "Informe con recomendaciones."),
], imagen_derecha=FOTO)

diapositiva_lista_de_bloques(prs, [(f"Bloque {i}", f"Texto corto del bloque {i}.") for i in range(1, 7)],
                             titulo="Seis bloques en dos columnas")

s, area = diapositiva_en_blanco(prs)
texto(s, area, "Diapositiva en blanco (muestra 3): el área libre empieza debajo del logo.", tam=20)

s = copiar_muestra(prs, 6)
reemplazar_texto(s, "Facultad de Ciencias Jurídicas y Humanísticas", "Facultad de Ingeniería")
reemplazar_texto(s, "Subtítulo de ejemplo", "Copia exacta de la muestra 6")
reemplazar_texto(s, "Este es un texto corto de ejemplo", "Hecha con copiar_muestra() y reemplazar_texto().")

cierre(prs, "¡Gracias por su atención!", "contacto@ejemplo.edu.co")

ruta = guardar(prs, RAIZ / "salida" / "todos_los_tipos.pptx")
print(f"Creado: {ruta.relative_to(RAIZ)} ({len(prs.slides)} diapositivas)")
