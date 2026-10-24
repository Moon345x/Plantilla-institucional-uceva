#!/usr/bin/env bash
# Convierte un .pptx en un PNG por diapositiva (por defecto 1280×720).
#
# Uso:
#   scripts/render_previews.sh ARCHIVO.pptx CARPETA_SALIDA [ANCHO]
#
# Resultado: CARPETA_SALIDA/slide01.png, slide02.png, ...
#
# Motor de conversión .pptx → PDF (se usa el primero que exista):
#   1. LibreOffice:  soffice / libreoffice  (--headless --convert-to pdf)
#   2. OnlyOffice:   x2t (viene con OnlyOffice Desktop Editors, en /opt/onlyoffice)
# Luego PDF → PNG con pdftoppm (paquete poppler / poppler-utils).
#
# Se puede forzar el motor con la variable UCEVA_RENDER=libreoffice|onlyoffice.
set -euo pipefail

if [[ $# -lt 2 ]]; then
  echo "Uso: $0 ARCHIVO.pptx CARPETA_SALIDA [ANCHO]" >&2
  exit 2
fi

ENTRADA="$(realpath "$1")"
SALIDA="$2"
ANCHO="${3:-1280}"
ALTO=$(( ANCHO * 9 / 16 ))

[[ -f "$ENTRADA" ]] || { echo "No existe: $ENTRADA" >&2; exit 2; }
command -v pdftoppm >/dev/null || { echo "Falta pdftoppm (instala poppler-utils)." >&2; exit 3; }

mkdir -p "$SALIDA"
SALIDA="$(realpath "$SALIDA")"
TMP="$(mktemp -d)"
trap 'rm -rf "$TMP"' EXIT

cp "$ENTRADA" "$TMP/entrada.pptx"
PDF="$TMP/entrada.pdf"

buscar_soffice() {
  for c in soffice libreoffice; do
    command -v "$c" >/dev/null && { echo "$c"; return; }
  done
  for c in /usr/lib/libreoffice/program/soffice /opt/libreoffice*/program/soffice \
           /Applications/LibreOffice.app/Contents/MacOS/soffice; do
    [[ -x "$c" ]] && { echo "$c"; return; }
  done
}

buscar_x2t() {
  for c in /opt/onlyoffice/desktopeditors/converter/x2t \
           /usr/lib/onlyoffice/desktopeditors/converter/x2t; do
    [[ -x "$c" ]] && { echo "$c"; return; }
  done
}

con_libreoffice() {
  local soffice="$1"
  "$soffice" --headless --norestore --convert-to pdf --outdir "$TMP" "$TMP/entrada.pptx" >/dev/null 2>&1
}

con_onlyoffice() {
  local x2t="$1"
  # x2t necesita la caché de fuentes (AllFonts.js). OnlyOffice Desktop la crea al abrirse una vez.
  local fuentes=""
  for d in "$HOME/.local/share/onlyoffice/desktopeditors/data/fonts" \
           "$HOME/.config/onlyoffice/desktopeditors/data/fonts"; do
    [[ -f "$d/AllFonts.js" ]] && { fuentes="$d"; break; }
  done
  if [[ -z "$fuentes" ]]; then
    echo "OnlyOffice encontrado, pero falta su caché de fuentes. Abre OnlyOffice Desktop una vez y reintenta." >&2
    return 1
  fi
  cat > "$TMP/x2t.xml" <<EOF
<?xml version="1.0" encoding="utf-8"?>
<TaskQueueDataConvert xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
<m_sFileFrom>$TMP/entrada.pptx</m_sFileFrom>
<m_sFileTo>$PDF</m_sFileTo>
<m_nFormatTo>513</m_nFormatTo>
<m_sFontDir>$fuentes</m_sFontDir>
<m_sAllFontsPath>$fuentes/AllFonts.js</m_sAllFontsPath>
<m_sThemeDir>$(dirname "$x2t")/templates</m_sThemeDir>
<m_bDontSaveAdditional>true</m_bDontSaveAdditional>
</TaskQueueDataConvert>
EOF
  "$x2t" "$TMP/x2t.xml" >/dev/null 2>&1
}

MOTOR="${UCEVA_RENDER:-}"
SOFFICE="$(buscar_soffice || true)"
X2T="$(buscar_x2t || true)"
if [[ -z "$MOTOR" ]]; then
  if [[ -n "$SOFFICE" ]]; then MOTOR=libreoffice
  elif [[ -n "$X2T" ]]; then MOTOR=onlyoffice
  else
    echo "No encontré LibreOffice (soffice) ni OnlyOffice (x2t). Instala uno de los dos." >&2
    exit 3
  fi
fi

case "$MOTOR" in
  libreoffice) [[ -n "$SOFFICE" ]] || { echo "Falta soffice" >&2; exit 3; }; con_libreoffice "$SOFFICE" ;;
  onlyoffice)  [[ -n "$X2T" ]] || { echo "Falta x2t" >&2; exit 3; }; con_onlyoffice "$X2T" ;;
  *) echo "UCEVA_RENDER debe ser libreoffice u onlyoffice" >&2; exit 2 ;;
esac

[[ -s "$PDF" ]] || { echo "La conversión a PDF falló (motor: $MOTOR)." >&2; exit 4; }

rm -f "$SALIDA"/slide*.png
pdftoppm -png -scale-to-x "$ANCHO" -scale-to-y "$ALTO" "$PDF" "$TMP/p"
n=0
for f in $(ls "$TMP"/p-*.png | sort -V); do
  n=$((n + 1))
  mv "$f" "$(printf '%s/slide%02d.png' "$SALIDA" "$n")"
done
echo "Motor: $MOTOR · $n diapositivas → $SALIDA"
