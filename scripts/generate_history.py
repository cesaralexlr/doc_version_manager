import html as html_lib
import os
import re
import subprocess
import sys
from datetime import datetime
from docx import Document
from google import genai
from google.genai import errors

SOURCE_DIR = "source"
OUTPUT_DIR = "output"

# Meses en español para evitar dependencia del locale del sistema
MESES_ES = [
    "enero",
    "febrero",
    "marzo",
    "abril",
    "mayo",
    "junio",
    "julio",
    "agosto",
    "septiembre",
    "octubre",
    "noviembre",
    "diciembre",
]


def fecha_en_espanol(d: datetime) -> str:
  """Devuelve la fecha en formato español: '26 de septiembre de 2026'."""
  return f"{d.day} de {MESES_ES[d.month - 1]} de {d.year}"


def extraer_metadatos_docx(ruta_docx: str) -> tuple[str, str]:
    """
    Busca 'Versión X.Y' y 'Última modificación' en los primeros 10 párrafos del .docx.
    """
    try:
        doc = Document(ruta_docx)
        # Concatenar los primeros 10 párrafos no vacíos con espacio en blanco
        primeros = " ".join(
            p.text for p in doc.paragraphs[:10] if p.text.strip()
        )
    except Exception as e:
        print(f"  [advertencia] No se pudo leer el .docx: {e}")
        primeros = ""

    m_ver = re.search(
        r'(?:versi[oó]n|v)\s*:?\s*([\d]+\.[\d]+)',
        primeros,
        re.IGNORECASE
    )
    version = m_ver.group(1).strip() if m_ver else "N/D"

    m_fec = re.search(
        r'[ÚU]ltima\s+modificaci[oó]n\s+(?:el\s+)?(\d{1,2}\s+de\s+[a-zA-Z]+\s+de\s+\d{4})',
        primeros,
        re.IGNORECASE
    )
    if m_fec:
        fecha = m_fec.group(1).strip()
    else:
        fecha = fecha_en_espanol(datetime.now())

    return version, fecha


def convertir_con_pandoc(
    docx_path: str, html_output_path: str, version: str, fecha: str
) -> None:

  media_dir = OUTPUT_DIR
  os.makedirs(media_dir, exist_ok=True)

  cmd = [
      "pandoc",
      docx_path,
      "-f",
      "docx+styles+empty_paragraphs",
      "-t",
      "html5",
      f"--extract-media={media_dir}",
      "-o",
      html_output_path,
      "--standalone",
      "--toc",
      "--toc-depth=3",
      "--number-sections",
      "--css=styles.css",
  ]
  try:
    subprocess.run(cmd, check=True)
  except subprocess.CalledProcessError as e:
    print(f"  [error] Fallo al ejecutar Pandoc: {e}")
    raise

  with open(html_output_path, "r", encoding="utf-8") as f:
    contenido_html = f.read()

    # 1. Corrección y estandarización de rutas de imágenes a formato web (/)
    contenido_html = re.sub(
        r'src=["\'][^"\']*?media[\\/](?:media[\\/])?([^"\'\\/]+)["\']',
        r'src="media/\1"',
        contenido_html
    )

    contenido_html = re.sub(
        r'(<img[^>]*?)\s+style="[^"]*"',
        r'\1', contenido_html
    )

    # 2. Eliminar el header redundante de Pandoc ANTES de limpiar párrafos
    contenido_html = re.sub(
        r'<header id="title-block-header">.*?</header>',
        '',
        contenido_html,
        flags=re.DOTALL
    )

    # 3. Eliminar líneas divisorias horizontales redundantes
    contenido_html = contenido_html.replace("<hr />", "").replace("<hr>", "")

    # 4. Limpieza profunda de párrafos vacíos
    #    Ahora captura también <p> con <span> vacío dentro
    contenido_html = re.sub(
        r'<p>\s*(?:&nbsp;|<br\s*/?>|\u200b|&ensp;|&emsp;|<span[^>]*>\s*</span>)*\s*</p>',
        '',
        contenido_html,
        flags=re.IGNORECASE
    )

    # 5. Eliminar <br/> huérfanos entre párrafos (generan espacios fantasma)
    contenido_html = re.sub(
        r'</p>\s*(?:<br\s*/?>\s*)+<p>',
        '</p><p>',
        contenido_html,
        flags=re.IGNORECASE
    )

    # 6. Colapsar múltiples <br/> consecutivos a uno solo
    contenido_html = re.sub(
        r'(?:<br\s*/?>\s*){2,}',
        '<br />',
        contenido_html,
        flags=re.IGNORECASE
    )

    # 7. Integrar el enlace "Volver al historial" como el primer ítem del TOC
    if '<nav id="TOC" role="doc-toc">' in contenido_html:
        enlace_item = (
            '<li class="toc-history-link">'
            '<a href="index.html" style="font-weight: bold; color: #e74c3c;">'
            '← Volver al Historial</a></li>'
        )
        contenido_html = re.sub(
            r'(<nav id="TOC"[^>]*>\s*<ul>)',
            r'\1\n' + enlace_item,
            contenido_html,
            count=1
        )

    # 8. Inyectar Barra Corporativa Fija limpia
    banner_superior = (
        '<header class="top-corporate-bar">\n'
        '    <div class="corporate-info">\n'
        '        <div class="doc-title-bar">Manual de Registro de Medicamentos en TASY</div>\n'
        f'        <div class="doc-version-tag">Centro Médico ABC | Versión {version} | Última modificación: {fecha}</div>\n'
        '    </div>\n'
        '</header>\n'
    )
    contenido_html = contenido_html.replace("<body>", f"<body>\n{banner_superior}")
    

  with open(html_output_path, "w", encoding="utf-8") as f:
    f.write(contenido_html)

  print(f"  Convertido e impreso: {html_output_path}")


def leer_documento(ruta_docx: str) -> str:
  """Extrae el texto plano del documento Word para análisis."""
  try:
    doc = Document(ruta_docx)
    return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
  except Exception as e:
    print(f"  [error] No se pudo leer el archivo Word: {e}")
    return ""


def leer_version_previa(output_dir: str) -> str | None:
  """Lee el HTML más reciente de output/ (excluyendo index.html) y extrae solo texto plano para Gemini."""
  if not os.path.exists(output_dir):
    return None
  htmls = [
      f
      for f in os.listdir(output_dir)
      if f.endswith(".html") and f != "index.html"
  ]
  if not htmls:
    return None
  htmls.sort(
      key=lambda f: os.path.getmtime(os.path.join(output_dir, f)), reverse=True
  )
  try:
    with open(os.path.join(output_dir, htmls[0]), "r", encoding="utf-8") as f:
      contenido = f.read()
      # Extraer texto plano eliminando etiquetas HTML
      texto_limpio = re.sub(r"<[^>]+>", " ", contenido)
      texto_limpio = re.sub(r"\s+", " ", texto_limpio)
      return texto_limpio[:6000]
  except Exception:
    return None


def generar_resumen_gemini(texto_actual: str, texto_previo: str | None) -> str:
  """Envía el contenido a Gemini para obtener un registro exclusivo de cambios."""
  api_key = os.environ.get("GEMINI_API_KEY")
  if not api_key:
    print("  [error] No se encontró la variable de entorno GEMINI_API_KEY.")
    return "Resumen de cambios no disponible (falta configurar la API Key)."

  client = genai.Client(api_key=api_key)

  contexto_previo = (
      texto_previo if texto_previo else "(no hay versión previa registrada)"
  )

  prompt = f"""
Eres un analista de control de versiones de documentación técnica.

Recibirás DOS versiones del MISMO documento técnico, en texto plano.
Tu única tarea es reportar QUÉ TEXTO cambió entre las dos versiones.

QUÉ CUENTA COMO CAMBIO:
- Títulos que cambiaron de nombre, se agregaron o se eliminaron.
- Párrafos cuyo texto cambió, se agregó o se eliminó.
- Listas cuyo contenido cambió.

QUÉ NO CUENTA COMO CAMBIO (ignóralo completamente):
- Numeración de títulos (1, 1.1, 1.2 → 1, 1.1, 1.3).
- Cambios en la estructura, orden o presentación.
- Metadatos (versión, fecha, autor).
- Índices, tablas de contenido, TOC.
- Cualquier mención a estilos, colores, formato o diseño.
- Cualquier mención a imágenes, tablas, CSS o HTML.

FORMATO DE SALIDA:
- Máximo 4 viñetas, una por cambio.
- Cada viñeta describe el cambio en texto plano.
- Si un título cambió: "Título 'X' cambiado a 'Y'".
- Si un párrafo se agregó: "Nuevo párrafo en [sección]: primeras palabras…".
- Si un párrafo se eliminó: "Eliminado párrafo en [sección]: primeras palabras…".
- Si el texto es idéntico: escribe exactamente "Sin cambios en el contenido textual."

---

VERSIÓN ANTERIOR (texto plano):
{contexto_previo}

---

VERSIÓN ACTUAL (texto plano):
{texto_actual[:6000]}
"""
  
  try:
    response = client.models.generate_content(
        model="gemini-3.8-flash",
        contents=prompt,
    )
    return response.text
  except errors.APIError as e:
    print(f"  [error] Gemini API: {e}")
    return (
        "No se pudo generar el registro de cambios debido a un error en la API."
    )


def actualizar_indice(
    historial_path: str,
    version_nombre: str,
    html_filename: str,
    resumen: str,
) -> None:
  """Actualiza o crea un archivo índice HTML que almacena el historial."""
  fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
  resumen_seguro = html_lib.escape(resumen).replace("\n", "<br>")

  nueva_entrada = f"""
    <div style="border: 1px solid #ddd; padding: 20px; margin-bottom: 20px; border-radius: 8px; background: #fff; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
        <h3 style="margin-top: 0; color: #2c3e50;">Versión: {html_lib.escape(version_nombre)} <span style="font-size: 0.8em; color: #666; font-weight: normal;">({fecha_actual})</span></h3>
        <p><a href="{html_lib.escape(html_filename)}" target="_blank" style="color: #3498db; text-decoration: none; font-weight: bold;">📄 Abrir versión HTML de lectura</a></p>
        <div style="margin-top: 10px;"><strong>Registro de cambios vs versión anterior:</strong></div>
        <div style="background: #f8f9fa; padding: 10px; border-left: 4px solid #3498db; margin-top: 5px;">{resumen_seguro}</div>
    </div>
    """

  if os.path.exists(historial_path):
    with open(historial_path, "r", encoding="utf-8") as f:
      contenido_actual = f.read()
    if "<!-- HISTORIAL_INJECT -->" in contenido_actual:
      contenido_nuevo = contenido_actual.replace(
          "<!-- HISTORIAL_INJECT -->",
          f"<!-- HISTORIAL_INJECT -->\n{nueva_entrada}",
      )
    else:
      contenido_nuevo = contenido_actual + "\n" + nueva_entrada
  else:
    contenido_nuevo = f"""<!DOCTYPE html>
        <html lang="es">
        <head>
            <meta charset="UTF-8">
            <title>Historial de Versiones - Documentación</title>
            <style>
                body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; max-width: 900px; margin: 40px auto; padding: 0 20px; background: #f4f6f7; color: #333; line-height: 1.6; }}
                h1 {{ color: #2c3e50; border-bottom: 2px solid #ccc; padding-bottom: 10px; }}
            </style>
        </head>
        <body>
            <h1>Historial de Versiones y Control de Cambios</h1>
            <p>Registro histórico automatizado de las versiones del documento maestro.</p>
            <!-- HISTORIAL_INJECT -->
            {nueva_entrada}
        </body>
        </html>
        """

  with open(historial_path, "w", encoding="utf-8") as f:
    f.write(contenido_nuevo)


if __name__ == "__main__":
  os.makedirs(OUTPUT_DIR, exist_ok=True)

  if not os.path.exists(SOURCE_DIR):
    print(f"La carpeta '{SOURCE_DIR}' no existe.")
    sys.exit(1)

  archivos_docx = [
      f
      for f in os.listdir(SOURCE_DIR)
      if f.endswith(".docx") and not f.startswith("~$")
  ]

  if not archivos_docx:
    print("No se encontraron archivos .docx válidos en la carpeta source/")
    sys.exit(0)

  # Procesar únicamente el .docx más reciente de la carpeta source/
  archivos_docx.sort(
      key=lambda f: os.path.getmtime(os.path.join(SOURCE_DIR, f)), reverse=True
  )
  archivos_docx = [archivos_docx[0]]

  for archivo in archivos_docx:
    ruta_docx = os.path.join(SOURCE_DIR, archivo)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre_base = os.path.splitext(archivo)[0]
    html_filename = f"{nombre_base}_{timestamp}.html"
    html_output_path = os.path.join(OUTPUT_DIR, html_filename)

    print(f"Procesando documento: {archivo}...")

    try:
      # Extraer metadatos ANTES de Pandoc
      version, fecha = extraer_metadatos_docx(ruta_docx)
      print(f"  Versión detectada: {version} | Fecha: {fecha}")

      # Convertir con Pandoc
      convertir_con_pandoc(ruta_docx, html_output_path, version, fecha)

      # Leer texto para Gemini
      texto_doc = leer_documento(ruta_docx)
      texto_previo = leer_version_previa(OUTPUT_DIR)
      resumen = generar_resumen_gemini(texto_doc, texto_previo)

      # Actualizar índice con la versión real (no el nombre del archivo)
      historial_path = os.path.join(OUTPUT_DIR, "index.html")
      actualizar_indice(historial_path, version, html_filename, resumen)
      print(f"  Historial actualizado para versión {version}.")

    except Exception as e:
      print(f"  [error] Fallo procesando {archivo}: {e}")
      continue