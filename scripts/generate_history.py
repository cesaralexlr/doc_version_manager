import os
import sys
import subprocess
from datetime import datetime
from docx import Document
from google import genai
from google.genai import errors

SOURCE_DIR = "source"
OUTPUT_DIR = "output"

def convertir_con_pandoc(docx_path, html_output_path):
    """Convierte un archivo docx a html5 usando pandoc, aplicando estilos y extrayendo imágenes."""
    media_dir = os.path.join(OUTPUT_DIR, "media")
    os.makedirs(media_dir, exist_ok=True)
    
    cmd = [
        "pandoc",
        docx_path,
        "-f", "docx",
        "-t", "html5",
        f"--extract-media={media_dir}",
        "-o", html_output_path,
        "--standalone",
        "--css=styles.css"  # Enlaza la hoja de estilos generada
    ]
    try:
        subprocess.run(cmd, check=True)
        print(f"Convertido exitosamente a HTML con estilos: {html_output_path}")
    except subprocess.CalledProcessError as e:
        print(f"Error al ejecutar Pandoc: {e}")
        sys.exit(1)

def leer_documento(ruta_docx):
    """Extrae el texto plano del documento Word para análisis."""
    try:
        doc = Document(ruta_docx)
        texto = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
        return texto
    except Exception as e:
        print(f"Error al leer el archivo Word: {e}")
        return ""

def generar_resumen_gemini(texto_actual):
    """Envía el contenido a Gemini para obtener un resumen ejecutivo y objetivo."""
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        print("Error: No se encontró la variable de entorno GEMINI_API_KEY.")
        return "Resumen no disponible (falta configurar la API Key)."

    client = genai.Client(api_key=api_key)
    
    prompt = f"""
    Actúa como un analista de documentación técnica. 
    Analiza el siguiente texto de un documento de Word y genera un resumen ejecutivo y objetivo de los puntos clave o cambios principales (máximo 4 viñetas directas):

    CONTENIDO DEL DOCUMENTO:
    {texto_actual[:8000]}
    """

    try:
        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=prompt,
        )
        return response.text
    except errors.APIError as e:
        print(f"Error en la API de Gemini: {e}")
        return "No se pudo generar el resumen automático debido a un error en la API."

def actualizar_indice(historial_path, version_nombre, html_filename, resumen):
    """Actualiza o crea un archivo índice HTML que almacena el historial y los enlaces a versiones anteriores."""
    fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    nueva_entrada = f"""
    <div style="border: 1px solid #ddd; padding: 20px; margin-bottom: 20px; border-radius: 8px; background: #fff; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
        <h3 style="margin-top: 0; color: #2c3e50;">Versión: {version_nombre} <span style="font-size: 0.8em; color: #666; font-weight: normal;">({fecha_actual})</span></h3>
        <p><a href="{html_filename}" target="_blank" style="color: #3498db; text-decoration: none; font-weight: bold;">📄 Abrir versión HTML de lectura</a></p>
        <div style="margin-top: 10px;"><strong>Resumen de cambios / Puntos clave:</strong></div>
        <div style="background: #f8f9fa; padding: 10px; border-left: 4px solid #3498db; margin-top: 5px;">{resumen.replace(chr(10), '<br>')}</div>
    </div>
    """
    
    if os.path.exists(historial_path):
        with open(historial_path, "r", encoding="utf-8") as f:
            contenido_actual = f.read()
        if "<!-- HISTORIAL_INJECT -->" in contenido_actual:
            contenido_nuevo = contenido_actual.replace("<!-- HISTORIAL_INJECT -->", f"<!-- HISTORIAL_INJECT -->\n{nueva_entrada}")
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
        
    archivos_docx = [f for f in os.listdir(SOURCE_DIR) if f.endswith(".docx") and not f.startswith("~$")]
    
    if not archivos_docx:
        print("No se encontraron archivos .docx válidos en la carpeta source/")
        sys.exit(0)
        
    for archivo in archivos_docx:
        ruta_docx = os.path.join(SOURCE_DIR, archivo)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nombre_base = os.path.splitext(archivo)[0]
        html_filename = f"{nombre_base}_{timestamp}.html"
        html_output_path = os.path.join(OUTPUT_DIR, html_filename)
        
        print(f"Procesando documento: {archivo}...")
        
        # 1. Convertir documento a HTML limpio con Pandoc
        convertir_con_pandoc(ruta_docx, html_output_path)
        
        # 2. Leer contenido y generar resumen inteligente con Gemini
        texto_doc = leer_documento(ruta_docx)
        resumen = generar_resumen_gemini(texto_doc)
        
        # 3. Registrar en el archivo índice de historial
        historial_path = os.path.join(OUTPUT_DIR, "index.html")
        actualizar_indice(historial_path, nombre_base, html_filename, resumen)
        print(f"Historial e índice actualizados correctamente para {archivo}.")