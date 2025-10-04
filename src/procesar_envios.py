import os
import glob
import re
import unicodedata
from enviar_factura import enviar_factura as enviar_correo # Alias para consistencia con GUI

def procesar_nombre_factura(nombre_archivo):
    """
    Extrae el nombre de la comunidad y el número de factura del nombre de un archivo PDF.
    Normaliza el nombre de la comunidad a minúsculas para facilitar la coincidencia.
    """
    # Eliminar la extensión .pdf
    nombre_base = os.path.splitext(nombre_archivo)[0]
    
    # Intentar dividir por el primer guion bajo
    partes = nombre_base.split('_', 1)
    
    if len(partes) == 2 and partes[0].isdigit():
        numero_factura = partes[0]
        nombre_comunidad = partes[1].strip()
    else:
        # Si no sigue el formato, se usa el nombre del archivo sin extensión como fallback.
        numero_factura = "N/A"
        nombre_comunidad = nombre_base.replace('_', ' ').strip()
        
    # Normalizar el nombre de la comunidad a minúsculas
    return nombre_comunidad.lower(), numero_factura

def _normalizar_nombre(texto: str) -> str:
    if texto is None:
        return ""
    s = str(texto).strip().lower()
    # Quitar acentos
    s = ''.join(c for c in unicodedata.normalize('NFKD', s) if not unicodedata.combining(c))
    # Unificar separadores y espacios
    s = s.replace('\u00ba', 'º')
    s = s.replace('\n', ' ').replace('\t', ' ')
    s = re.sub(r"[\s\-_/]+", " ", s)
    s = re.sub(r"\s+", " ", s)
    s = s.strip()
    return s

# Las funciones leer_tabla y buscar_factura no se usan en el nuevo flujo.

def procesar_envios(directorio_pdfs, mapeo_comunidades_excel, log_callback, stop_callback):
    """
    Escanea un directorio en busca de archivos PDF, extrae nombres de comunidades de los nombres de archivo
    y busca correos electrónicos correspondientes en el mapeo_comunidades_excel.
    Prepara una lista de datos para la ventana de asignación de correos en la GUI.

    Args:
        directorio_pdfs (str): Ruta al directorio que contiene los archivos PDF.
        mapeo_comunidades_excel (dict): Diccionario con nombres de comunidad como claves y correos como valores.
        log_callback (function): Función para registrar mensajes de progreso/error.
        stop_callback (function): Función que devuelve True si el proceso debe detenerse.

    Returns:
        list: Una lista de diccionarios, cada uno representando un PDF y su información asociada.
              Ej: [{'pdf_path': '...', 'nombre_comunidad': '...', 'correo_asignado': '...'}]
    """
    if not log_callback:
        log_callback = print # Fallback si no se proporciona log_callback

    log_callback(f"Iniciando procesamiento de PDFs en: {directorio_pdfs}")
    archivos_pdf = glob.glob(os.path.join(directorio_pdfs, '*.pdf'))
    log_callback(f"Se encontraron {len(archivos_pdf)} archivos PDF.")

    resultados_procesamiento = []

    for i, pdf_path in enumerate(archivos_pdf):
        if stop_callback():
            log_callback("Proceso detenido por el usuario.")
            break
        
        nombre_archivo = os.path.basename(pdf_path)
        nombre_comunidad, numero_factura = procesar_nombre_factura(nombre_archivo)
        nombre_comunidad_norm = _normalizar_nombre(nombre_comunidad)

        # Búsqueda en el mapeo de Excel usando normalización consistente con la carga del Excel
        correo_asignado = mapeo_comunidades_excel.get(nombre_comunidad_norm)

        if correo_asignado:
            log_callback(f"[{i+1}/{len(archivos_pdf)}] '{nombre_comunidad}' -> Correo: {correo_asignado} (Mapeo Excel)")
        else:
            log_callback(f"ADVERTENCIA: [{i+1}/{len(archivos_pdf)}] '{nombre_comunidad}' no encontrado en mapeo Excel. Se requerirá asignación manual.")

        data_item = {
            'ruta_pdf': pdf_path,
            'numero_factura': numero_factura,
            'nombre_comunidad': nombre_comunidad,
            'correo_asignado': correo_asignado if correo_asignado else "", # Dejar vacío si no hay mapeo
        }
        resultados_procesamiento.append(data_item)

    log_callback(f"Procesamiento de {len(resultados_procesamiento)} PDFs completado.")
    return resultados_procesamiento

