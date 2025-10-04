import os
import glob
import re
from normalizar_texto import normalizar_texto # normalizar_texto.py está ahora en el mismo directorio src
from enviar_factura import enviar_factura as enviar_correo # Alias para consistencia con GUI

# Las funciones leer_tabla y buscar_factura no se usan en el nuevo flujo.

def procesar_envios(directorio_pdfs, mapeo_comunidades_excel, log_callback, stop_callback):
    """
    Escanea un directorio en busca de archivos PDF, extrae nombres de comunidades de los nombres de archivo,
    los normaliza y busca correos electrónicos correspondientes en el mapeo_comunidades_excel.
    Prepara una lista de datos para la ventana de asignación de correos en la GUI.

    Args:
        directorio_pdfs (str): Ruta al directorio que contiene los archivos PDF.
        mapeo_comunidades_excel (dict): Diccionario con nombres de comunidad normalizados como claves y correos como valores.
        log_callback (function): Función para registrar mensajes de progreso/error.
        stop_callback (function): Función que devuelve True si el proceso debe detenerse.

    Returns:
        list: Una lista de diccionarios, cada uno representando un PDF y su información asociada.
              Ej: [{'pdf_path': '...', 'nombre_comunidad_original': '...', 
                    'nombre_comunidad_normalizado': '...', 'correo_asignado': '...', 'match_type': '...'}]
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
        # Intenta extraer el nombre de la comunidad del nombre del archivo
        # Este es un ejemplo simple, podría necesitar una lógica más robusta (ej. regex)
        # Suponemos que el nombre de la comunidad es todo antes del primer guion bajo o punto.
        match_nombre = re.match(r"^([^_.]+)", nombre_archivo)
        nombre_comunidad_original = match_nombre.group(1) if match_nombre else nombre_archivo.replace('.pdf', '')
        nombre_comunidad_normalizado = normalizar_texto(nombre_comunidad_original)

        correo_asignado = mapeo_comunidades_excel.get(nombre_comunidad_normalizado)
        match_type = ""

        if correo_asignado:
            match_type = "Mapeo Exacto (Excel)"
            log_callback(f"[{i+1}/{len(archivos_pdf)}] '{nombre_comunidad_original}' -> Correo: {correo_asignado} (Mapeo Excel)")
        else:
            match_type = "Sin Mapeo"
            log_callback(f"ADVERTENCIA: [{i+1}/{len(archivos_pdf)}] '{nombre_comunidad_original}' (Normalizado: '{nombre_comunidad_normalizado}') no encontrado en mapeo Excel. Se requerirá asignación manual.")

        data_item = {
            'pdf_path': pdf_path,
            'nombre_comunidad_original': nombre_comunidad_original,
            'nombre_comunidad_normalizado': nombre_comunidad_normalizado,
            'correo_asignado': correo_asignado if correo_asignado else "", # Dejar vacío si no hay mapeo
            'match_type': match_type,
            'enviar': bool(correo_asignado) # Marcar para enviar por defecto si hay correo
        }
        resultados_procesamiento.append(data_item)

    log_callback(f"Procesamiento de {len(resultados_procesamiento)} PDFs completado.")
    return resultados_procesamiento

