import os
from pdf2image import convert_from_path
import easyocr
import tempfile
import numpy as np
from PIL import Image
import re

def extraer_comunidad_de_pdf(pdf_path, reader):
    """Extrae el nombre de la comunidad de un PDF usando un reader de EasyOCR preinicializado."""
    if reader is None:
        # Esto es un fallback por si acaso, aunque no debería ocurrir si se usa correctamente.
        # Considerar lanzar un error o loguear una advertencia severa.
        print("ADVERTENCIA: El reader de EasyOCR no fue proporcionado. Se inicializará uno nuevo (esto es ineficiente).")
        reader = easyocr.Reader(['es'], verbose=False)

    # Convierte la primera página a imagen temporal
    with tempfile.TemporaryDirectory() as tempdir:
        try:
            images = convert_from_path(
                pdf_path,
                dpi=300,
                first_page=1,
                last_page=1,
                poppler_path=r"C:\Users\David\Desktop\PROGRAMAR\poppler-24.08.0\Library\bin"
            )
        except Exception as e_pdf2img:
            print(f"Error convirtiendo PDF a imagen ({os.path.basename(pdf_path)}): {e_pdf2img}")
            return '[ERROR CONVERSION PDF]'
            
        if not images:
            return '[NO SE PUDO CONVERTIR PDF]'
        
        img = images[0]
        # w, h = img.size # No se usa w, h directamente
        
        comunidad = None
        prop_box = None
        
        try:
            # OCR de toda la página con bounding boxes
            ocr_boxes = reader.readtext(np.array(img), detail=1, paragraph=False)
        except Exception as e_ocr:
            print(f"Error durante OCR en ({os.path.basename(pdf_path)}): {e_ocr}")
            return '[ERROR OCR]'

        # Expresión regular más flexible para variantes de 'Cdad. de Propietarios'
        patron_prop = re.compile(r"cdad\s*\.?\s*de\s*propietarios", re.IGNORECASE)
        patron_prop_simple = re.compile(r"cdad\s*\.?\s*propietarios", re.IGNORECASE)
        patron_solo_prop = re.compile(r"propietarios", re.IGNORECASE)
        
        for box, texto, conf in ocr_boxes:
            if patron_prop.search(texto) or patron_prop_simple.search(texto):
                prop_box = box
                break
        
        if not prop_box:
            for box, texto, conf in ocr_boxes:
                # Usar fullmatch para 'Propietarios' solo si es la palabra exacta, ignorando espacios alrededor
                if patron_solo_prop.fullmatch(texto.strip()): 
                    prop_box = box
                    break
        
        if prop_box:
            x_centro = sum([p[0] for p in prop_box]) / 4
            y_base = max([p[1] for p in prop_box])
            mejor_candidato_texto = None
            min_distancia_y = float('inf')
            
            for candid_box, candid_texto, candid_conf in ocr_boxes:
                # Asegurarse que la caja candidata está debajo de prop_box
                y_candid_top = min([p[1] for p in candid_box])
                x_candid_centro = sum([p[0] for p in candid_box]) / 4
                
                if y_candid_top > y_base: 
                    distancia_y_actual = y_candid_top - y_base
                    # Verificar alineación horizontal (dentro de un umbral)
                    if abs(x_candid_centro - x_centro) < 100: # Umbral de 100px para alineación X
                        # Verificar que no sea una palabra clave no deseada y que tenga contenido
                        texto_limpio = candid_texto.strip()
                        if texto_limpio and not any(palabra_clave in texto_limpio.upper() for palabra_clave in ["FACTURA", "FECHA", "NRO", "TOTAL", "SUBTOTAL", "I.V.A", "FIRMA", "SELLO", "CANTIDAD"]):
                            if distancia_y_actual < min_distancia_y:
                                min_distancia_y = distancia_y_actual
                                mejor_candidato_texto = texto_limpio
            
            comunidad = mejor_candidato_texto if mejor_candidato_texto else '[COMUNIDAD NO DETECTADA POST-PROP]' # Más específico
        else:
            comunidad = '[ETIQUETA PROPIETARIOS NO ENCONTRADA]'

        # Fallback: método antiguo si no se detecta nada con el método de cajas
        # Este fallback podría necesitar también el reader, pero ya lo tiene.
        if comunidad and ('[' in comunidad and ']' in comunidad): # Si los métodos anteriores fallaron en encontrar algo concreto
            try:
                ocr_result_paragraph = reader.readtext(np.array(img), detail=0, paragraph=True)
                texto_completo_parrafos = '\n'.join(ocr_result_paragraph)
                lineas_parrafos = texto_completo_parrafos.split('\n')
                comunidad_fallback = None
                for i, linea in enumerate(lineas_parrafos):
                    if 'cdad. de propietarios' in linea.lower():
                        # Buscar en las siguientes líneas
                        for siguiente_linea in lineas_parrafos[i+1:]:
                            siguiente_linea_limpia = siguiente_linea.strip()
                            if siguiente_linea_limpia and not any(palabra_clave in siguiente_linea_limpia.upper() for palabra_clave in ["FACTURA", "FECHA", "NRO", "TOTAL", "SUBTOTAL", "I.V.A", "FIRMA", "SELLO"]):
                                # Evitar que sea un número largo (como un CIF o teléfono)
                                if sum(c.isdigit() for c in siguiente_linea_limpia) < 6:
                                    comunidad_fallback = siguiente_linea_limpia
                                    break # Tomar la primera coincidencia válida
                        if comunidad_fallback: break # Salir del bucle exterior si se encontró
                if comunidad_fallback:
                    comunidad = comunidad_fallback
                # else: # Si el fallback tampoco encuentra, se mantiene el mensaje de error anterior
            except Exception as e_ocr_fallback:
                print(f"Error durante OCR fallback en ({os.path.basename(pdf_path)}): {e_ocr_fallback}")
                # Mantener el mensaje de error anterior si el fallback falla

        return comunidad if comunidad else '[NO DETECTADA FINAL]'

def extraer_comunidades_de_carpeta(carpeta, reader_global=None):
    """Extrae comunidades de todos los PDFs en una carpeta, usando un reader opcional preinicializado."""
    if reader_global is None:
        print("Inicializando reader de EasyOCR para extraer_comunidades_de_carpeta (uso de prueba)...")
        reader_global = easyocr.Reader(['es'], verbose=False) # Inicializar si no se provee
        
    pdfs = [os.path.join(carpeta, f) for f in os.listdir(carpeta) if f.lower().endswith('.pdf')]
    resultado = {}
    for pdf_path_iter in pdfs:
        nombre_base_pdf = os.path.basename(pdf_path_iter)
        print(f"[DEBUG] Procesando (desde carpeta): {nombre_base_pdf}")
        comunidad_extraida = extraer_comunidad_de_pdf(pdf_path_iter, reader_global)
        resultado[nombre_base_pdf] = comunidad_extraida
        print(f"[DEBUG] {nombre_base_pdf} => {comunidad_extraida}")
    return resultado

if __name__ == "__main__":
    print("Ejecutando prueba de extracción de comunidades...")
    # Inicializar el reader una vez para todas las operaciones en esta prueba
    reader_main = easyocr.Reader(['es'], verbose=False)
    print("Reader de EasyOCR inicializado para la prueba.")
    
    carpeta_pruebas = r"C:\Users\David\Desktop\PROGRAMAR\ENVIADOR DE FACTURAS\FACTURAS\FACTURAS COMPLETA\05 - Mayo - 2025"
    # carpeta_pruebas = r"ruta/a/tu/carpeta/de/pdfs"
    if os.path.isdir(carpeta_pruebas):
        resultados_prueba = extraer_comunidades_de_carpeta(carpeta_pruebas, reader_main)
        print("\n--- Resultados de la prueba ---")
        for nombre_pdf, comunidad_detectada in resultados_prueba.items():
            print(f"{nombre_pdf}: {comunidad_detectada}")
    else:
        print(f"Error: La carpeta de pruebas '{carpeta_pruebas}' no existe.")
    print("Prueba finalizada.")
