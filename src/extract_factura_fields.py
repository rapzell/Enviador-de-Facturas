import re
import easyocr
from PIL import Image
import cv2

def extraer_fecha_comunidad(imagen_path, usar_easyocr=True):
    """
    Extrae la fecha (mes y año) y la comunidad (nombre) de una factura estándar.
    Devuelve: {'fecha': '31 de Mayo de 2025', 'mes': 'mayo', 'año': '2025', 'comunidad': '...'}
    """
    # OCR completo de la imagen
    # Verifica que la imagen existe y se puede abrir
    img = cv2.imread(imagen_path)
    if img is None:
        print(f"[ERROR] No se pudo abrir la imagen: {imagen_path}")
        return {'fecha':'','mes':'','año':'','comunidad':'','texto':'','valida':False}
    if usar_easyocr:
        try:
            reader = easyocr.Reader(['es'], verbose=False)
            result = reader.readtext(imagen_path, detail=0, paragraph=True)
            texto = '\n'.join(result)
        except Exception as e:
            print(f"[ERROR] EasyOCR falló para {imagen_path}: {e}")
            return {'fecha':'','mes':'','año':'','comunidad':'','texto':'','valida':False}
    else:
        try:
            pil_img = Image.fromarray(cv2.cvtColor(img, cv2.COLOR_BGR2RGB))
            texto = pytesseract.image_to_string(pil_img, lang='spa')
        except Exception as e:
            print(f"[ERROR] pytesseract falló para {imagen_path}: {e}")
            return {'fecha':'','mes':'','año':'','comunidad':'','texto':'','valida':False}
    # Buscar número de factura
    m_num = re.search(r'FACTURA NRO\.?\s*[:\-]?\s*(\d+)', texto, re.IGNORECASE)
    numero_factura = m_num.group(1) if m_num else ''
    # Buscar línea de fecha
    m_fecha = re.search(r'FECHA:.*?(\d{1,2} de [A-Za-záéíóúñ]+ de \d{4})', texto, re.IGNORECASE)
    fecha = m_fecha.group(1) if m_fecha else ''
    mes = ''
    año = ''
    if fecha:
        m_mes = re.search(r'de ([A-Za-záéíóúñ]+) de (\d{4})', fecha, re.IGNORECASE)
        if m_mes:
            mes = m_mes.group(1).lower()
            año = m_mes.group(2)
    # Buscar comunidad (Cdad. de Propietarios ... o similar)
    comunidad = ''
    m_com = re.search(r'Cdad\. de Propietarios.*?([\w\s\.]+)', texto)
    if m_com:
        comunidad = m_com.group(1).strip()
    else:
        # Alternativa: buscar justo debajo de CIF
        m_cif = re.search(r'C\.I\.F\.:.*?\n(.+)', texto)
        if m_cif:
            comunidad = m_cif.group(1).strip()
    return {'fecha': fecha, 'mes': mes, 'año': año, 'comunidad': comunidad, 'numero_factura': numero_factura, 'texto': texto, 'valida': True}
