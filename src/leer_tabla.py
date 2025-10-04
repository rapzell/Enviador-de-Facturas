import cv2
import numpy as np
from PIL import Image
import easyocr
import re
import os

def extraer_correos_y_filas(img_rgb, ocr_reader):
    """
    Extrae los correos y las posiciones de las filas de comunidad.
    Devuelve una lista de tuplas (fila_inicio, fila_fin, correo, [comunidades])
    """
    # Convertir a escala de grises y binarizar
    img_gray = cv2.cvtColor(img_rgb, cv2.COLOR_BGR2GRAY)
    _, img_bin = cv2.threshold(img_gray, 180, 255, cv2.THRESH_BINARY_INV)
    # Buscar líneas horizontales largas (cabeceras de bloque)
    horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (img_bin.shape[1]//2, 1))
    detect_horizontal = cv2.morphologyEx(img_bin, cv2.MORPH_OPEN, horizontal_kernel, iterations=2)
    contours, _ = cv2.findContours(detect_horizontal, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    line_positions = [cv2.boundingRect(c)[1] for c in contours if cv2.boundingRect(c)[3] < 10]
    line_positions = sorted(line_positions)
    # OCR para buscar correos
    ocr_result = ocr_reader.readtext(img_rgb)
    correos = [(d[0][1], d[1]) for d in ocr_result if re.search(r'@', d[1])]
    bloques = []
    for i, (y, correo) in enumerate(sorted(correos)):
        y_inicio = y
        y_fin = line_positions[i+1] if i+1 < len(line_positions) else img_rgb.shape[0]
        bloques.append((y_inicio, y_fin, correo))
    return bloques

def segmentar_celdas(img_rgb):
    """
    Devuelve una lista de celdas: (x1, y1, x2, y2)
    """
    img_gray = cv2.cvtColor(img_rgb, cv2.COLOR_BGR2GRAY)
    _, img_bin = cv2.threshold(img_gray, 180, 255, cv2.THRESH_BINARY_INV)
    # Detectar líneas
    vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, img_bin.shape[0]//40))
    horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (img_bin.shape[1]//40, 1))
    vertical_lines = cv2.morphologyEx(img_bin, cv2.MORPH_OPEN, vertical_kernel, iterations=2)
    horizontal_lines = cv2.morphologyEx(img_bin, cv2.MORPH_OPEN, horizontal_kernel, iterations=2)
    grid = cv2.addWeighted(vertical_lines, 0.5, horizontal_lines, 0.5, 0.0)
    contours, _ = cv2.findContours(grid, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
    celdas = [cv2.boundingRect(c) for c in contours if 40 < cv2.boundingRect(c)[2] < img_rgb.shape[1] and 10 < cv2.boundingRect(c)[3] < img_rgb.shape[0]]
    celdas = sorted(celdas, key=lambda b: (b[1], b[0]))
    return celdas

def leer_tabla(imagenes, usar_easyocr=True):
    """
    Extrae comunidades, correos, meses y facturas de una o varias imágenes de tabla manuscrita.
    Devuelve: comunidades (lista de dicts), meses (lista), facturas (dict)
    """
    if isinstance(imagenes, str):
        imagenes = [imagenes]
    reader = easyocr.Reader(['es'], verbose=False) if usar_easyocr else None
    comunidades_total = []
    meses_total = set()
    facturas_total = dict()
    for imagen_path in imagenes:
        img_rgb = cv2.imread(imagen_path)
        if img_rgb is None:
            continue
        # Extraer cabecera de meses
        header_area = img_rgb[0:120, :]
        ocr_header = reader.readtext(header_area) if usar_easyocr else []
        meses = []
        for d in ocr_header:
            txt = d[1].lower()
            if any(m in txt for m in ['enero','febrero','marzo','abril','mayo','junio','diciemb']):
                meses += re.findall(r'(diciemb\w*|enero|febrero|marzo|abril|mayo|junio)', txt)
        meses = [m.replace('diciemb','diciembre') for m in meses]
        if not meses:
            meses = ['diciembre','enero','febrero','marzo','abril','mayo','junio']
        meses_total.update(meses)
        # Extraer bloques de comunidades y correos
        bloques = extraer_correos_y_filas(img_rgb, reader)
        # Segmentar celdas
        celdas = segmentar_celdas(img_rgb)
        # Agrupar celdas por filas (aproximar por coordenada Y)
        filas = {}
        for x, y, w, h in celdas:
            # Validación robusta: asegúrate de que bloque[0] y bloque[1] son enteros
            posibles_bloques = [bloque for bloque in bloques if isinstance(bloque, (list, tuple)) and len(bloque) >= 2 and isinstance(bloque[0], int) and isinstance(bloque[1], int)]
            bloques_validos = [bloque for bloque in posibles_bloques if y >= bloque[0] and y < bloque[1]]
            fila_candidato = min((bloque[1] for bloque in bloques_validos if isinstance(bloque[1], int)), default=y)
            # Asegura que la clave sea SIEMPRE int
            if not isinstance(fila_candidato, int):
                print(f"[DEBUG] fila_candidato no es int: {fila_candidato} (type: {type(fila_candidato)})")
                continue
            fila = fila_candidato
            if fila not in filas:
                filas[fila] = []
            filas[fila].append((x, y, w, h))
        # Depuración: mostrar tipos de claves de filas
        print(f"[DEBUG] Claves de filas: {[type(k) for k in filas.keys()]}")
        # Para cada bloque/correo, asociar comunidades
        for i, (y_inicio, y_fin, correo) in enumerate(bloques):
            # Solo compara si k es int
            filas_bloque = [k for k in filas if isinstance(k, int) and y_inicio < k < y_fin]
            # Depuración: muestra claves de filas y valores de y_inicio/y_fin
            print(f"[DEBUG] y_inicio={y_inicio}, y_fin={y_fin}, claves filas={list(filas.keys())}")
            filas_bloque = sorted(filas_bloque)
            for idx, fila_y in enumerate(filas_bloque):
                celdas_fila = sorted(filas[fila_y], key=lambda b: b[0])
                # Extraer nombre de comunidad de la primera celda
                x0, y0, w0, h0 = celdas_fila[0]
                celda_comunidad_img = img_rgb[y0:y0+h0, x0:x0+w0]
                nombre_comunidad = ''
                if usar_easyocr:
                    ocr_nombre = reader.readtext(celda_comunidad_img)
                    if ocr_nombre:
                        nombre_comunidad = ocr_nombre[0][1].strip()
                else:
                    pil_img = Image.fromarray(cv2.cvtColor(celda_comunidad_img, cv2.COLOR_BGR2RGB))
                    nombre_comunidad = pytesseract.image_to_string(pil_img, config='--psm 7').strip()
                if not nombre_comunidad:
                    nombre_comunidad = f'Fila_{fila_y}'
                comunidades_total.append({'nombre': nombre_comunidad, 'correo': correo})
                for mes_idx, celda in enumerate(celdas_fila[1:len(meses)+1]):
                    x, y, w, h = celda
                    celda_img = img_rgb[y:y+h, x:x+w]
                    texto_celda = ''
                    if usar_easyocr:
                        ocr_celda = reader.readtext(celda_img)
                        if ocr_celda:
                            texto_celda = ocr_celda[0][1]
                    else:
                        pil_img = Image.fromarray(cv2.cvtColor(celda_img, cv2.COLOR_BGR2RGB))
                        texto_celda = pytesseract.image_to_string(pil_img, config='--psm 7')
                    codigo = None
                    m = re.search(r'([A-Z]-\d{2,4})', texto_celda.upper())
                    if m:
                        codigo = m.group(1)
                    if codigo:
                        facturas_total[(nombre_comunidad, meses[mes_idx])] = codigo
    return comunidades_total, list(meses_total), facturas_total
