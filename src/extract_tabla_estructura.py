import cv2
import numpy as np
import easyocr

def extraer_tabla_estructura(imagen_path):
    """
    Extrae la matriz de la tabla manuscrita de estructura.
    Devuelve: {'meses': [...], 'comunidades': [...], 'facturas': {(comunidad, mes): numero_factura, ...}}
    """
    # Leer imagen y binarizar
    import os, shutil, tempfile
    ruta_abierta = imagen_path
    temp_file = None
    if imagen_path.lower().endswith('.pnj'):
        # Copia temporal como .png
        temp_file = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
        shutil.copyfile(imagen_path, temp_file.name)
        ruta_abierta = temp_file.name
    img = cv2.imread(ruta_abierta)
    if img is None:
        if temp_file:
            os.remove(temp_file.name)
        raise ValueError(f"No se pudo abrir la imagen: {imagen_path}")
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    _, binarizada = cv2.threshold(gray, 180, 255, cv2.THRESH_BINARY_INV)
    # Detectar líneas horizontales y verticales
    kernel_h = cv2.getStructuringElement(cv2.MORPH_RECT, (40,1))
    kernel_v = cv2.getStructuringElement(cv2.MORPH_RECT, (1,20))
    horizontal = cv2.morphologyEx(binarizada, cv2.MORPH_OPEN, kernel_h)
    vertical = cv2.morphologyEx(binarizada, cv2.MORPH_OPEN, kernel_v)
    tabla = cv2.add(horizontal, vertical)
    # Encontrar contornos de celdas
    contours, _ = cv2.findContours(tabla, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
    celdas = [cv2.boundingRect(cnt) for cnt in contours if cv2.contourArea(cnt) > 500]
    # Ordenar celdas por Y, luego X
    celdas = sorted(celdas, key=lambda b: (b[1], b[0]))
    # Agrupar por filas
    filas = []
    fila_actual = []
    last_y = None
    for x, y, w, h in celdas:
        if last_y is None or abs(y - last_y) < 20:
            fila_actual.append((x, y, w, h))
            last_y = y
        else:
            filas.append(sorted(fila_actual, key=lambda b: b[0]))
            fila_actual = [(x, y, w, h)]
            last_y = y
    if fila_actual:
        filas.append(sorted(fila_actual, key=lambda b: b[0]))
    # OCR de cabecera de meses (recorte manual)
    reader = easyocr.Reader(['es'], verbose=False)
    alto_img = img.shape[0]
    ancho_img = img.shape[1]
    y_cabecera_ini = 60  # Ajusta si tu tabla está más arriba/abajo
    y_cabecera_fin = 110
    franja_cabecera = img[y_cabecera_ini:y_cabecera_fin, :]
    ocr_cabecera = reader.readtext(franja_cabecera, detail=0)
    # Limpiar y unir los textos de la cabecera
    cabecera_texto = ' '.join([t.strip() for t in ocr_cabecera if t.strip()])
    import re
    meses_detectados = re.findall(r'(diciemb\w*|enero|febrero|marzo|abril|mayo|junio)', cabecera_texto.lower())
    meses_detectados = [m.replace('diciemb','diciembre') for m in meses_detectados]

    # --- SEGMENTACIÓN AVANZADA DE FILAS PARA EXTRAER COMUNIDADES ---
    x_comunidad_ini = 0
    x_comunidad_fin = 220  # Ajusta según la imagen
    margen_sup = 110  # Salta cabecera
    margen_inf = 10   # Salta pie si lo hay
    alto_img = img.shape[0]
    franja = img[margen_sup:alto_img-margen_inf, x_comunidad_ini:x_comunidad_fin]
    franja_gray = cv2.cvtColor(franja, cv2.COLOR_BGR2GRAY)
    # Binarización fuerte
    franja_bin = cv2.adaptiveThreshold(franja_gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY_INV, 15, 10)
    # Morfología para resaltar líneas horizontales
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (franja.shape[1]//2, 2))
    horizontal = cv2.morphologyEx(franja_bin, cv2.MORPH_OPEN, kernel)
    # Detección de líneas horizontales
    lines = cv2.HoughLinesP(horizontal, 1, np.pi/180, threshold=60, minLineLength=franja.shape[1]//3, maxLineGap=15)
    y_lines = []
    if lines is not None:
        for l in lines:
            for x1, y1, x2, y2 in l:
                y_lines.append(y1)
    y_lines = sorted(set(y_lines))
    # Añadir inicio y fin de la franja para no perder la primera/última
    y_lines = [0] + y_lines + [franja.shape[0]-1]
    # Filtrar líneas muy juntas (ruido)
    y_lines_filtradas = [y_lines[0]]
    for y in y_lines[1:]:
        if y - y_lines_filtradas[-1] > 15:
            y_lines_filtradas.append(y)
    comunidades_detectadas_raw = []
    print(f'[DEBUG] Líneas horizontales detectadas para filas de comunidades: {y_lines_filtradas}')
    for i in range(len(y_lines_filtradas)-1):
        y0 = y_lines_filtradas[i]
        y1 = y_lines_filtradas[i+1]
        if y1-y0 < 12:
            continue
        celda_img = franja[y0:y1, :]
        celda_gray = cv2.cvtColor(celda_img, cv2.COLOR_BGR2GRAY)
        celda_bin = cv2.adaptiveThreshold(celda_gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 15, 10)
        celda_bin = cv2.cvtColor(celda_bin, cv2.COLOR_GRAY2BGR)
        ocr = reader.readtext(celda_bin, detail=0)
        texto = ocr[0].strip().replace('"','').replace("'",'') if ocr else ''
        print(f'[DEBUG] Comunidad detectada (fila {i+1}, y={y0}-{y1}): {texto}')
        comunidades_detectadas_raw.append(texto)
    # --- FILTRO Y LIMPIEZA ---
    comunidades_filtradas = []
    for idx, txt in enumerate(comunidades_detectadas_raw):
        if not txt or len(txt) < 3:
            continue
        if '@' in txt:
            continue
        if txt.isdigit():
            continue
        if txt in comunidades_filtradas:
            continue
        comunidades_filtradas.append(txt)
    print(f'[DEBUG] Comunidades filtradas finales: {comunidades_filtradas}')
    comunidades_detectadas = comunidades_filtradas

    # OCR de cada celda (como antes)
    matriz = []
    for fila in filas:
        fila_texto = []
        for x, y, w, h in fila:
            celda_img = img[y:y+h, x:x+w]
            ocr = reader.readtext(celda_img, detail=0)
            texto = ocr[0].strip() if ocr else ''
            fila_texto.append(texto)
        matriz.append(fila_texto)
    # Procesar matriz
    import logging
    print('[DEBUG] MATRIZ OCR:')
    for fila in matriz:
        print(fila)
    if len(matriz) < 2:
        raise ValueError('No se detectó una tabla válida en la imagen')
    # Normalizar ancho de filas
    ancho_max = max(len(fila) for fila in matriz)
    matriz = [fila + [''] * (ancho_max - len(fila)) if len(fila) < ancho_max else fila[:ancho_max] for fila in matriz]
    # Usar la cabecera de meses extraída manualmente si es válida
    meses = meses_detectados if meses_detectados else matriz[0][1:]
    # Usar la columna de comunidades extraída manualmente si es válida
    comunidades = comunidades_detectadas if len(comunidades_detectadas) > 3 else []
    facturas = {}
    print('[DEBUG] MATRIZ OCR FINAL:')
    for idx, fila in enumerate(matriz[1:]):
        # Si tenemos la columna de comunidades extraída, la usamos para cada fila
        if comunidades:
            if idx < len(comunidades):
                comunidad = comunidades[idx]
            else:
                continue
        else:
            comunidad = fila[0].replace('"', '').replace("'", '').strip()
        if not comunidad or comunidad.isspace():
            continue
        print(f'Fila {idx+1}: Comunidad={comunidad}, Facturas={fila[1:]}')
        for j, celda in enumerate(fila[1:]):
            if j >= len(meses):
                print(f'[WARN] Fila {idx+1} tiene más columnas ({len(fila)-1}) que la cabecera de meses ({len(meses)}). Se ignoran las sobrantes.')
                break
            mes = meses[j]
            if celda and not celda.isspace():
                facturas[(comunidad, mes)] = celda.strip()
    if any(len(fila)-1 != len(meses) for fila in matriz[1:]):
        print(f'[WARN] No todas las filas tienen el mismo número de celdas que la cabecera de meses.')
    print(f'[DEBUG] Comunidades extraídas: {comunidades}')
    print(f'[DEBUG] Meses extraídos (cabecera manual): {meses}')
    print(f'[DEBUG] Facturas extraídas: {facturas}')
    return {'meses': meses, 'comunidades': comunidades, 'facturas': facturas}
