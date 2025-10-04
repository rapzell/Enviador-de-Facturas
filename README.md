# Envío automatizado de facturas

Este sistema permite analizar una imagen de tabla (tipo Excel impreso), extraer los datos de comunidades, meses y facturas, y enviar automáticamente las facturas por correo electrónico.

## Estructura principal
- `/src/leer_tabla.py`: OCR y parsing de la tabla desde la imagen.
- `/src/buscar_factura.py`: Búsqueda y validación de archivos de factura.
- `/src/enviar_factura.py`: Composición y envío de correos.
- `/src/procesar_envios.py`: Lógica principal de procesamiento.
- `/src/gui/interface.py`: Interfaz gráfica mínima.
- `/facturas/`: Carpeta para tus archivos de facturas.
- `start_gui.py`: Lanza la interfaz gráfica.

## Instalación

1. Instala dependencias:
   ```
   pip install -r requirements.txt
   ```
2. Instala Tesseract OCR en tu sistema (https://github.com/tesseract-ocr/tesseract)

## Uso

1. Coloca la imagen de la tabla y las facturas en las carpetas correspondientes.
2. Ejecuta:
   ```
   python start_gui.py
   ```
3. Completa los campos y pulsa "Ejecutar envío de facturas".

## Notas
- Puedes ampliar los formatos soportados editando la lista EXTENSIONES en `buscar_factura.py`.
- El parsing OCR de la tabla es un ejemplo y debe adaptarse a tu formato real.
