# Enviador de Facturas (GUI)

Aplicación de escritorio en Python para enviar facturas PDF por correo electrónico a comunidades, mapeando el nombre de la comunidad con su email a partir de un archivo Excel.

## Tecnologías
- **GUI**: `tkinter` (`ttk`, `ScrolledText`).
- **Excel**: `pandas` + `openpyxl`.
- **Emails**: `smtplib`, `email.mime` (Gmail SMTP).
- **Texto**: `re`, `unicodedata` (normalización de nombres y detección de emails).
- **Varios**: `threading`, `os`.

## Requisitos
- Python 3.9+ (Windows: puedes usar el lanzador `py`).
- Dependencias Python:
  ```bash
  pip install -r requirements.txt
  ```
- Habilitar "App Password" en Gmail para el remitente (SMTP).

## Estructura clave
- `src/gui/interface_funcional_DEFINITIVA_2.py`: Interfaz principal (selección de carpeta PDFs y Excel, logs, confirmación y envío).
- `src/procesar_envios.py`: Lee PDFs, extrae nombre de comunidad y lo normaliza.
- `src/enviar_factura.py`: Envía correos con adjunto PDF mediante Gmail SMTP.
- `requirements.txt`: Dependencias.

## Cómo ejecutar
Desde la raíz del proyecto:

```bash
# Windows
py src\gui\interface_funcional_DEFINITIVA_2.py

# Alternativa
py start_gui.py
```

### Campos en la GUI
- Directorio de PDFs: carpeta que contiene directamente los .pdf (no recursivo).
- Remitente (Gmail) y App Password.
- Archivo Excel de Mapeo.

Pulsa "Iniciar Proceso" para cargar el mapeo, preprocesar PDFs y abrir la ventana de confirmación para editar emails antes de enviar.

## Formato del Excel (mapeo comunidad → email)
La aplicación soporta dos disposiciones comunes por hoja (se analizan todas las hojas):
- **Por filas**: una fila contiene el email en alguna celda y, en esa misma fila, el resto de celdas son nombres de comunidades asociados a ese email.
- **Por columnas (fallback)**: una columna contiene un email en alguna de sus celdas (p. ej. cabecera) y el resto de celdas de esa columna son comunidades asociadas.

Reglas de lectura:
- Se detectan emails en cualquier celda (se extrae del texto si está mezclado).
- Cada comunidad se normaliza: minúsculas, sin acentos, separadores unificados y espacios colapsados.
- Si una comunidad aparece varias veces con correos distintos, prevalece la última ocurrencia y se loguea una advertencia.

## Nombres de archivos PDF
- Se espera que el nombre de archivo siga el patrón: `NNN_Nombre de la comunidad.pdf` (ej.: `123_Urb. Pie de Rey blq.1.pdf`).
- El programa elimina el prefijo `NNN_` y normaliza el nombre resultante para buscarlo en el mapeo del Excel.

## Logs y reporte
- Los logs se muestran en la ventana principal.
- Al finalizar, se genera un reporte en `logs/` con el resumen (enviados, errores, saltados).

## Problemas comunes
- **Mapeo = 0 comunidades**: verifica que el Excel tenga emails en alguna celda de cada fila/columna útil. Prueba con otra hoja. Asegúrate de seleccionar la carpeta de PDFs correcta (no el directorio padre).
- **No se encuentra la comunidad**: revisa el nombre exacto del PDF tras quitar `NNN_` y que exista en el Excel (la normalización elimina acentos y unifica separadores).
- **SMTP Gmail**: usa una "App Password" y activa acceso SMTP. Revisa bloqueos de Google si falla la autenticación.

## Seguridad
- No almacenes contraseñas en el repositorio. `.gitignore` excluye `logs/`, artefactos de build y binarios pesados.

## Contribuir
PRs y sugerencias bienvenidos. Para cambios grandes, abre primero un issue describiendo el caso de uso.

