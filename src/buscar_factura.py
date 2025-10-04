import os

EXTENSIONES = ['.jpg', '.jpeg', '.png', '.pdf']

FACTURAS_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'facturas')

def buscar_factura(codigo_factura):
    """
    Busca el archivo de factura por código en la carpeta /facturas/.
    Devuelve la ruta si existe, o None si no existe.
    """
    for ext in EXTENSIONES:
        archivo = os.path.join(FACTURAS_DIR, f"{codigo_factura}{ext}")
        if os.path.exists(archivo):
            return archivo
    return None
