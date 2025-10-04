import re
import unicodedata

def normalizar_texto(texto):
    """
    Normaliza un texto: convierte a minúsculas, quita acentos, 
    reemplaza caracteres no alfanuméricos (excepto ñ) por espacios, 
    y elimina espacios extra.
    """
    if not isinstance(texto, str):
        return "" 
    try:
        texto_norm = texto.lower()
        texto_norm = ''.join(c for c in unicodedata.normalize('NFD', texto_norm) if unicodedata.category(c) != 'Mn')
        texto_norm = re.sub(r'[^a-z0-9ñ\s]', ' ', texto_norm)
        texto_norm = re.sub(r'\s+', ' ', texto_norm).strip()
        return texto_norm
    except Exception as e:
        print(f"Error al normalizar texto '{texto}': {e}")
        return texto
