print('>>> CARGADO test_interface.py')
import sys
import os
import traceback

# Asegurarse de que el directorio src esté en el PATH
src_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), 'src'))
if src_dir not in sys.path:
    sys.path.append(src_dir)

try:
    print("Intentando importar la interfaz...")
    from gui.interface import run_gui
    
    print("Iniciando la interfaz gráfica...")
    run_gui()
    print("Interfaz cerrada.")
    
except Exception as e:
    print(f"Error al ejecutar la aplicación: {str(e)}")
    print("\nDetalle del error:")
    traceback.print_exc()
    input("Presiona Enter para salir...")
