import sys
import os
import traceback

def main():
    # Asegurarse de que el directorio src esté en el PATH
    src_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), 'src'))
    if src_dir not in sys.path:
        sys.path.append(src_dir)

    try:
        print("=== Iniciando aplicación de envío de facturas ===")
        print(f"Directorio de trabajo: {os.getcwd()}")
        print(f"Python path: {sys.version}")
        print(f"Python path: {sys.path}")
        
        # Importar después de configurar el path
        from gui.interface import run_gui
        
        print("\nIniciando interfaz gráfica...")
        run_gui()
        print("Aplicación cerrada correctamente.")
        
    except ImportError as e:
        print(f"\nError de importación: {e}")
        print("\nTraceback completo:")
        traceback.print_exc()
        input("\nPresiona Enter para salir...")
    except Exception as e:
        print(f"\nError inesperado: {e}")
        print("\nTraceback completo:")
        traceback.print_exc()
        input("\nPresiona Enter para salir...")

if __name__ == "__main__":
    main()
