print('>>> CARGADO debug_interface.py')
import sys
import os
import traceback
import tkinter as tk

# Asegurarse de que el directorio src esté en el PATH
src_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), 'src'))
if src_dir not in sys.path:
    sys.path.append(src_dir)

def main():
    try:
        print("Intentando importar la interfaz...")
        from gui.interface import run_gui
        
        print("Iniciando la interfaz gráfica...")
        
        # Crear la ventana principal de Tkinter
        root = tk.Tk()
        root.withdraw()  # Ocultar la ventana principal temporalmente
        
        # Llamar a la función run_gui
        run_gui()
        
        # Iniciar el bucle principal de Tkinter
        root.mainloop()
        
        print("Interfaz cerrada.")
        
    except Exception as e:
        print(f"Error al ejecutar la aplicación: {str(e)}")
        print("\nDetalle del error:")
        traceback.print_exc()
        input("Presiona Enter para salir...")

if __name__ == "__main__":
    main()
