import sys
import os
import tkinter as tk
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), 'src')))

try:
    from gui.interface import run_gui
    print("Módulos importados correctamente")
    run_gui()
    print("Interfaz ejecutada")
    # Mantener la ventana abierta
    tk.mainloop()
except Exception as e:
    print(f"Error al ejecutar la aplicación: {str(e)}")
    import traceback
    traceback.print_exc()
    input("Presione Enter para salir...")
