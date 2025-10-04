import os
import sys
import subprocess

# Ruta al script principal
script_path = os.path.join(os.path.dirname(__file__), 'src', 'gui', 'interface.py')

# Ejecutar el script con la consola visible
subprocess.Popen([sys.executable, script_path], creationflags=subprocess.CREATE_NEW_CONSOLE)
