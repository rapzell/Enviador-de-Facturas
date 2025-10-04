import tkinter as tk
from tkinter import ttk, messagebox
import os

def mostrar_mensaje(mensaje):
    # Reemplaza cualquier carácter no ASCII por un signo de interrogación
    mensaje_ascii = mensaje.encode('ascii', errors='replace').decode('ascii')
    print(f"[DEBUG] {mensaje_ascii}")
    messagebox.showinfo("Info", mensaje_ascii)

def main():
    print("Iniciando aplicación...")
    
    try:
        # Crear ventana principal
        root = tk.Tk()
        root.title("Envío de Facturas - Versión Simple")
        root.geometry("600x400")
        
        # Variables
        comunidades = [
            {'nombre': 'Comunidad de Prueba 1', 'correo': 'test1@example.com', 'pdf': 'factura1.pdf'},
            {'nombre': 'Comunidad de Prueba 2', 'correo': 'test2@example.com', 'pdf': 'factura2.pdf'}
        ]
        
        mostrar_mensaje("Aplicaci�n iniciada correctamente")
        
        # Función para abrir la ventana de asignación
        def abrir_asignacion():
            try:
                mostrar_mensaje("Abriendo ventana de asignación...")
                win = tk.Toplevel(root)
                win.title("Asignar Correos")
                
                # Frame principal
                main_frame = ttk.Frame(win, padding="10")
                main_frame.pack(fill=tk.BOTH, expand=True)
                
                # Título
                ttk.Label(main_frame, text="Comunidades y Correos", font=('Arial', 12, 'bold')).pack(pady=5)
                
                # Frame para la tabla
                frame_tabla = ttk.Frame(main_frame)
                frame_tabla.pack(fill=tk.BOTH, expand=True, pady=10)
                
                # Tabla
                for i, com in enumerate(comunidades):
                    # Fila
                    frame_fila = ttk.Frame(frame_tabla)
                    frame_fila.pack(fill=tk.X, pady=2)
                    
                    # Nombre (editable)
                    var_nombre = tk.StringVar(value=com['nombre'])
                    entry_nombre = ttk.Entry(frame_fila, textvariable=var_nombre, width=30)
                    entry_nombre.pack(side=tk.LEFT, padx=2)
                    
                    # Correo (editable)
                    var_correo = tk.StringVar(value=com['correo'])
                    entry_correo = ttk.Entry(frame_fila, textvariable=var_correo, width=30)
                    entry_correo.pack(side=tk.LEFT, padx=2)
                    
                    # Archivo (solo lectura)
                    ttk.Label(frame_fila, text=com['pdf'], width=20).pack(side=tk.LEFT, padx=2)
                
                # Botón de guardar
                def guardar():
                    mostrar_mensaje("Guardando cambios...")
                    win.destroy()
                
                btn_guardar = ttk.Button(main_frame, text="Guardar Cambios", command=guardar)
                btn_guardar.pack(pady=10)
                
                # Centrar ventana
                win.update_idletasks()
                width = 600
                height = 400
                x = (win.winfo_screenwidth() // 2) - (width // 2)
                y = (win.winfo_screenheight() // 2) - (height // 2)
                win.geometry(f'{width}x{height}+{x}+{y}')
                
                mostrar_mensaje("Ventana de asignación lista")
                
            except Exception as e:
                mostrar_mensaje(f"Error al abrir la ventana: {str(e)}")
        
        # Botón principal
        btn_asignar = ttk.Button(root, text="Asignar Correos", command=abrir_asignacion)
        btn_asignar.pack(pady=50)
        
        # Botón de salir
        btn_salir = ttk.Button(root, text="Salir", command=root.quit)
        btn_salir.pack(pady=10)
        
        mostrar_mensaje("Interfaz lista")
        
        # Iniciar el bucle principal
        root.mainloop()
        
    except Exception as e:
        mostrar_mensaje(f"Error crítico: {str(e)}")
        import traceback
        traceback.print_exc()
    finally:
        print("Aplicación finalizada")

if __name__ == "__main__":
    main()
