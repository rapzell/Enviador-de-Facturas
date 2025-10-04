import tkinter as tk
from tkinter import ttk, messagebox
import os

def main():
    root = tk.Tk()
    root.title("Envío de Facturas")
    
    # Variables
    comunidades = [
        {'nombre': 'Comunidad 1', 'correo': 'ejemplo1@test.com', 'pdf': 'factura1.pdf'},
        {'nombre': 'Comunidad 2', 'correo': 'ejemplo2@test.com', 'pdf': 'factura2.pdf'}
    ]
    
    # Función para abrir la ventana de asignación de correos
    def abrir_asignacion_correos():
        win = tk.Toplevel(root)
        win.title("Asignar Correos")
        
        # Frame principal con scrollbar
        main_frame = ttk.Frame(win)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Canvas y scrollbar
        canvas = tk.Canvas(main_frame)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Variables para los campos editables
        check_vars = []
        entry_vars = []
        nombre_vars = []
        
        # Tabla de comunidades
        tabla = ttk.Frame(scrollable_frame)
        tabla.pack(fill='x', padx=5, pady=5)
        
        # Encabezados
        ttk.Label(tabla, text='Enviar', width=6).grid(row=0, column=0, padx=2, pady=2)
        ttk.Label(tabla, text='Comunidad', width=32).grid(row=0, column=1, padx=2, pady=2)
        ttk.Label(tabla, text='Correo', width=36).grid(row=0, column=2, padx=2, pady=2)
        ttk.Label(tabla, text='Archivo', width=14).grid(row=0, column=3, padx=2, pady=2)
        
        # Función para guardar y continuar
        def guardar_y_continuar():
            comunidades_actualizadas = []
            for i, com in enumerate(comunidades):
                # Actualizar nombre y correo
                com['nombre'] = nombre_vars[i].get()
                com['correo'] = entry_vars[i].get()
                
                # Agregar a la lista si está marcado para enviar
                if check_vars[i].get() and entry_vars[i].get():
                    comunidades_actualizadas.append(com)
            
            if not comunidades_actualizadas:
                messagebox.showwarning("Advertencia", "No hay comunidades seleccionadas para enviar")
                return
                
            # Mostrar confirmación
            mensaje = "¿Desea continuar con el envío a las siguientes comunidades?\n\n"
            mensaje += "\n".join([f"- {c['nombre']} ({c['correo']})" for c in comunidades_actualizadas])
            
            if messagebox.askyesno("Confirmar envío", mensaje):
                print("Iniciando envío de correos...")
                # Aquí iría la lógica de envío
                messagebox.showinfo("Éxito", "Los correos se han enviado correctamente")
            
            win.destroy()
        
        # Agregar filas de comunidades
        for i, com in enumerate(comunidades):
            # Variables
            var_envio = tk.BooleanVar(value=True)
            var_correo = tk.StringVar(value=com['correo'])
            var_nombre = tk.StringVar(value=com['nombre'])
            
            # Guardar referencias
            check_vars.append(var_envio)
            entry_vars.append(var_correo)
            nombre_vars.append(var_nombre)
            
            # Fila en la tabla
            row = i + 1
            
            # Checkbox para seleccionar
            cb = ttk.Checkbutton(tabla, variable=var_envio)
            cb.grid(row=row, column=0, padx=4, pady=2, sticky='w')
            
            # Campo de nombre de comunidad (editable)
            entry_nombre = ttk.Entry(tabla, textvariable=var_nombre, width=35)
            entry_nombre.grid(row=row, column=1, padx=2, pady=2, sticky='we')
            
            # Campo de correo (editable)
            entry_correo = ttk.Entry(tabla, textvariable=var_correo, width=40)
            entry_correo.grid(row=row, column=2, padx=2, pady=2, sticky='we')
            
            # Nombre del archivo (solo lectura)
            ttk.Label(tabla, text=com['pdf'], width=20).grid(row=row, column=3, padx=2, pady=2)
        
        # Configurar el grid
        tabla.grid_columnconfigure(2, weight=1)
        
        # Botón de guardar
        btn_frame = ttk.Frame(win)
        btn_frame.pack(fill='x', padx=10, pady=10)
        
        btn_guardar = ttk.Button(btn_frame, text="Guardar y continuar", command=guardar_y_continuar)
        btn_guardar.pack(side='right')
        
        # Configurar el scroll
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Ajustar tamaño de la ventana
        win.geometry("800x400")
        win.minsize(600, 300)
    
    # Botón para abrir la asignación de correos
    btn_abrir = ttk.Button(root, text="Asignar correos", command=abrir_asignacion_correos)
    btn_abrir.pack(pady=20)
    
    # Botón para probar el autocompletado
    def autocompletar_demo():
        # Aquí iría la lógica de autocompletado
        messagebox.showinfo("Demo", "Función de autocompletado demo")
    
    btn_demo = ttk.Button(root, text="Autocompletar demo", command=autocompletar_demo)
    btn_demo.pack(pady=10)
    
    root.mainloop()

if __name__ == "__main__":
    main()
