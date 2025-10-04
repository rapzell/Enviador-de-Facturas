import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
import threading
import os
import sys
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from procesar_envios import procesar_envios

def abrir_archivo(archivo):
    try:
        os.startfile(archivo)
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo abrir el archivo: {str(e)}")

def buscar_facturas_numeradas():
    carpeta = filedialog.askdirectory(title="Seleccionar carpeta de facturas")
    if not carpeta:
        return
    carpeta_var.set(carpeta)
    
    # Buscar archivos numerados (ej: factura_001.pdf, 002.pdf, etc.)
    import glob
    archivos = glob.glob(os.path.join(carpeta, '*.pdf')) + glob.glob(os.path.join(carpeta, 'factura_*.pdf'))
    
    if not archivos:
        messagebox.showinfo("Facturas", "No se encontraron facturas numeradas en la carpeta")
        return
    
    # Extraer nombre de comunidad de cada PDF
    try:
        from src.extractor_comunidad_pdf import extraer_comunidad_de_pdf
    except ImportError:
        from extractor_comunidad_pdf import extraer_comunidad_de_pdf
    resumen = f"Se encontraron {len(archivos)} facturas:\n"
    comunidades_pdf = {}
    for i, archivo in enumerate(sorted(archivos), 1):
        comunidad = extraer_comunidad_de_pdf(archivo)
        comunidades_pdf[archivo] = comunidad
        resumen += f"{i}. {os.path.basename(archivo)} | Comunidad: {comunidad if comunidad else '[NO DETECTADA]'}\n"
    messagebox.showinfo("Facturas encontradas", resumen)

def run_gui():
    import traceback
    from tkinter import messagebox
    import threading
    
    # Variables globales para controlar el hilo
    procesando = False
    detener_proceso = False
    
    def detener_analisis():
        nonlocal detener_proceso
        detener_proceso = True
        btn_parar.config(state='disabled')
        log_box.insert(tk.END, "\nDeteniendo análisis...\n")
    
    def abrir_asignacion_correos(comunidades):
        win = tk.Toplevel(root)
        win.title("Asignación de Correos")
        win.geometry("1000x600")
        
        frame_tabla = tk.Frame(win)
        frame_tabla.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Cabecera y filas en un solo Frame usando grid
        tabla = tk.Frame(frame_tabla)
        tabla.pack(fill='both', expand=True)
        
        # Configurar el grid para que las columnas se expandan correctamente
        tabla.columnconfigure(1, weight=1)  # Columna del nombre
        tabla.columnconfigure(2, weight=1)  # Columna del correo
        
        # Encabezados
        tk.Label(tabla, text='Enviar', width=8, anchor='center', font=('Arial', 9, 'bold')).grid(row=0, column=0, padx=2, pady=2, sticky='ew')
        tk.Label(tabla, text='Comunidad', width=32, anchor='w', font=('Arial', 9, 'bold')).grid(row=0, column=1, padx=2, pady=2, sticky='ew')
        tk.Label(tabla, text='Correo', width=36, anchor='w', font=('Arial', 9, 'bold')).grid(row=0, column=2, padx=2, pady=2, sticky='ew')
        tk.Label(tabla, text='Archivo', width=14, anchor='center', font=('Arial', 9, 'bold')).grid(row=0, column=3, padx=2, pady=2, sticky='ew')
        tk.Label(tabla, text='', width=8).grid(row=0, column=4, sticky='ew')

        # Variables para checkboxes, nombres y correos
        check_vars = []
        entry_vars = []
        nombre_vars = []
        
        # Agregar filas
        for i, com in enumerate(comunidades):
            var_envio = tk.BooleanVar(value=True)
            var_correo = tk.StringVar(value=com.get('correo', 'davidvr1994@gmail.com'))
            var_nombre = tk.StringVar(value=com.get('nombre', ''))
            
            check_vars.append(var_envio)
            entry_vars.append(var_correo)
            nombre_vars.append(var_nombre)
            
            # Checkbox para seleccionar
            cb = tk.Checkbutton(tabla, variable=var_envio)
            cb.grid(row=i+1, column=0, padx=4, pady=2, sticky='w')
            
            # Campo editable para el nombre de la comunidad
            entry_nombre = tk.Entry(tabla, textvariable=var_nombre, state='normal')
            entry_nombre.grid(row=i+1, column=1, padx=2, pady=2, sticky='we')
            
            # Campo de correo editable
            entry_correo = tk.Entry(tabla, textvariable=var_correo, state='normal')
            entry_correo.grid(row=i+1, column=2, padx=2, pady=2, sticky='we')
            
            # Nombre del archivo PDF
            lbl_archivo = tk.Label(tabla, text=os.path.basename(com['pdf']), width=14, anchor='center')
            lbl_archivo.grid(row=i+1, column=3, padx=2, pady=2)
            
            # Botón para abrir PDF
            btn_abrir = tk.Button(tabla, text='Abrir', 
                                command=lambda f=com['pdf']: abrir_archivo(f) if os.path.exists(f) else None)
            btn_abrir.grid(row=i+1, column=4, padx=2, pady=2, sticky='ew')
            
            # Deshabilitar botón si el PDF no existe
            if not os.path.exists(com['pdf']):
                btn_abrir.config(state='disabled')

        # Botón inferior
        btn_guardar = ttk.Button(win, text='Guardar y continuar', command=guardar_y_continuar)
        btn_guardar.pack(pady=10)

        def guardar_y_continuar():
            comunidades_actualizadas = []
            for i, com in enumerate(comunidades):
                # Actualizar siempre el nombre y correo, esté o no seleccionado
                com['nombre'] = nombre_vars[i].get()
                com['correo'] = entry_vars[i].get()
                
                # Solo agregar a la lista de envío si está seleccionado
                if check_vars[i].get() and entry_vars[i].get():
                    comunidades_actualizadas.append(com)
            
            if not comunidades_actualizadas:
                messagebox.showwarning("Advertencia", "No hay comunidades seleccionadas para enviar")
                return
                
            win.destroy()
            mostrar_confirmacion(comunidades_actualizadas)

        # Centrar ventana
        win.update_idletasks()
        width = win.winfo_width()
        height = win.winfo_height()
        x = (win.winfo_screenwidth() // 2) - (width // 2)
        y = (win.winfo_screenheight() // 2) - (height // 2)
        win.geometry(f'{width}x{height}+{x}+{y}')
        win.transient(root)
        win.grab_set()
        root.wait_window(win)
    
    # Resto del código de run_gui...
    
    try:
        print('INICIANDO GUI')
        root = tk.Tk()
        root.title("Envío automatizado de facturas")
        root.geometry("800x600")
        
        # --- INICIALIZACIÓN DE WIDGETS PRINCIPALES ---
        top_frame = tk.Frame(root)
        top_frame.pack(fill='x', padx=5, pady=5)
        btn_buscar = tk.Button(top_frame, text=" Buscar facturas", command=buscar_facturas_numeradas)
        btn_buscar.pack(side='right', padx=5)

        main_frame = tk.Frame(root)
        main_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Resto de la interfaz...
        
    except Exception as e:
        messagebox.showerror("Error", f"Error inesperado: {str(e)}\n\n{traceback.format_exc()}")
        raise

if __name__ == "__main__":
    run_gui()
