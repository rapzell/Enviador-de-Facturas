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
    # Puedes guardar comunidades_pdf en variable global o pasarla al flujo de asignación

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
        tk.Label(tabla, text='Enviar', width=8, anchor='center', font=('Arial', 9, 'bold')).grid(row=0, column=0, padx=2, pady=2)
        tk.Label(tabla, text='Comunidad', width=32, anchor='w', font=('Arial', 9, 'bold')).grid(row=0, column=1, padx=2, pady=2)
        tk.Label(tabla, text='Correo', width=36, anchor='w', font=('Arial', 9, 'bold')).grid(row=0, column=2, padx=2, pady=2)
        tk.Label(tabla, text='Archivo', width=14, anchor='center', font=('Arial', 9, 'bold')).grid(row=0, column=3, padx=2, pady=2)
        tk.Label(tabla, text='', width=8).grid(row=0, column=4)

        # Variables para checkboxes y entries
        check_vars = []
        entry_vars = []
        
        # Agregar filas
        for i, com in enumerate(comunidades):
            var_envio = tk.BooleanVar(value=True)
            var_correo = tk.StringVar(value='davidvr1994@gmail.com')
            check_vars.append(var_envio)
            entry_vars.append(var_correo)
            cb = tk.Checkbutton(tabla, variable=var_envio)
            cb.grid(row=i+1, column=0, padx=4, sticky='w')
            tk.Label(tabla, text=com['nombre'], width=32, anchor='w').grid(row=i+1, column=1, sticky='w')
            entry = tk.Entry(tabla, textvariable=var_correo, state='normal')
            entry.grid(row=i+1, column=2, padx=2, sticky='we')
            tabla.grid_columnconfigure(2, weight=1)
            tk.Label(tabla, text=os.path.basename(com['pdf']), width=14, anchor='center').grid(row=i+1, column=3)
            btn = tk.Button(tabla, text='Abrir', command=lambda f=com['pdf']: abrir_archivo(f))
            btn.grid(row=i+1, column=4, padx=2)

        # Botón inferior
        btn_guardar = ttk.Button(win, text='Guardar y continuar', command=lambda: guardar_y_continuar())
        btn_guardar.pack(pady=10)

        def guardar_y_continuar():
            comunidades_actualizadas = []
            for i, com in enumerate(comunidades):
                if check_vars[i].get() and entry_vars[i].get():
                    com['correo'] = entry_vars[i].get()
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
        frame_img = tk.Frame(main_frame)
        frame_img.pack(fill='x')
        carpeta_var = tk.StringVar()
        tk.Entry(frame_img, textvariable=carpeta_var, width=50).pack(side='left', padx=2, expand=True, fill='x')
        def browse_folder():
            d = filedialog.askdirectory()
            if d:
                carpeta_var.set(d)
        tk.Button(frame_img, text="Examinar", command=browse_folder).pack(side='left', padx=5)
        tk.Label(main_frame, text="Mes (ej: mayo):").pack(anchor='w')
        mes_var = tk.StringVar()
        tk.Entry(main_frame, textvariable=mes_var).pack(fill='x')
        tk.Label(main_frame, text="Correo Gmail remitente:").pack(anchor='w')
        remitente_var = tk.StringVar()
        tk.Entry(main_frame, textvariable=remitente_var).pack(fill='x')
        tk.Label(main_frame, text="Contraseña de aplicación Gmail:").pack(anchor='w')
        gmail_pass_var = tk.StringVar()
        tk.Entry(main_frame, textvariable=gmail_pass_var, show='*').pack(fill='x')
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill='x', pady=10)
        def autocompletar_demo():
            carpeta_var.set(r'C:\Users\David\Desktop\PROGRAMAR\ENVIADOR DE FACTURAS\FACTURAS\FACTURAS COMPLETA\05 - Mayo - 2025')
            mes_var.set('mayo')
            remitente_var.set('davidvr1994@gmail.com')
            gmail_pass_var.set('')
        tk.Button(button_frame, text="Autocompletar demo", command=autocompletar_demo, bg='#2196F3', fg='white').pack(side='left', padx=5)
        
        def procesar_en_hilo():
            nonlocal procesando, detener_proceso
            
            # Resetear banderas
            procesando = True
            detener_proceso = False
            
            try:
                # Obtener valores de la interfaz
                carpeta = carpeta_var.get()
                mes = mes_var.get()
                remitente = remitente_var.get()
                gmail_pass = gmail_pass_var.get()
                
                # Verificar campos obligatorios
                if not carpeta or not mes or not remitente or not gmail_pass:
                    messagebox.showerror("Error", "Por favor complete todos los campos")
                    return
                
                # Actualizar interfaz
                root.after(0, lambda: [
                    btn_parar.config(state='normal'),
                    btn_ejecutar.config(state='disabled')
                ])
                
                log_box.delete(1.0, tk.END)
                log_box.insert(tk.END, "Iniciando análisis de facturas...\n")
                root.update_idletasks()
                
                # 1. Buscar facturas en la carpeta
                log_box.insert(tk.END, "Buscando facturas en la carpeta...\n")
                root.update_idletasks()
                
                import glob
                archivos = glob.glob(os.path.join(carpeta, '*.pdf')) + glob.glob(os.path.join(carpeta, 'factura_*.pdf'))
                
                if detener_proceso:
                    return
                    
                if not archivos:
                    messagebox.showerror("Error", "No se encontraron archivos PDF en la carpeta especificada")
                    return
                    
                log_box.insert(tk.END, f"Se encontraron {len(archivos)} facturas.\n")
                
                # 2. Extraer información de cada factura
                log_box.insert(tk.END, "\nExtrayendo información de las facturas...\n")
                root.update_idletasks()
                
                try:
                    from src.extractor_comunidad_pdf import extraer_comunidad_de_pdf
                except ImportError:
                    from extractor_comunidad_pdf import extraer_comunidad_de_pdf
                    
                comunidades = []
                for archivo in archivos:
                    if detener_proceso:
                        log_box.insert(tk.END, "\nAnálisis detenido por el usuario\n")
                        break
                        
                    comunidad = extraer_comunidad_de_pdf(archivo)
                    if comunidad:
                        comunidades.append({
                            'nombre': comunidad,
                            'correo': '',  # Se puede extraer del PDF si está disponible
                            'pdf': archivo,
                            'mes': mes
                        })
                        log_box.insert(tk.END, f"Procesada: {os.path.basename(archivo)} - {comunidad}\n")
                    else:
                        log_box.insert(tk.END, f"[ADVERTENCIA] No se pudo extraer comunidad de: {os.path.basename(archivo)}\n")
                    root.update_idletasks()
                    
                # Si se detuvo el proceso, abrir ventana con lo que se pudo analizar
                if detener_proceso and comunidades:
                    root.after(0, lambda: abrir_asignacion_correos(comunidades))
                    return
                    
                if not comunidades:
                    messagebox.showerror("Error", "No se pudo extraer información de ninguna factura")
                    return
                    
                # Si llegamos aquí, no se detuvo el proceso
                # 3. Configurar SMTP y procesar envíos
                root.after(0, lambda: abrir_asignacion_correos(comunidades))
                    
            except Exception as e:
                log_box.insert(tk.END, f"\n[ERROR] {str(e)}\n")
                messagebox.showerror("Error en el proceso", str(e))
            finally:
                # Restaurar estado de la interfaz
                procesando = False
                detener_proceso = False
                root.after(0, lambda: [
                    btn_ejecutar.config(state='normal'),
                    btn_parar.config(state='disabled')
                ])
            
        # Botón de ejecutar
        btn_ejecutar = tk.Button(button_frame, text="[EJECUTAR] Envío de facturas", 
                               command=lambda: threading.Thread(target=procesar_en_hilo, daemon=True).start(), 
                               bg='#4CAF50', fg='white')
        btn_ejecutar.pack(side='left', padx=5)
        
        # Botón de parar (inicialmente deshabilitado)
        btn_parar = tk.Button(button_frame, text="[PARAR] Análisis", command=detener_analisis, 
                            bg='#f44336', fg='white', state='disabled')
        btn_parar.pack(side='left', padx=5)
        
        progress_frame = tk.Frame(main_frame)
        progress = ttk.Progressbar(progress_frame, mode='indeterminate', length=200)
        progress_label = tk.Label(progress_frame, text="Procesando, por favor espera...")
        log_box = scrolledtext.ScrolledText(main_frame, height=9)
        log_box.pack(fill='both', padx=2, pady=2, expand=True)
        # --- FIN INICIALIZACIÓN DE WIDGETS ---
        # Puedes restaurar aquí tus funciones internas y lógica, asegurando que todo esté dentro del try
        root.mainloop()
    except Exception as e:
        print('ERROR EN TKINTER:', e)
        print(traceback.format_exc())
        try:
            messagebox.showerror('Error en GUI', str(e) + '\n' + traceback.format_exc())
        except Exception:
            pass

    def abrir_asignacion_correos(comunidades):
        win = tk.Toplevel(root)
        win.title("Asignación de Correos")
        win.geometry("1000x600")
        
        frame_tabla = tk.Frame(win)
        frame_tabla.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Cabecera y filas en un solo Frame usando grid
        tabla = tk.Frame(frame_tabla)
        tabla.pack(fill='both', expand=True)
        tk.Label(tabla, text='Enviar', width=8, anchor='center', font=('Arial', 9, 'bold')).grid(row=0, column=0, padx=2, pady=2)
        tk.Label(tabla, text='Comunidad', width=32, anchor='w', font=('Arial', 9, 'bold')).grid(row=0, column=1, padx=2, pady=2)
        tk.Label(tabla, text='Correo', width=36, anchor='w', font=('Arial', 9, 'bold')).grid(row=0, column=2, padx=2, pady=2)
        tk.Label(tabla, text='Archivo', width=14, anchor='center', font=('Arial', 9, 'bold')).grid(row=0, column=3, padx=2, pady=2)
        tk.Label(tabla, text='', width=8).grid(row=0, column=4)

        # Variables para checkboxes y entries
        check_vars = []
        entry_vars = []
        
        # Agregar filas
        for i, com in enumerate(comunidades):
            var_envio = tk.BooleanVar(value=True)
            var_correo = tk.StringVar(value='davidvr1994@gmail.com')
            check_vars.append(var_envio)
            entry_vars.append(var_correo)
            cb = tk.Checkbutton(tabla, variable=var_envio)
            cb.grid(row=i+1, column=0, padx=4, sticky='w')
            tk.Label(tabla, text=com['nombre'], width=32, anchor='w').grid(row=i+1, column=1, sticky='w')
            entry = tk.Entry(tabla, textvariable=var_correo, state='normal')
            entry.grid(row=i+1, column=2, padx=2, sticky='we')
            tabla.grid_columnconfigure(2, weight=1)
            tk.Label(tabla, text=os.path.basename(com['pdf']), width=14, anchor='center').grid(row=i+1, column=3)
            btn = tk.Button(tabla, text='Abrir', command=lambda f=com['pdf']: abrir_archivo(f))
            btn.grid(row=i+1, column=4, padx=2)

        # Botón inferior
        btn_guardar = ttk.Button(win, text='Guardar y continuar', command=lambda: guardar_y_continuar())
        btn_guardar.pack(pady=10)

        def guardar_y_continuar():
            comunidades_actualizadas = []
            for i, com in enumerate(comunidades):
                if check_vars[i].get() and entry_vars[i].get():
                    com['correo'] = entry_vars[i].get()
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

