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
    comunidades = []
    
    for archivo in archivos:
        comunidad = extraer_comunidad_de_pdf(archivo)
        if not comunidad:
            comunidad = "Comunidad no detectada"
        comunidades.append({
            'nombre': comunidad,
            'pdf': archivo,
            'correo': 'davidvr1994@gmail.com',
            'enviar': True
        })
        resumen += f"- {os.path.basename(archivo)} | Comunidad: {comunidad}\n"
    
    messagebox.showinfo("Facturas encontradas", resumen)
    
    # Abrir la ventana de asignación de correos
    abrir_asignacion_correos(comunidades)

def abrir_asignacion_correos(comunidades):
    win = tk.Toplevel(root)
    win.title("Asignación de Correos")
    win.geometry("1000x600")
    
    # Frame principal con scroll
    main_frame = tk.Frame(win)
    main_frame.pack(fill='both', expand=True, padx=10, pady=10)
    
    # Canvas y scrollbar
    canvas = tk.Canvas(main_frame)
    scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas)
    
    # Configurar el canvas
    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )
    
    # Crear ventana en el canvas para el frame desplazable
    canvas_frame = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    
    # Configurar el frame desplazable
    def on_frame_configure(event):
        canvas.configure(scrollregion=canvas.bbox("all"))
    
    def on_canvas_configure(event):
        canvas.itemconfig(canvas_frame, width=event.width)
    
    scrollable_frame.bind("<Configure>", on_frame_configure)
    canvas.bind("<Configure>", on_canvas_configure)
    
    # Cabecera de la tabla
    header_frame = tk.Frame(scrollable_frame)
    header_frame.pack(fill='x')
    
    # Configurar el grid de la cabecera
    tk.Label(header_frame, text='Enviar', width=8, anchor='center', 
             font=('Arial', 9, 'bold')).grid(row=0, column=0, padx=2, pady=2)
    tk.Label(header_frame, text='Comunidad', anchor='w', 
             font=('Arial', 9, 'bold')).grid(row=0, column=1, padx=2, pady=2, sticky='ew')
    tk.Label(header_frame, text='Correo', anchor='w', 
             font=('Arial', 9, 'bold')).grid(row=0, column=2, padx=2, pady=2, sticky='ew')
    tk.Label(header_frame, text='Archivo', width=20, anchor='center', 
             font=('Arial', 9, 'bold')).grid(row=0, column=3, padx=2, pady=2)
    
    # Configurar pesos de columnas
    header_frame.columnconfigure(1, weight=1)
    header_frame.columnconfigure(2, weight=2)
    
    # Variables para almacenar referencias
    check_vars = []
    comunidad_entries = []
    correo_entries = []
    
    # Agregar filas
    for i, com in enumerate(comunidades, 1):
        row_frame = tk.Frame(scrollable_frame)
        row_frame.pack(fill='x', pady=2)
        
        # Checkbox para seleccionar
        var_envio = tk.BooleanVar(value=com.get('enviar', True))
        check_vars.append(var_envio)
        cb = tk.Checkbutton(row_frame, variable=var_envio)
        cb.pack(side='left', padx=4)
        
        # Entrada para el nombre de la comunidad
        var_comunidad = tk.StringVar(value=com['nombre'])
        comunidad_entries.append(var_comunidad)
        entry_comunidad = tk.Entry(row_frame, textvariable=var_comunidad, width=30)
        entry_comunidad.pack(side='left', padx=2, fill='x', expand=True)
        
        # Entrada para el correo
        var_correo = tk.StringVar(value=com.get('correo', ''))
        correo_entries.append(var_correo)
        entry_correo = tk.Entry(row_frame, textvariable=var_correo, width=40)
        entry_correo.pack(side='left', padx=2, fill='x', expand=True)
        
        # Nombre del archivo y botón para abrir
        file_frame = tk.Frame(row_frame)
        file_frame.pack(side='left', padx=2, fill='x')
        
        file_label = tk.Label(file_frame, text=os.path.basename(com['pdf']), 
                            width=20, anchor='w', relief='sunken', padx=5, pady=2)
        file_label.pack(side='left', fill='x', expand=True)
        
        btn_abrir = tk.Button(file_frame, text='Abrir', 
                            command=lambda f=com['pdf']: abrir_archivo(f))
        btn_abrir.pack(side='left', padx=2)
    
    # Frame para los botones inferiores
    btn_frame = tk.Frame(win)
    btn_frame.pack(fill='x', pady=10)
    
    def guardar_y_continuar():
        # Actualizar las comunidades con los valores editados
        for i, com in enumerate(comunidades):
            com['enviar'] = check_vars[i].get()
            com['nombre'] = comunidad_entries[i].get()
            com['correo'] = correo_entries[i].get()
        
        # Filtrar solo las comunidades seleccionadas
        comunidades_seleccionadas = [c for c in comunidades if c['enviar']]
        
        if not comunidades_seleccionadas:
            messagebox.showwarning("Advertencia", "No hay comunidades seleccionadas para enviar")
            return
            
        win.destroy()
        mostrar_confirmacion(comunidades_seleccionadas)
    
    # Botón para guardar y continuar
    btn_guardar = ttk.Button(btn_frame, text='Guardar y continuar', 
                           command=guardar_y_continuar)
    btn_guardar.pack(side='right', padx=5)
    
    # Botón para cancelar
    btn_cancelar = ttk.Button(btn_frame, text='Cancelar', 
                            command=win.destroy)
    btn_cancelar.pack(side='right', padx=5)
    
    # Configurar el canvas y scrollbar
    canvas.pack(side='left', fill='both', expand=True)
    scrollbar.pack(side='right', fill='y')
    
    # Centrar ventana
    win.update_idletasks()
    width = min(1000, win.winfo_screenwidth() - 100)
    height = min(600, win.winfo_screenheight() - 100)
    x = (win.winfo_screenwidth() // 2) - (width // 2)
    y = (win.winfo_screenheight() // 2) - (height // 2)
    win.geometry(f'{width}x{height}+{x}+{y}')
    
    win.transient(root)
    win.grab_set()
    root.wait_window(win)

def mostrar_confirmacion(comunidades):
    # Crear ventana de confirmación
    confirm_win = tk.Toplevel(root)
    confirm_win.title("Confirmar envío")
    
    # Frame principal
    main_frame = ttk.Frame(confirm_win, padding=10)
    main_frame.pack(fill='both', expand=True)
    
    # Título
    ttk.Label(main_frame, 
             text="¿Desea enviar los correos con las siguientes facturas?",
             font=('Arial', 10, 'bold')).pack(pady=5)
    
    # Crear frame con scroll para la lista de facturas
    frame = ttk.Frame(main_frame)
    frame.pack(fill='both', expand=True, pady=10)
    
    # Canvas y scrollbar
    canvas = tk.Canvas(frame, height=200)
    scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
    scrollable_frame = ttk.Frame(canvas)
    
    # Configurar el canvas
    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")
        )
    )
    
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    
    # Mostrar la lista de facturas
    for com in comunidades:
        frame_item = ttk.Frame(scrollable_frame)
        frame_item.pack(fill='x', pady=2)
        
        ttk.Label(frame_item, text=f"• {com['nombre']}").pack(side='left')
        ttk.Label(frame_item, text=os.path.basename(com['pdf']), 
                 font=('Arial', 8), foreground='gray').pack(side='right')
    
    # Empaquetar canvas y scrollbar
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    
    # Frame para botones
    btn_frame = ttk.Frame(main_frame)
    btn_frame.pack(fill='x', pady=10)
    
    # Botones
    ttk.Button(btn_frame, text="Cancelar", 
              command=confirm_win.destroy).pack(side='right', padx=5)
    
    def enviar_correos():
        confirm_win.destroy()
        # Aquí iría el código para enviar los correos
        messagebox.showinfo("Éxito", "Los correos se han enviado correctamente")
    
    ttk.Button(btn_frame, text="Enviar correos", 
              command=enviar_correos).pack(side='right', padx=5)
    
    # Centrar ventana
    confirm_win.update_idletasks()
    width = 500
    height = 400
    x = (confirm_win.winfo_screenwidth() // 2) - (width // 2)
    y = (confirm_win.winfo_screenheight() // 2) - (height // 2)
    confirm_win.geometry(f'{width}x{height}+{x}+{y}')
    confirm_win.transient(root)
    confirm_win.grab_set()
    root.wait_window(confirm_win)

def iniciar_analisis():
    """Inicia el análisis de las facturas"""
    global detener_proceso
    detener_proceso = False
    
    if 'btn_ejecutar' in globals():
        btn_ejecutar.config(state='disabled')
    if 'btn_parar' in globals():
        btn_parar.config(state='normal')
    if 'log_box' in globals():
        log_box.delete(1.0, tk.END)
        log_box.insert(tk.END, "Iniciando análisis de facturas...\n")
        log_box.see(tk.END)
    
    # Aquí iría la lógica de análisis
    # Por ahora, simulamos un análisis
    import time
    
    def analizar():
        for i in range(1, 11):
            if detener_proceso:
                break
            time.sleep(0.5)  # Simulación de trabajo
            if 'log_box' in globals():
                log_box.insert(tk.END, f"Analizando factura {i}/10...\n")
                log_box.see(tk.END)
        
        if not detener_proceso and 'log_box' in globals():
            log_box.insert(tk.END, "Análisis completado.\n")
            log_box.see(tk.END)
        
        if 'btn_ejecutar' in globals():
            btn_ejecutar.config(state='normal')
        if 'btn_parar' in globals():
            btn_parar.config(state='disabled')
    
    # Ejecutar en un hilo separado para no bloquear la interfaz
    threading.Thread(target=analizar, daemon=True).start()

def detener_analisis():
    """Detiene el análisis en curso"""
    global detener_proceso
    detener_proceso = True
    if 'btn_parar' in globals():
        btn_parar.config(state='disabled')
    if 'log_box' in globals():
        log_box.insert(tk.END, "\nAnálisis detenido por el usuario\n")
        log_box.see(tk.END)

def run_gui():
    global root, carpeta_var, detener_proceso, log_box, btn_ejecutar, btn_parar, btn_buscar
    detener_proceso = False  # Variable para controlar la detención
    
    # Configurar estilo para botones
    style = ttk.Style()
    style.configure('Accent.TButton', font=('Arial', 10, 'bold'))
    
    try:
        print('INICIANDO GUI')
        root = tk.Tk()
        root.title("Envío automatizado de facturas")
        root.geometry("800x600")
        
        # Variables globales
        carpeta_var = tk.StringVar()
        
        # --- INTERFAZ PRINCIPAL ---
        # Frame superior
        top_frame = ttk.Frame(root, padding="10")
        top_frame.pack(fill='x')
        
        # Título
        ttk.Label(top_frame, text="Envío de Facturas", 
                 font=('Arial', 16, 'bold')).pack(pady=10)
        
        # Frame principal
        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill='both', expand=True)
        
        # Selección de carpeta
        frame_carpeta = ttk.LabelFrame(main_frame, text="Carpeta de facturas", padding=10)
        frame_carpeta.pack(fill='x', pady=5)
        
        ttk.Entry(frame_carpeta, textvariable=carpeta_var, width=50).pack(side='left', fill='x', expand=True, padx=(0, 5))
        ttk.Button(frame_carpeta, text="Examinar", 
                  command=buscar_facturas_numeradas).pack(side='left')
        
        # Configuración de correo
        frame_config = ttk.LabelFrame(main_frame, text="Configuración de correo", padding=10)
        frame_config.pack(fill='x', pady=5)
        
        # Mes
        ttk.Label(frame_config, text="Mes (ej: mayo):").pack(anchor='w')
        mes_var = tk.StringVar()
        ttk.Entry(frame_config, textvariable=mes_var).pack(fill='x', pady=2)
        
        # Correo remitente
        ttk.Label(frame_config, text="Correo Gmail remitente:").pack(anchor='w')
        remitente_var = tk.StringVar()
        ttk.Entry(frame_config, textvariable=remitente_var).pack(fill='x', pady=2)
        
        # Contraseña de aplicación
        ttk.Label(frame_config, text="Contraseña de aplicación Gmail:").pack(anchor='w')
        gmail_pass_var = tk.StringVar()
        ttk.Entry(frame_config, textvariable=gmail_pass_var, show='*').pack(fill='x', pady=2)
        
        # Botón de autocompletar demo
        def autocompletar_demo():
            carpeta_var.set(r'C:\Users\David\Desktop\PROGRAMAR\ENVIADOR DE FACTURAS\FACTURAS\FACTURAS COMPLETA\05 - Mayo - 2025')
            mes_var.set('mayo')
            remitente_var.set('davidvr1994@gmail.com')
            gmail_pass_var.set('')
        
        ttk.Button(main_frame, text="Autocompletar demo", 
                  command=autocompletar_demo).pack(pady=5, anchor='w')
        
        # Área de logs
        log_frame = ttk.LabelFrame(main_frame, text="Registro de actividad", padding=5)
        log_frame.pack(fill='both', expand=True, pady=5)
        
        global log_box
        log_box = scrolledtext.ScrolledText(log_frame, height=10)
        log_box.pack(fill='both', expand=True, pady=5)
        
        # Frame para botones de acción
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill='x', pady=10)
        
        global btn_ejecutar, btn_parar, btn_buscar
        
        # Botón para buscar facturas
        btn_buscar = ttk.Button(btn_frame, text="Buscar facturas", 
                              command=buscar_facturas_numeradas)
        btn_buscar.pack(side='left', padx=5)
        
        # Botón para iniciar análisis
        btn_ejecutar = ttk.Button(btn_frame, text="Iniciar análisis", 
                                command=iniciar_analisis,
                                style='Accent.TButton')
        btn_ejecutar.pack(side='left', padx=5)
        
        # Botón para detener
        btn_parar = ttk.Button(btn_frame, text="Detener", 
                             command=detener_analisis,
                             state='disabled')
        btn_parar.pack(side='left', padx=5)
        
        # Centrar ventana
        root.update_idletasks()
        width = 800
        height = 600
        x = (root.winfo_screenwidth() // 2) - (width // 2)
        y = (root.winfo_screenheight() // 2) - (height // 2)
        root.geometry(f'{width}x{height}+{x}+{y}')
        
        # Iniciar el bucle principal
        root.mainloop()
        
    except Exception as e:
        messagebox.showerror("Error", f"Se produjo un error: {str(e)}")
        print(f"Error en la interfaz: {str(e)}")
        if 'root' in globals() and root:
            root.destroy()

if __name__ == "__main__":
    run_gui()
