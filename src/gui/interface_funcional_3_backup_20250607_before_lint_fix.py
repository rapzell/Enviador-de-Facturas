import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
import threading
import os
import sys
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from procesar_envios import procesar_envios, enviar_correo # Actualizado para incluir enviar_correo
import pandas as pd # Para leer Excel
# openpyxl será usado por pandas implícitamente para .xlsx, asegúrate de que esté instalado.
import re


# Ruta al archivo Excel de mapeo por defecto
RUTA_EXCEL_POR_DEFECTO = r"C:\Users\David\Desktop\PROGRAMAR\ENVIADOR DE FACTURAS\ESTRUCTURA\correos_y_comunidades.xlsx"
# lista_mapeos_global se elimina ya que los mapeos serán un diccionario cargado directamente.

def normalizar_nombre_comunidad(nombre):
    """Normalización básica: minúsculas, elimina 'urb' y 'blq' (y variantes), luego no alfanuméricos. CONSERVA NÚMEROS."""
    if not nombre:
        return ""
    nombre_str = str(nombre).lower()

    # Eliminar 'blq' seguido opcionalmente por un punto y luego opcionalmente por espacios.
    nombre_str = re.sub(r'blq\.?' + r'\s*', '', nombre_str, flags=re.IGNORECASE)
    
    # Eliminar 'urb' seguido opcionalmente por un punto y luego opcionalmente por espacios.
    nombre_str = re.sub(r'urb\.?' + r'\s*', '', nombre_str, flags=re.IGNORECASE)

    # Eliminar la mayoría de los caracteres no alfanuméricos restantes,
    # pero conservar letras (incluyendo ñ, acentuadas) y números.
    nombre_limpio = re.sub(r'[^a-z0-9ñáéíóúü]', '', nombre_str, flags=re.IGNORECASE)
    return nombre_limpio

def cargar_mapeo_desde_excel(log_func):
    """Carga el mapeo de comunidades y correos desde el archivo Excel especificado."""
    mapeo_para_procesamiento = {}
    if not os.path.exists(RUTA_EXCEL_POR_DEFECTO):
        log_func(f"ERROR: Archivo Excel de mapeo no encontrado en {RUTA_EXCEL_POR_DEFECTO}. No se cargarán mapeos.")
        return mapeo_para_procesamiento

    try:
        df = pd.read_excel(RUTA_EXCEL_POR_DEFECTO, header=None) # Asumimos que no hay encabezados
        log_func(f"Cargando mapeo de correos desde Excel: {RUTA_EXCEL_POR_DEFECTO}")

        for index, row in df.iterrows():
            if row.empty or pd.isna(row.iloc[0]): # Omitir filas vacías o sin correo
                continue

            correo = str(row.iloc[0]).strip()
            if not correo: # Omitir si el correo está vacío después de limpiar
                continue

            # Iterar sobre las celdas subsecuentes en la fila para nombres de comunidad
            for i in range(1, len(row)):
                if pd.isna(row.iloc[i]): # Detenerse si la celda es NaN (vacía)
                    break 
                
                nombre_comunidad_excel = str(row.iloc[i]).strip()
                if not nombre_comunidad_excel: # Omitir nombres de comunidad vacíos
                    continue

                nombre_normalizado = normalizar_nombre_comunidad(nombre_comunidad_excel)
                if nombre_normalizado:
                    if nombre_normalizado in mapeo_para_procesamiento and mapeo_para_procesamiento[nombre_normalizado] != correo:
                        log_func(f"ADVERTENCIA: Comunidad duplicada en Excel '{nombre_comunidad_excel}' (normalizada: '{nombre_normalizado}') "
                                 f"mapeada previamente a '{mapeo_para_procesamiento[nombre_normalizado]}' y ahora a '{correo}'. "
                                 f"Se utilizará la última asignación encontrada: '{correo}'.")
                    mapeo_para_procesamiento[nombre_normalizado] = correo
        
        log_func(f"Mapeo de correos cargado desde Excel: {len(mapeo_para_procesamiento)} comunidades mapeadas.")
        
    except FileNotFoundError:
        log_func(f"ERROR CRÍTICO: Archivo Excel de mapeo NO ENCONTRADO en {RUTA_EXCEL_POR_DEFECTO}. Verifique la ruta y el nombre del archivo.")
        messagebox.showerror("Error de Archivo", f"No se encontró el archivo de mapeo Excel en:\n{RUTA_EXCEL_POR_DEFECTO}\n\nPor favor, asegúrese de que el archivo existe en esa ubicación.")
    except ImportError:
        log_func("ERROR: La librería pandas o su dependencia openpyxl no está instalada. Instálela con 'pip install pandas openpyxl'.")
        messagebox.showerror("Error de Dependencia", "La librería pandas o openpyxl no está instalada.\nPor favor, instálela para leer archivos Excel.")
    except Exception as e:
        log_func(f"Error inesperado al cargar el mapeo desde Excel: {e}")
        import traceback
        log_func(traceback.format_exc())
        messagebox.showerror("Error al Leer Excel", f"Ocurrió un error al leer el archivo Excel:\n{e}")
    return mapeo_para_procesamiento

def abrir_archivo(archivo):
    try:
        os.startfile(archivo)
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo abrir el archivo: {str(e)}")

def buscar_facturas_numeradas():
    """Función para buscar facturas PDF y mostrar un resumen de las comunidades detectadas"""
    carpeta = filedialog.askdirectory(title="Seleccionar carpeta de facturas")
    if not carpeta:
        return
    carpeta_var.set(carpeta)
    
    import glob
    # Buscar todos los PDFs en la carpeta seleccionada
    archivos = glob.glob(os.path.join(carpeta, '*.pdf')) + glob.glob(os.path.join(carpeta, 'factura_*.pdf'))
    
    if not archivos:
        messagebox.showinfo("Facturas", "No se encontraron facturas PDF en la carpeta seleccionada")
        return
    
    # Importar la función de extracción de comunidad
    try:
        from src.extractor_comunidad_pdf import extraer_comunidad_de_pdf
    except ImportError:
        from extractor_comunidad_pdf import extraer_comunidad_de_pdf
    
    # Construir resumen de las facturas encontradas
    resumen = f"Se encontraron {len(archivos)} facturas:\n"
    for i, archivo in enumerate(sorted(archivos), 1):
        comunidad_original = extraer_comunidad_de_pdf(archivo)
        comunidad_normalizada = normalizar_nombre_comunidad(comunidad_original) if comunidad_original else ''
        
        # Mostrar tanto el nombre original como el normalizado
        resumen += f"{i}. {os.path.basename(archivo)} | Comunidad: {comunidad_original if comunidad_original else '[NO DETECTADA]'}"
        if comunidad_original and comunidad_normalizada:
            resumen += f" (Normalizada: {comunidad_normalizada})"
        resumen += "\n"
    
    # Mostrar resumen en un cuadro de diálogo
    messagebox.showinfo("Facturas encontradas", resumen)

def run_gui():
    """Función principal que maneja la interfaz gráfica y toda la lógica del programa de envío de facturas."""
    try:
        import traceback
        print('INICIANDO RUN_GUI')

        root = tk.Tk()
        root.title("Envío de Facturas a Comunidades")
        root.geometry("600x600")
        root.minsize(500, 500)

        # Declarar todas las variables globales al principio de la función
        # Estas variables serán accedidas por múltiples funciones anidadas
        global carpeta_var, remitente_var, gmail_pass_var

        # Variables para control de estado
        procesando = False
        detener_proceso_solicitado = False
        
        # Inicializar el diccionario de mapeo
        mapeo_comunidades_actual = {}
        
        # Crear las variables Tkinter
        carpeta_var = tk.StringVar(root)
        remitente_var = tk.StringVar(root)
        gmail_pass_var = tk.StringVar(root)
        
        # Cargar mapeo de comunidades desde Excel al iniciar
        try:
            # Se utiliza una función anónima para la llamada a cargar_mapeo_desde_excel
            # ya que el log_box aún no existe
            mapeo_comunidades_actual = cargar_mapeo_desde_excel(lambda msg: print(msg))
        except Exception as e_excel:
            print(f"ERROR CRÍTICO: No se pudo cargar el mapeo desde Excel: {e_excel}")
            messagebox.showerror("Error de Mapeo", f"No se pudo cargar el mapeo desde Excel: {e_excel}")

        def detener_proceso_actual():
            nonlocal detener_proceso_solicitado, procesando
            if procesando:
                detener_proceso_solicitado = True
                messagebox.showinfo("Info", "Se ha solicitado detener el proceso. Espere un momento...")
            else:
                messagebox.showinfo("Info", "No hay ningún proceso en ejecución para detener.")
            
        # Contenedor principal
        main_frame = tk.Frame(root, padx=10, pady=10)
        main_frame.pack(fill='both', expand=True)
        
        def log_func_hilo(mensaje):
            """Función para manejar logs desde hilos secundarios"""
            log_box.config(state='normal')
            log_box.insert(tk.END, f"{mensaje}\n")
            log_box.see(tk.END)
            log_box.config(state='disabled')
            root.update_idletasks()
        if not mapeo_comunidades_actual:
            log_func_hilo("ADVERTENCIA: No se pudo cargar el mapeo desde Excel o está vacío.")
        else:
            log_func_hilo(f"Mapeo cargado correctamente. {len(mapeo_comunidades_actual)} comunidades mapeadas.")
    except Exception as e_excel:
        log_func_hilo(f"ERROR al cargar mapeo desde Excel: {str(e_excel)}")
        # No usar messagebox aquí porque root aún no está completamente configurado

    def _logica_envio_final(comunidades_para_envio_final, remitente_email, remitente_password, ventana_confirmacion):
        """Función interna para manejar el envío de correos una vez confirmado"""
        nonlocal detener_proceso_solicitado, procesando
        # Usamos los parámetros pasados explícitamente en lugar de variables externas
        ventana_confirmacion.destroy()
        log_func_hilo("Confirmación final recibida. Iniciando envío de correos...")
        
        # Estadísticas de envío
        envios_exitosos = 0
        envios_fallidos = 0

        for idx, item in enumerate(comunidades_para_envio_final):
            if detener_proceso_solicitado:
                log_func_hilo("Envío de correos cancelado por el usuario.")
                break
            
            # Preparar datos para el envío
            nombre_comunidad = item['nombre']
            correo_destino = item['correo']
            ruta_pdf = item['pdf']
            asunto = f"Factura: {nombre_comunidad}"
            cuerpo = f"Estimado cliente,\n\nAdjuntamos su factura para {nombre_comunidad}.\n\nSaludos cordiales."
            
            log_func_hilo(f"Enviando ({idx+1}/{len(comunidades_para_envio_final)}): {nombre_comunidad} a {correo_destino}")
            try:
                # Usar la función de envío de correo
                enviar_correo(remitente_email, remitente_password, correo_destino, asunto, cuerpo, ruta_pdf)
                log_func_hilo(f"  Éxito: Factura enviada a {correo_destino} para {nombre_comunidad}")
                envios_exitosos += 1
            except Exception as e_envio:
                log_func_hilo(f"  ERROR al enviar a {correo_destino} para {nombre_comunidad}: {e_envio}")
                envios_fallidos += 1
                
        # Resumen final
        mensaje_final = f"\nProceso de envío finalizado.\nFacturas procesadas: {len(comunidades_para_envio_final)}\nEnviadas con éxito: {envios_exitosos}\nFallidas: {envios_fallidos}"
        log_func_hilo(mensaje_final)
        messagebox.showinfo("Resultado del Envío", mensaje_final)

    def mostrar_dialogo_confirmacion_final(comunidades_a_enviar, datos_remitente):
        """Muestra un diálogo de confirmación final con la lista de correos a enviar"""
        global remitente_var, gmail_pass_var, carpeta_var
        # Usamos datos recibidos explícitamente y/o las variables globales
        
        # Validaciones previas
        if not comunidades_a_enviar:
            messagebox.showinfo("Sin datos", "No hay facturas seleccionadas para enviar.")
            return
        
        # Usamos los valores que recibimos como parámetro
        remitente = datos_remitente['remitente'].strip() if datos_remitente.get('remitente') else ''
        password = datos_remitente.get('password', '')
        
        if not remitente or not password:
            messagebox.showerror("Credenciales incompletas", 
                               "Debe proporcionar un correo remitente válido y la contraseña de aplicación de Gmail.")
            return
        
        # Crear ventana de confirmación
        confirm_win = tk.Toplevel(root)
        confirm_win.title("Confirmación Final de Envío")
        confirm_win.geometry("650x500")
        
        tk.Label(confirm_win, text="Se enviarán correos a las siguientes comunidades:", 
               font=('Arial', 12, 'bold')).pack(pady=10)
        
        # Área de desplazamiento para mostrar la lista de facturas
        text_frame = tk.Frame(confirm_win)
        text_frame.pack(pady=5, padx=10, fill='both', expand=True)
        
        list_text = scrolledtext.ScrolledText(text_frame, height=18, width=70, wrap=tk.WORD)
        list_text.pack(side='left', fill='both', expand=True)
        
        scrollbar = ttk.Scrollbar(text_frame, orient='vertical', command=list_text.yview)
        scrollbar.pack(side='right', fill='y')
        list_text['yscrollcommand'] = scrollbar.set
        
        # Mostrar datos de las facturas a enviar
        list_text.config(state='normal')
        for i, com in enumerate(comunidades_a_enviar, 1):
            list_text.insert(tk.END, f"{i}. Comunidad: {com['nombre']}\n")
            list_text.insert(tk.END, f"   Correo: {com['correo']}\n")
            list_text.insert(tk.END, f"   Archivo: {os.path.basename(com['pdf'])}\n")
            list_text.insert(tk.END, "----------------------------------------\n")
        list_text.config(state='disabled')
        
        # Botones de acción
        btn_frame_confirm = tk.Frame(confirm_win)
        btn_frame_confirm.pack(pady=10)

        ttk.Button(btn_frame_confirm, text="Enviar Correos", 
                 command=lambda: _logica_envio_final(
                     comunidades_a_enviar, remitente, password, confirm_win)
                 ).pack(side='left', padx=10)
                 
        ttk.Button(btn_frame_confirm, text="Cancelar", 
                 command=confirm_win.destroy).pack(side='right', padx=10)

        # Configuración final de la ventana
        confirm_win.transient(root)
        confirm_win.grab_set()
        root.wait_window(confirm_win)
    
    def abrir_ventana_asignacion_correos(lista_comunidades_procesadas):
        """Abre una ventana para revisar y ajustar la asignación de correos a comunidades"""
        nonlocal mapeo_comunidades_actual
        global remitente_var, gmail_pass_var, carpeta_var
        
        # Validar si hay comunidades para procesar
        if not lista_comunidades_procesadas:
            messagebox.showinfo("Sin datos", "No hay comunidades para procesar.")
            return
            
        # Crear la ventana de asignación
        ventana_asignacion = tk.Toplevel(root)
        ventana_asignacion.title("Asignación de Correos por Comunidad")
        ventana_asignacion.geometry("1000x600")
        
        # Crear el marco principal con scroll
        main_frame = tk.Frame(ventana_asignacion)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Área con scroll
        canvas = tk.Canvas(main_frame)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)
        
        # Configurar scroll
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Posicionar canvas y scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Frame para la tabla
        frame_tabla = tk.Frame(scrollable_frame)
        frame_tabla.pack(fill='both', expand=True)
        
        # Encabezados de la tabla
        tabla = tk.Frame(frame_tabla)
        tabla.pack(fill='both', expand=True)
        ttk.Label(tabla, text='Enviar', width=8, anchor='center', font=('Arial', 10, 'bold')).grid(row=0, column=0, padx=5, pady=5)
        ttk.Label(tabla, text='Comunidad Original', width=25, anchor='w', font=('Arial', 10, 'bold')).grid(row=0, column=1, padx=5, pady=5)
        ttk.Label(tabla, text='Comunidad Normalizada', width=20, anchor='w', font=('Arial', 10, 'bold')).grid(row=0, column=2, padx=5, pady=5)
        ttk.Label(tabla, text='Correo Asignado', width=35, anchor='w', font=('Arial', 10, 'bold')).grid(row=0, column=3, padx=5, pady=5)
        ttk.Label(tabla, text='Archivo PDF', width=15, anchor='center', font=('Arial', 10, 'bold')).grid(row=0, column=4, padx=5, pady=5)
        ttk.Label(tabla, text='Ver', width=5).grid(row=0, column=5, padx=5, pady=5)

        # Variables para gestionar los datos en la tabla
        check_vars = []
        comunidad_vars = []
        correo_vars = []
        
        # Rellenar la tabla con los datos de las comunidades
        for i, item in enumerate(lista_comunidades_procesadas):
            # Obtener el nombre normalizado para esta comunidad
            nombre_original = item.get('nombre_comunidad_original', '')
            nombre_normalizado = item.get('nombre_comunidad_normalizado', '') 
            
            # Determinar el correo basado en el mapeo (o usar el valor predeterminado)
            correo_asignado = item.get('correo_asignado', 'no@asignado.com')
            
            # Crear variables para cada fila
            var_enviar = tk.BooleanVar(value=True)  # Por defecto se marcan para enviar
            var_comunidad = tk.StringVar(value=nombre_original)
            var_correo = tk.StringVar(value=correo_asignado)
            
            # Guardar referencias a estas variables
            check_vars.append(var_enviar)
            comunidad_vars.append(var_comunidad)
            correo_vars.append(var_correo)
            
            # Checkbox para seleccionar
            check = tk.Checkbutton(tabla, variable=var_enviar)
            check.grid(row=i+1, column=0, padx=5, pady=2)
            
            # Campo de nombre de comunidad original (solo lectura)
            entry_comunidad = tk.Entry(tabla, textvariable=var_comunidad, width=25, state='readonly')
            entry_comunidad.grid(row=i+1, column=1, padx=5, pady=2, sticky='we')
            
            # Mostrar nombre normalizado (solo para información)
            tk.Label(tabla, text=nombre_normalizado, width=20, anchor='w').grid(row=i+1, column=2, padx=5, pady=2)
            
            # Campo editable para el correo
            entry_correo = tk.Entry(tabla, textvariable=var_correo, width=35)
            entry_correo.grid(row=i+1, column=3, padx=5, pady=2, sticky='we')
            
            # Nombre de archivo PDF
            nombre_archivo = os.path.basename(item['ruta_pdf']) if 'ruta_pdf' in item else 'N/A'
            tk.Label(tabla, text=nombre_archivo, width=15, anchor='w').grid(row=i+1, column=4, padx=5, pady=2)
            
            # Botón para abrir el PDF
            btn_abrir = tk.Button(tabla, text="Abrir", command=lambda f=item.get('ruta_pdf', ''): abrir_archivo(f))
            btn_abrir.grid(row=i+1, column=5, padx=5, pady=2)

        # Frame para los botones de acción
        btn_frame = tk.Frame(ventana_asignacion)
        btn_frame.pack(fill='x', pady=15)
        
        # Botón para guardar y continuar
        ttk.Button(btn_frame, text='Confirmar y Continuar', 
                 command=lambda: procesar_seleccion_comunidades()).pack(pady=5)

        def procesar_seleccion_comunidades():
            """Procesa la selección de comunidades y correos para preparar el envío"""
            nonlocal lista_comunidades_procesadas, check_vars, correo_vars, ventana_asignacion
            global remitente_var, gmail_pass_var, carpeta_var
            
            # Filtrar solo las comunidades seleccionadas para envío
            comunidades_para_enviar = []
            
            for idx, item_original in enumerate(lista_comunidades_procesadas):
                if check_vars[idx].get() and correo_vars[idx].get().strip():
                    # El ítem está marcado para enviar y tiene un correo válido
                    item_para_enviar = {
                        'nombre_comunidad_original': item_original.get('nombre_comunidad_original', ''),
                        'nombre_comunidad_normalizado': item_original.get('nombre_comunidad_normalizado', ''),
                        'correo_asignado': correo_vars[idx].get().strip(),
                        'ruta_pdf': item_original.get('ruta_pdf', '')
                    }
                    comunidades_para_enviar.append(item_para_enviar)
            
            # Validar que haya comunidades seleccionadas
            if not comunidades_para_enviar:
                messagebox.showwarning("Advertencia", "No hay comunidades seleccionadas para envío o faltan correos.")
                return
            
            # Llamar a la función de confirmación final con las comunidades filtradas
            if comunidades_para_enviar:
                ventana_asignacion.destroy()
                # Obtenemos los valores antes de cerrar la ventana
                datos_remitente = {
                    'remitente': remitente_var.get() if remitente_var else '',
                    'password': gmail_pass_var.get() if gmail_pass_var else '',
                    'carpeta': carpeta_var.get() if carpeta_var else ''
                }
                # Pasamos los valores explícitamente en lugar de usar nonlocal
                mostrar_dialogo_confirmacion_final(comunidades_para_enviar, datos_remitente)

        # Centrar la ventana
        ventana_asignacion.update_idletasks()
        width = ventana_asignacion.winfo_width()
        height = ventana_asignacion.winfo_height()
        x = (ventana_asignacion.winfo_screenwidth() // 2) - (width // 2)
        y = (ventana_asignacion.winfo_screenheight() // 2) - (height // 2)
        ventana_asignacion.geometry(f'{width}x{height}+{x}+{y}')
        
        # Hacer modal
        ventana_asignacion.transient(root)
        ventana_asignacion.grab_set()
        root.wait_window(ventana_asignacion)
    
    try:
        print('INICIANDO GUI')
        root = tk.Tk()
        root.title("Envío automatizado de facturas")
        root.geometry("800x600")
        
        top_frame = tk.Frame(root)
        top_frame.pack(fill='x', padx=5, pady=5)
        # El botón de búsqueda de facturas ha sido eliminado según lo solicitado

        main_frame = tk.Frame(root)
        main_frame.pack(fill='both', expand=True, padx=10, pady=5)
        frame_img = tk.Frame(main_frame)
        frame_img.pack(fill='x')
        
        global carpeta_var, mes_var, remitente_var, gmail_pass_var, log_box, btn_parar, btn_ejecutar # Declare global for assignment
        carpeta_var = tk.StringVar()
        tk.Entry(frame_img, textvariable=carpeta_var, width=50).pack(side='left', padx=2, expand=True, fill='x')
        
        def browse_folder():
            d = filedialog.askdirectory()
            if d:
                carpeta_var.set(d)
        tk.Button(frame_img, text="Examinar", command=browse_folder).pack(side='left', padx=5)
        
        tk.Label(main_frame, text="Correo Gmail remitente:").pack(anchor='w')
        remitente_var = tk.StringVar()
        tk.Entry(main_frame, textvariable=remitente_var).pack(fill='x')
        tk.Label(main_frame, text="Contraseña de aplicación Gmail:").pack(anchor='w')
        gmail_pass_var = tk.StringVar()
        tk.Entry(main_frame, textvariable=gmail_pass_var, show='*').pack(fill='x')
        
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill='x', pady=10)
        
        def autocompletar_demo():
            # Ruta actualizada para el botón de autocompletar demo
            carpeta_var.set(r'C:\Users\David\Desktop\PROGRAMAR\ENVIADOR DE FACTURAS\FACTURAS\FACTURAS COMPLETA\05 - Mayo - 2025')
            
        # Validar que haya comunidades seleccionadas
        if not comunidades_para_enviar:
            messagebox.showwarning("Advertencia", "No hay comunidades seleccionadas para envío o faltan correos.")
            return
            
        # Llamar a la función de confirmación final con las comunidades filtradas
        if comunidades_para_enviar:
            ventana_asignacion.destroy()
            # Obtenemos los valores antes de cerrar la ventana
            datos_remitente = {
                'remitente': remitente_var.get() if remitente_var else '',
                'password': gmail_pass_var.get() if gmail_pass_var else '',
                'carpeta': carpeta_var.get() if carpeta_var else ''
            }
            # Pasamos los valores explícitamente en lugar de usar nonlocal
            mostrar_dialogo_confirmacion_final(comunidades_para_enviar, datos_remitente)

    # Centrar la ventana
    ventana_asignacion.update_idletasks()
    width = ventana_asignacion.winfo_width()
    height = ventana_asignacion.winfo_height()
    x = (ventana_asignacion.winfo_screenwidth() // 2) - (width // 2)
    y = (ventana_asignacion.winfo_screenheight() // 2) - (height // 2)
    ventana_asignacion.geometry(f'{width}x{height}+{x}+{y}')
        
    # Hacer modal
    ventana_asignacion.transient(root)
    ventana_asignacion.grab_set()
    root.wait_window(ventana_asignacion)
    
try:
    print('INICIANDO GUI')
    root = tk.Tk()
    root.title("Envío automatizado de facturas")
    root.geometry("800x600")
        
    top_frame = tk.Frame(root)
    top_frame.pack(fill='x', padx=5, pady=5)
    # El botón de búsqueda de facturas ha sido eliminado según lo solicitado

    main_frame = tk.Frame(root)
    main_frame.pack(fill='both', expand=True, padx=10, pady=5)
    frame_img = tk.Frame(main_frame)
    frame_img.pack(fill='x')
        
    global carpeta_var, mes_var, remitente_var, gmail_pass_var, log_box, btn_parar, btn_ejecutar # Declare global for assignment
    carpeta_var = tk.StringVar()
    tk.Entry(frame_img, textvariable=carpeta_var, width=50).pack(side='left', padx=2, expand=True, fill='x')
        
    def browse_folder():
        d = filedialog.askdirectory()
        if d:
            carpeta_var.set(d)
    tk.Button(frame_img, text="Examinar", command=browse_folder).pack(side='left', padx=5)
        
    tk.Label(main_frame, text="Correo Gmail remitente:").pack(anchor='w')
    remitente_var = tk.StringVar()
    tk.Entry(main_frame, textvariable=remitente_var).pack(fill='x')
    tk.Label(main_frame, text="Contraseña de aplicación Gmail:").pack(anchor='w')
    gmail_pass_var = tk.StringVar()
    tk.Entry(main_frame, textvariable=gmail_pass_var, show='*').pack(fill='x')
        
    button_frame = tk.Frame(main_frame)
    button_frame.pack(fill='x', pady=10)
        
    def autocompletar_demo():
        # Ruta actualizada para el botón de autocompletar demo
        carpeta_var.set(r'C:\Users\David\Desktop\PROGRAMAR\ENVIADOR DE FACTURAS\FACTURAS\FACTURAS COMPLETA\05 - Mayo - 2025')
        remitente_var.set('davidvr1994@gmail.com')
        gmail_pass_var.set('') # No poner la contraseña real aquí
    tk.Button(button_frame, text="Autocompletar demo", command=autocompletar_demo, bg='#2196F3', fg='white').pack(side='left', padx=5)
        
    def iniciar_proceso_wrapper():
        """Función para iniciar el procesamiento de facturas"""
        nonlocal procesando, detener_proceso_solicitado, mapeo_comunidades_actual
        global carpeta_var, remitente_var, gmail_pass_var
                messagebox.showerror("Error", "Debe seleccionar un directorio válido con facturas PDF.")
                return
                
            # Validar que exista el mapeo de comunidades (cargado desde Excel)
            if not mapeo_comunidades_actual:
                log_func_hilo("ERROR: No se ha podido cargar el mapeo de comunidades desde Excel.")
                messagebox.showerror("Error", "No se ha podido cargar el mapeo de comunidades desde Excel.")
                return
                
            # Cambiar estado a procesando
            procesando = True
            detener_proceso_solicitado = False
            # Actualizar estado de botones
            for widget in button_frame.winfo_children():
                if isinstance(widget, tk.Button):
                    if "Iniciar" in widget["text"]:
                        widget.config(state=tk.DISABLED)
                    elif "Detener" in widget["text"]:
                        widget.config(state=tk.NORMAL)
            
            log_func_hilo(f"Iniciando procesamiento de facturas en: {directorio_seleccionado}")
            log_func_hilo(f"Mapeando {len(mapeo_comunidades_actual)} comunidades desde Excel")
            
            try:
                # Buscar archivos PDF en el directorio
                archivos_pdf = [os.path.join(directorio_seleccionado, f) for f in os.listdir(directorio_seleccionado)
                             if f.lower().endswith('.pdf')]
                
                if not archivos_pdf:
                    log_func_hilo("No se encontraron archivos PDF en el directorio seleccionado.")
                    messagebox.showinfo("Sin datos", "No se encontraron archivos PDF en el directorio seleccionado.")
                    procesando = False
                    
                    # Actualizar estado de botones
                    for widget in button_frame.winfo_children():
                        if isinstance(widget, tk.Button):
                            if "Iniciar" in widget["text"]:
                                widget.config(state=tk.NORMAL)
                            elif "Detener" in widget["text"]:
                                widget.config(state=tk.DISABLED)
                    return
                    
                log_func_hilo(f"Se encontraron {len(archivos_pdf)} archivos PDF para procesar")
                
                # Lista para almacenar los resultados del procesamiento
                comunidades_procesadas = []
                
                # Procesar cada PDF para extraer la comunidad y buscar su correo
                for idx, pdf_path in enumerate(archivos_pdf):
                    if detener_proceso_solicitado:
                        log_func_hilo("Proceso detenido por el usuario.")
                        break
                        
                    pdf_filename = os.path.basename(pdf_path)
                    log_func_hilo(f"Procesando archivo {idx+1}/{len(archivos_pdf)}: {pdf_filename}")
                    
                    try:
                        # Extraer nombre de la comunidad desde el PDF
                        try:
                            from src.extractor_comunidad_pdf import extraer_comunidad
                        except ImportError:
                            from extractor_comunidad_pdf import extraer_comunidad
                        
                        nombre_comunidad = extraer_comunidad(pdf_path)
                        
                        if not nombre_comunidad:
                            log_func_hilo(f"  No se pudo extraer el nombre de comunidad del PDF: {pdf_filename}")
                            continue
                            
                        # Normalizar el nombre para la búsqueda en el mapeo
                        nombre_normalizado = normalizar_nombre_comunidad(nombre_comunidad)
                        
                        # Buscar el correo en el mapeo
                        correo = None
                        for comunidad_map, email_map in mapeo_comunidades_actual.items():
                            comunidad_map_norm = normalizar_nombre_comunidad(comunidad_map)
                            if nombre_normalizado == comunidad_map_norm:
                                correo = email_map
                                break
                        
                        # Guardar los datos procesados
                        comunidades_procesadas.append({
                            'nombre_comunidad_original': nombre_comunidad,
                            'nombre_comunidad_normalizado': nombre_normalizado,
                            'correo_asignado': correo if correo else '',
                            'ruta_pdf': pdf_path
                        })
                        
                        log_func_hilo(f"  Comunidad: {nombre_comunidad}")
                        log_func_hilo(f"  Normalizado: {nombre_normalizado}")
                        log_func_hilo(f"  Correo asignado: {correo if correo else 'NO ENCONTRADO'}")
                        
                    except Exception as e_pdf:
                        log_func_hilo(f"  ERROR al procesar el PDF {pdf_filename}: {str(e_pdf)}")
                
                # Cambiar estado cuando termina el procesamiento
                procesando = False
                
                # Actualizar estado de botones
                for widget in button_frame.winfo_children():
                    if isinstance(widget, tk.Button):
                        if "Iniciar" in widget["text"]:
                            widget.config(state=tk.NORMAL)
                        elif "Detener" in widget["text"]:
                            widget.config(state=tk.DISABLED)
                
                # Si no se encontraron comunidades, mostrar mensaje
                if not comunidades_procesadas:
                    log_func_hilo("No se pudieron extraer nombres de comunidades de los PDFs.")
                    messagebox.showinfo("Sin datos", "No se pudieron extraer nombres de comunidades de los PDFs.")
                    return
                    
                # Mostrar resumen del procesamiento
                total_con_correo = sum(1 for c in comunidades_procesadas if c.get('correo_asignado'))
                total_sin_correo = len(comunidades_procesadas) - total_con_correo
                
                log_func_hilo(f"\nRESUMEN DEL PROCESAMIENTO:")
                log_func_hilo(f"Total de facturas procesadas: {len(comunidades_procesadas)}")
                log_func_hilo(f"Comunidades con correo asignado: {total_con_correo}")
                log_func_hilo(f"Comunidades sin correo asignado: {total_sin_correo}\n")
                
                # Abrir ventana de asignación de correos
                root.after(100, lambda: abrir_ventana_asignacion_correos(comunidades_procesadas))
                
            except Exception as e:
                log_func_hilo(f"ERROR inesperado durante el procesamiento: {str(e)}")
                procesando = False
                
                # Actualizar estado de botones
                for widget in button_frame.winfo_children():
                    if isinstance(widget, tk.Button):
                        if "Iniciar" in widget["text"]:
                            widget.config(state=tk.NORMAL)
                        elif "Detener" in widget["text"]:
                            widget.config(state=tk.DISABLED)
                
        def procesar_en_hilo_wrapper():
            """Inicia el procesamiento en un hilo separado para evitar bloquear la GUI"""
            nonlocal procesando
            if procesando:
                messagebox.showinfo("Aviso", "Ya hay un proceso en ejecución.")
                return
            threading.Thread(target=iniciar_proceso_wrapper, daemon=True).start()

        btn_ejecutar = tk.Button(button_frame, text="Iniciar Proceso", command=lambda: threading.Thread(target=procesar_en_hilo_wrapper, daemon=True).start(), 
                               bg='#4CAF50', fg='white')
        btn_ejecutar.pack(side='left', padx=5)
        btn_parar = tk.Button(button_frame, text="Parar Proceso", command=detener_analisis, 
                            bg='#f44336', fg='white', state='disabled')
        btn_parar.pack(side='left', padx=5)
        
        log_box = scrolledtext.ScrolledText(main_frame, height=9)
        log_box.pack(fill='both', padx=2, pady=2, expand=True)
        
        root.mainloop()
    except Exception as e_gui:
        print('ERROR EN TKINTER:', e_gui)
        print(traceback.format_exc())
        try:
            messagebox.showerror('Error en GUI', str(e_gui) + '\n' + traceback.format_exc())
        except Exception:
            pass # In case root is already destroyed

# La segunda definición de abrir_asignacion_correos y mostrar_confirmacion que estaban al final
# del archivo original han sido omitidas ya que la primera es la que se integra con el flujo principal.

if __name__ == "__main__":
    run_gui()
