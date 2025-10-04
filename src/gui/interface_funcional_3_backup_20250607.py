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
import easyocr # Para el motor OCR


# Ruta al archivo Excel de mapeo por defecto
RUTA_EXCEL_POR_DEFECTO = r"C:\Users\David\Desktop\PROGRAMAR\ENVIADOR DE FACTURAS\ESTRUCTURA\correos_y_comunidades.xlsx"
# lista_mapeos_global se elimina ya que los mapeos serán un diccionario cargado directamente.

# Cuerpo del correo electrónico predeterminado
CUERPO_EMAIL_PREDETERMINADO = """Saludos...
"""

def normalizar_nombre_comunidad(nombre):
    """Normalización básica: minúsculas, elimina 'urb' y 'blq' (y variantes), luego no alfanuméricos. CONSERVA NÚMEROS."""
    if not nombre:
        return ""
    nombre_str = str(nombre).lower()
    if nombre_str == "nan":
        return "" # Tratar 'nan' (de celdas Excel vacías) como inválido

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
        global carpeta_var, remitente_var, gmail_pass_var, excel_file_var # excel_file_var is also used globally within run_gui scope
        global log_box # Declare log_box for use in log_func_hilo
        log_box = None  # Initialize log_box to None

        # Variables de estado para el procesamiento
        procesando = False
        parada_evento = threading.Event() # Usar threading.Event para la señal de parada
        mapeo_comunidades_actual = {} # Inicializado como diccionario
        # excel_file_var is already initialized globally and set via previous edits
        excel_file_var = tk.StringVar(root) # Inicializar nueva variable
        excel_file_var.set(RUTA_EXCEL_POR_DEFECTO) # Establecer valor por defecto

        # --- Early GUI Helper Functions ---
        def log_func_hilo(mensaje):
            print(f"LOG_HILO: {mensaje}") 
            if log_box is not None and log_box.winfo_exists():
                try:
                    current_state = log_box.cget("state")
                    log_box.config(state=tk.NORMAL)
                    log_box.insert(tk.END, f"{mensaje}\n")
                    log_box.see(tk.END)
                    log_box.config(state=current_state)
                    # root.update_idletasks() # Consider if needed for every log, can be intensive
                except Exception as e_log_gui:
                    print(f"LOG_HILO_ERROR (gui): {e_log_gui}")

        def browse_excel_file():
            current_excel_path = excel_file_var.get() or RUTA_EXCEL_POR_DEFECTO
            initial_dir_excel = os.path.dirname(current_excel_path) if current_excel_path and os.path.exists(os.path.dirname(current_excel_path)) else os.getcwd()
            
            filepath = filedialog.askopenfilename(
                title="Seleccionar archivo Excel de mapeo",
                filetypes=(("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")),
                initialdir=initial_dir_excel
            )
            if filepath:
                excel_file_var.set(filepath)
                log_func_hilo(f"Archivo Excel de mapeo seleccionado: {filepath}")
            else:
                log_func_hilo(f"Selección de archivo Excel cancelada. Se mantiene: {excel_file_var.get()}")
        # --- End of Early GUI Helper Functions ---

        # El mapeo de comunidades se cargará dinámicamente en iniciar_proceso_wrapper
        # usando la ruta de excel_file_var. Ya no se carga aquí al inicio.

        def solicitar_parada_proceso():
            nonlocal procesando, parada_evento # Correctly declare nonlocals
            if procesando:
                parada_evento.set() # Signal the event to stop
                log_func_hilo("Solicitud de detención de proceso enviada.")
                messagebox.showinfo("Info", "Se ha solicitado detener el proceso. El proceso se detendrá después de completar la factura actual.")
            else:
                messagebox.showinfo("Info", "No hay ningún proceso en ejecución para detener.")
            
        # Contenedor principal
        main_frame = tk.Frame(root, padx=10, pady=10)
        main_frame.pack(fill='both', expand=True)
        
        # El log sobre el estado del mapeo (incluyendo el archivo Excel por defecto)
        # se hará más adelante, después de que log_func_hilo y log_box estén completamente inicializados.
        # Una llamada similar como log_func_hilo(f"Aplicación GUI iniciada. Mapeo por defecto: {excel_file_var.get()}")
        # ya existe en la versión completa proporcionada anteriormente y está correctamente ubicada.
        
        def abrir_ventana_asignacion_correos(lista_comunidades_procesadas):
            nonlocal root # Necesario para Toplevel. log_func_hilo es accesible.
            log_func_hilo(f"LOG_DEBUG: abrir_ventana_asignacion_correos llamada con {len(lista_comunidades_procesadas)} elementos.")

            # Copiar los datos para no modificar la lista original directamente y añadir estado de selección
            comunidades_procesadas_data = []
            for i, item_orig in enumerate(lista_comunidades_procesadas):
                item_copy = item_orig.copy()
                item_copy['seleccionada_para_envio'] = True # Por defecto, todas seleccionadas
                # item_copy['original_index'] = i # Podría ser útil si se reordena, pero iid del treeview será el índice
                comunidades_procesadas_data.append(item_copy)

            ventana_asignacion = tk.Toplevel(root)
            ventana_asignacion.title("Asignación y Envío de Correos")
            ventana_asignacion.geometry("1000x700") # Tamaño ampliado

            main_frame = ttk.Frame(ventana_asignacion, padding="10")
            main_frame.pack(fill=tk.BOTH, expand=True)

            tree_frame = ttk.Frame(main_frame)
            tree_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

            edit_frame = ttk.LabelFrame(main_frame, text="Editar Fila Seleccionada", padding="10")
            edit_frame.pack(fill=tk.X, pady=(0, 10))

            action_frame = ttk.Frame(main_frame, padding="5") # Reducir padding para botones de acción
            action_frame.pack(fill=tk.X)

            columns = ("enviar", "comunidad", "correo") # pdf_path no se muestra, se usa internamente
            tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="browse")
            
            tree.heading("enviar", text="Enviar?")
            tree.heading("comunidad", text="Comunidad")
            tree.heading("correo", text="Correo Electrónico")

            tree.column("enviar", width=60, anchor=tk.CENTER, stretch=tk.NO)
            tree.column("comunidad", width=350)
            tree.column("correo", width=350)

            vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
            vsb.pack(side='right', fill='y')
            hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
            hsb.pack(side='bottom', fill='x')
            tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
            tree.pack(fill=tk.BOTH, expand=True)

            def populate_tree():
                for item in tree.get_children():
                    tree.delete(item)
                for idx, data in enumerate(comunidades_procesadas_data):
                    enviar_str = "Sí" if data.get('seleccionada_para_envio', True) else "No"
                    tree.insert("", tk.END, values=(enviar_str, data.get('nombre_comunidad', ''), data.get('correo_asignado', '')), iid=str(idx))
            
            populate_tree()

            ttk.Label(edit_frame, text="Comunidad:").grid(row=0, column=0, padx=5, pady=2, sticky="w")
            comunidad_var = tk.StringVar()
            comunidad_entry = ttk.Entry(edit_frame, textvariable=comunidad_var, width=60)
            comunidad_entry.grid(row=0, column=1, padx=5, pady=2, sticky="ew")

            ttk.Label(edit_frame, text="Correo:").grid(row=1, column=0, padx=5, pady=2, sticky="w")
            correo_var = tk.StringVar()
            correo_entry = ttk.Entry(edit_frame, textvariable=correo_var, width=60)
            correo_entry.grid(row=1, column=1, padx=5, pady=2, sticky="ew")
            
            pdf_path_display_var = tk.StringVar()
            ttk.Label(edit_frame, text="Archivo PDF:").grid(row=2, column=0, padx=5, pady=2, sticky="w")
            ttk.Label(edit_frame, textvariable=pdf_path_display_var, foreground="blue", cursor="hand2").grid(row=2, column=1, padx=5, pady=2, sticky="w")
            edit_frame.columnconfigure(1, weight=1)

            selected_item_iid = None

            def on_tree_select(event):
                nonlocal selected_item_iid
                selection = tree.selection()
                if not selection:
                    selected_item_iid = None
                    comunidad_var.set("")
                    correo_var.set("")
                    pdf_path_display_var.set("")
                    return
                selected_item_iid = selection[0]
                try:
                    item_idx = int(selected_item_iid)
                    data = comunidades_procesadas_data[item_idx]
                    comunidad_var.set(data.get('nombre_comunidad', ''))
                    correo_var.set(data.get('correo_asignado', ''))
                    pdf_path_display_var.set(os.path.basename(data.get('ruta_pdf', '')))
                except (ValueError, IndexError) as e:
                    log_func_hilo(f"Error al seleccionar item del tree: {e}. IID: {selected_item_iid}")
                    selected_item_iid = None
                    comunidad_var.set(""); correo_var.set(""); pdf_path_display_var.set("")
            tree.bind("<<TreeviewSelect>>", on_tree_select)

            def toggle_send_status(event):
                clicked_iid = tree.identify_row(event.y)
                log_func_hilo(f"LOG_DEBUG: toggle_send_status - event.y: {event.y}, clicked_iid: {clicked_iid}") # Nuevo log
                if not clicked_iid: return
                try:
                    item_idx = int(clicked_iid)
                    data = comunidades_procesadas_data[item_idx]
                    data['seleccionada_para_envio'] = not data.get('seleccionada_para_envio', True)
                    log_func_hilo(f"Comunidad '{data['nombre_comunidad']}' marcada para envío: {data['seleccionada_para_envio']}")
                    populate_tree()
                    tree.selection_set(clicked_iid) # Mantener selección visual
                    tree.focus(clicked_iid)
                except (ValueError, IndexError, KeyError) as e:
                    log_func_hilo(f"Error al cambiar estado de envío: {e}. IID: {clicked_iid}")
            tree.bind("<Double-1>", toggle_send_status)

            def save_row_changes():
                if not selected_item_iid:
                    messagebox.showwarning("Sin Selección", "Seleccione una fila para guardar.", parent=ventana_asignacion)
                    return
                try:
                    item_idx = int(selected_item_iid)
                    comunidades_procesadas_data[item_idx]['nombre_comunidad'] = comunidad_var.get()
                    comunidades_procesadas_data[item_idx]['correo_asignado'] = correo_var.get()
                    log_func_hilo(f"Cambios guardados para: {comunidades_procesadas_data[item_idx]['nombre_comunidad']}")
                    populate_tree()
                    tree.selection_set(selected_item_iid); tree.focus(selected_item_iid)
                    messagebox.showinfo("Guardado", "Cambios guardados.", parent=ventana_asignacion)
                except (ValueError, IndexError, KeyError) as e:
                    messagebox.showerror("Error", f"No se pudieron guardar cambios: {e}", parent=ventana_asignacion)

            def view_pdf_for_selected_row(): # Renombrado para claridad
                if not selected_item_iid:
                    messagebox.showwarning("Sin Selección", "Seleccione una fila para ver el PDF.", parent=ventana_asignacion)
                    return
                try:
                    item_idx = int(selected_item_iid)
                    pdf_path = comunidades_procesadas_data[item_idx].get('ruta_pdf')
                    if pdf_path and os.path.exists(pdf_path):
                        log_func_hilo(f"Abriendo PDF: {pdf_path}")
                        os.startfile(pdf_path)
                    else:
                        messagebox.showerror("Error", f"PDF no encontrado: {pdf_path}", parent=ventana_asignacion)
                except (ValueError, IndexError, KeyError) as e:
                    messagebox.showerror("Error", f"No se pudo abrir PDF: {e}", parent=ventana_asignacion)
            
            # Vincular clic en la etiqueta del PDF para abrirlo también
            pdf_label_widget = edit_frame.grid_slaves(row=2, column=1)[0]
            pdf_label_widget.bind("<Button-1>", lambda e: view_pdf_for_selected_row())

            btn_save_changes = ttk.Button(edit_frame, text="Guardar Cambios en Fila", command=save_row_changes)
            btn_save_changes.grid(row=3, column=0, padx=5, pady=5, sticky="ew")
            btn_view_pdf = ttk.Button(edit_frame, text="Ver PDF de la Fila", command=view_pdf_for_selected_row)
            btn_view_pdf.grid(row=3, column=1, padx=5, pady=5, sticky="ew")

            def send_selected_action():
                remitente_email = remitente_var.get()
                remitente_password = gmail_pass_var.get()

                if not remitente_email or not remitente_password:
                    messagebox.showerror("Error de Credenciales", "Por favor, ingrese el correo remitente y la contraseña de Gmail en la pantalla principal.", parent=ventana_asignacion)
                    return

                items_para_enviar_raw = [d for d in comunidades_procesadas_data if d.get('seleccionada_para_envio', True) and d.get('correo_asignado', '').strip()]
                
                if not items_para_enviar_raw:
                    messagebox.showinfo("Nada para Enviar", "No hay facturas seleccionadas para enviar o ninguna tiene un correo asignado válido.", parent=ventana_asignacion)
                    return

                correos_agrupados = {}
                for item in items_para_enviar_raw:
                    correo_destinatario = item.get('correo_asignado').strip()
                    if correo_destinatario not in correos_agrupados:
                        correos_agrupados[correo_destinatario] = []
                    correos_agrupados[correo_destinatario].append({
                        'nombre_comunidad': item.get('nombre_comunidad', 'N/A'),
                        'ruta_pdf': item.get('ruta_pdf', 'N/A')
                    })
                
                if not correos_agrupados:
                     messagebox.showinfo("Nada para Enviar", "No se pudo agrupar ninguna factura para envío (error inesperado).", parent=ventana_asignacion)
                     return

                log_func_hilo(f"Iniciando proceso de envío real para {len(correos_agrupados)} destinatario(s).")
                
                enviados_count = 0
                errores_count = 0
                detalles_errores = []

                for destinatario_email, adjuntos_info_list in correos_agrupados.items():
                    rutas_pdfs_adjuntar = [adj_info['ruta_pdf'] for adj_info in adjuntos_info_list if adj_info.get('ruta_pdf') != 'N/A']
                    nombres_comunidades_list = sorted(list(set(adj_info['nombre_comunidad'] for adj_info in adjuntos_info_list))) # sorted unique names
                    nombres_comunidades = ", ".join(nombres_comunidades_list)
                    
                    if not rutas_pdfs_adjuntar:
                        log_func_hilo(f"No hay PDFs válidos para enviar a {destinatario_email} para comunidades {nombres_comunidades}. Saltando.")
                        errores_count +=1
                        detalles_errores.append(f"Destinatario {destinatario_email} (Comunidades: {nombres_comunidades}): Sin PDFs válidos.")
                        continue

                    asunto = f"Facturas Comunidades: {nombres_comunidades}"
                    cuerpo = CUERPO_EMAIL_PREDETERMINADO 

                    log_func_hilo(f"Intentando enviar a: {destinatario_email}, Asunto: \"{asunto}\", Adjuntos: {len(rutas_pdfs_adjuntar)}")
                    
                    try:
                        # La función importada es 'enviar_correo' de 'procesar_envios'
                        # enviar_correo(remitente, password, destinatario, asunto, cuerpo, lista_adjuntos_paths)
                        enviar_correo(remitente_email, remitente_password, destinatario_email, asunto, cuerpo, rutas_pdfs_adjuntar)
                        log_func_hilo(f"Correo enviado exitosamente a {destinatario_email} para {nombres_comunidades}.")
                        enviados_count += 1
                    except Exception as e:
                        error_msg = f"Error al enviar correo a {destinatario_email} para {nombres_comunidades}: {str(e)}"
                        log_func_hilo(error_msg)
                        errores_count += 1
                        detalles_errores.append(error_msg)

                resumen_msg = f"Proceso de envío completado.\n\nCorreos enviados exitosamente: {enviados_count}\nCorreos con errores: {errores_count}"
                if errores_count > 0:
                    resumen_msg += "\n\nDetalles de errores:\n" + "\n".join(detalles_errores)
                
                log_func_hilo(resumen_msg) # Log the full summary
                
                # Determine title and type of messagebox based on errors
                if errores_count > 0 and enviados_count > 0:
                    msg_title = "Resultado del Envío (con errores)"
                    messagebox.showwarning(msg_title, resumen_msg, parent=ventana_asignacion)
                elif errores_count > 0 and enviados_count == 0:
                    msg_title = "Error de Envío"
                    messagebox.showerror(msg_title, resumen_msg, parent=ventana_asignacion)
                else: # No errors
                    msg_title = "Envío Exitoso"
                    messagebox.showinfo(msg_title, resumen_msg, parent=ventana_asignacion)
                
                # Opcional: cerrar después de enviar
                # ventana_asignacion.destroy()

            btn_send_selected = ttk.Button(action_frame, text="Enviar Seleccionados", command=send_selected_action)
            btn_send_selected.pack(side=tk.LEFT, padx=10, pady=5)
            btn_cerrar_asignacion = ttk.Button(action_frame, text="Cerrar", command=ventana_asignacion.destroy)
            btn_cerrar_asignacion.pack(side=tk.RIGHT, padx=10, pady=5)
            
            ventana_asignacion.protocol("WM_DELETE_WINDOW", ventana_asignacion.destroy) # Asegurar limpieza
            ventana_asignacion.grab_set()
            ventana_asignacion.transient(root)
            ventana_asignacion.wait_window()

        
        top_frame = tk.Frame(root)
        top_frame.pack(fill='x', padx=5, pady=5)
        # El botón de búsqueda de facturas ha sido eliminado según lo solicitado

        main_frame = tk.Frame(root)
        main_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        frame_img = tk.Frame(main_frame) # Frame para la selección de carpeta
        frame_img.pack(fill='x')
        
        # Variable y Entry para la ruta de la carpeta
        # Asegurarse que carpeta_var se define antes de ser usada por autocompletar_demo
        global carpeta_var
        carpeta_var = tk.StringVar()
        tk.Entry(frame_img, textvariable=carpeta_var, width=50).pack(side='left', padx=2, expand=True, fill='x')
        
        # Botón para examinar carpetas
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
        
        # --- Excel File Frame ---
        excel_file_frame = tk.LabelFrame(main_frame, text="Archivo de Mapeo Excel", padx=10, pady=5)
        excel_file_frame.pack(fill='x', pady=5, before=button_frame if 'button_frame' in locals() else None) # Try to place it before button_frame
        
        tk.Label(excel_file_frame, text="Ruta:").grid(row=0, column=0, sticky='w', padx=(0,5), pady=2)
        excel_entry = tk.Entry(excel_file_frame, textvariable=excel_file_var, width=60) 
        excel_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=2)
        
        excel_browse_button = tk.Button(excel_file_frame, text="Examinar...", command=browse_excel_file)
        excel_browse_button.grid(row=0, column=2, padx=(5,0), pady=2)
        excel_file_frame.columnconfigure(1, weight=1) # Make entry expandable
        # --- End of Excel File Frame ---

        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill='x', pady=10)
        
        # La primera definición de autocompletar_demo ha sido eliminada, la siguiente es la correcta.
        def autocompletar_demo():
            # Ruta actualizada para el botón de autocompletar demo
            carpeta_var.set(r'C:\Users\David\Desktop\PROGRAMAR\ENVIADOR DE FACTURAS\FACTURAS\FACTURAS COMPLETA\05 - Mayo - 2025')
            remitente_var.set('limpiezasmaylinsl@gmail.com')
            gmail_pass_var.set('emjt aedb luqc xkbr')
        tk.Button(button_frame, text="Autocompletar demo", command=autocompletar_demo, bg='#2196F3', fg='white').pack(side='left', padx=5)
        
        def iniciar_proceso_wrapper():
            """Función para iniciar el procesamiento de facturas"""
            log_func_hilo("DEBUG: iniciar_proceso_wrapper llamado")
            nonlocal procesando, parada_evento, mapeo_comunidades_actual # Correctly use parada_evento
            global carpeta_var, remitente_var, gmail_pass_var, excel_file_var

            directorio_seleccionado = carpeta_var.get() # Define and get the directory
            log_func_hilo(f"DEBUG: Directorio seleccionado: '{directorio_seleccionado}'")
            if not directorio_seleccionado: # Check if the directory is selected
                messagebox.showerror("Error", "Debe seleccionar un directorio válido con facturas PDF.")
                return
                    
            # Cargar mapeo de comunidades desde Excel (ruta dinámica, formato: Email | Comunidad1 | Comunidad2 | ...)
            ruta_excel_seleccionada = excel_file_var.get()
            log_func_hilo(f"Intentando cargar mapeo desde: {ruta_excel_seleccionada}")
            try:
                if not ruta_excel_seleccionada or not os.path.exists(ruta_excel_seleccionada):
                    error_msg = f"La ruta del archivo Excel no es válida o el archivo no existe: {ruta_excel_seleccionada}"
                    log_func_hilo(f"ERROR: {error_msg}")
                    messagebox.showerror("Error de Archivo Excel", error_msg)
                    return # Salir si el archivo no es válido
                df_mapeo = pd.read_excel(ruta_excel_seleccionada, header=None) # Leer sin encabezados

                if df_mapeo.empty:
                    error_msg = "El archivo Excel de mapeo está vacío."
                    log_func_hilo(f"ERROR: {error_msg}")
                    messagebox.showerror("Error de Mapeo", error_msg)
                    return

                mapeo_comunidades_actual.clear() # Limpiar mapeo previo

                for index, row in df_mapeo.iterrows():
                    if row.empty or pd.isna(row.iloc[0]):
                        log_func_hilo(f"Advertencia: Fila {index + 1} en Excel vacía o sin email en la primera columna. Omitiendo fila.")
                        continue
                    
                    email_actual = str(row.iloc[0]).strip()
                    if not email_actual: # Doble chequeo por si acaso es solo espacios
                        log_func_hilo(f"Advertencia: Fila {index + 1} en Excel con email vacío en la primera columna. Omitiendo fila.")
                        continue

                    comunidades_en_fila = 0
                    for i in range(1, len(row)):
                        comunidad_raw = str(row.iloc[i])
                        if pd.isna(comunidad_raw) or comunidad_raw.strip() == "":
                            continue # Saltar celdas de comunidad vacías
                        
                        comunidad_normalizada = normalizar_nombre_comunidad(comunidad_raw.strip())
                        if comunidad_normalizada:
                            if comunidad_normalizada in mapeo_comunidades_actual and mapeo_comunidades_actual[comunidad_normalizada] != email_actual:
                                log_func_hilo(f"ADVERTENCIA: Comunidad duplicada '{comunidad_normalizada}' (de '{comunidad_raw}') en fila {index+1} asignada a '{email_actual}'. Ya estaba asignada a '{mapeo_comunidades_actual[comunidad_normalizada]}'. Se SOBREESCRIBIRÁ con el email de esta fila.")
                            mapeo_comunidades_actual[comunidad_normalizada] = email_actual
                            comunidades_en_fila += 1
                    
                    if comunidades_en_fila == 0:
                        log_func_hilo(f"Advertencia: Fila {index+1} con email '{email_actual}' no tiene comunidades asociadas válidas.")

                if not mapeo_comunidades_actual:
                    error_msg = "No se cargaron datos de mapeo válidos desde Excel. Verifique el archivo y los logs."
                    log_func_hilo(f"ERROR: {error_msg}")
                    messagebox.showerror("Error de Mapeo", error_msg)
                    return
                else:
                    log_func_hilo(f"Mapeo cargado exitosamente: {len(mapeo_comunidades_actual)} comunidades mapeadas.")

            except FileNotFoundError: # Este bloque podría volverse redundante con el chequeo os.path.exists de arriba
                error_msg = f"No se encontró el archivo Excel de mapeo en la ruta: {ruta_excel_seleccionada}"
                log_func_hilo(f"ERROR: {error_msg}")
                messagebox.showerror("Error de Mapeo", error_msg)
                return
            except Exception as e:
                error_msg = f"Ocurrió un error al leer el archivo Excel de mapeo: {str(e)}"
                log_func_hilo(f"ERROR: {error_msg}")
                messagebox.showerror("Error de Mapeo", error_msg)
                return

            # Log de depuración después del intento de carga
            log_func_hilo(f"DEBUG: Mapeo comunidades (después de carga): {mapeo_comunidades_actual is not None}, len: {len(mapeo_comunidades_actual) if mapeo_comunidades_actual is not None else 'N/A'}")
            if not mapeo_comunidades_actual: # Doble chequeo, por si acaso algo muy raro pasó
                log_func_hilo("ERROR CRÍTICO: Mapeo de comunidades vacío después del intento de carga, a pesar de no errores previos.")
                messagebox.showerror("Error Interno", "El mapeo de comunidades está vacío después de intentar cargarlo.")
                return

            # El siguiente bloque es el que maneja el procesamiento de PDFs
            # y debe estar al mismo nivel de indentación que el bloque try-except de carga de Excel.

            # --- Inicializar EasyOCR Reader --- 
            ocr_reader = None
            try:
                log_func_hilo("Inicializando motor OCR (EasyOCR)... Esto puede tardar un momento.")
                ocr_reader = easyocr.Reader(['es'], verbose=False) # verbose=False para menos spam en consola
                log_func_hilo("Motor OCR inicializado correctamente.")
            except Exception as e_ocr_init:
                log_func_hilo(f"ERROR CRÍTICO: No se pudo inicializar el motor OCR: {str(e_ocr_init)}")
                messagebox.showerror("Error OCR", f"No se pudo inicializar el motor OCR: {str(e_ocr_init)}. El proceso no puede continuar.")
                # No es necesario cambiar 'procesando' aquí, ya que aún no se ha establecido en True para el procesamiento de PDFs
                # Re-habilitar botón Iniciar si es necesario, aunque el flujo principal se detiene.
                for widget_btn in button_frame.winfo_children():
                    if isinstance(widget_btn, tk.Button) and "Iniciar" in widget_btn["text"]:
                        widget_btn.config(state=tk.NORMAL)
                return # Salir de la función si el OCR no se puede cargar
            
            # Cambiar estado a procesando
            procesando = True
            parada_evento.clear() # Ensure the event is clear at the start of processing
            
            # Actualizar estado de botones
            for widget in button_frame.winfo_children():
                if isinstance(widget, tk.Button):
                    if "Iniciar" in widget["text"]:
                        widget.config(state=tk.DISABLED)
                    elif widget["text"] == "Parar Proceso":
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
                    log_func_hilo(f"LOG_DEBUG: Bucle PDF {idx+1} - Antes de check parada_evento.is_set(): {parada_evento.is_set()}")
                    if parada_evento.is_set(): # Check if the stop event has been signaled
                        log_func_hilo("LOG_DEBUG: Bucle PDF - parada_evento.is_set() es TRUE. Interrumpiendo bucle.")
                        log_func_hilo("Proceso detenido por el usuario.")
                        break
                        
                    pdf_filename = os.path.basename(pdf_path)
                    log_func_hilo(f"Procesando archivo {idx+1}/{len(archivos_pdf)}: {pdf_filename}")
                    
                    try: # Inicio del try para procesar un PDF individual
                        # Extraer nombre de la comunidad desde el PDF
                        try:
                            from src.extractor_comunidad_pdf import extraer_comunidad_de_pdf
                        except ImportError:
                            from extractor_comunidad_pdf import extraer_comunidad_de_pdf # Asumiendo que está en el mismo dir si src falla
                        
                        nombre_comunidad = extraer_comunidad_de_pdf(pdf_path, ocr_reader) # Pasar el reader inicializado
                        
                        if not nombre_comunidad:
                            log_func_hilo(f"  No se pudo extraer el nombre de comunidad del PDF: {pdf_filename}")
                            continue # Saltar al siguiente PDF
                            
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
                            'nombre_comunidad': nombre_comunidad, # Clave cambiada de 'nombre_comunidad_original'
                            'nombre_comunidad_normalizado': nombre_normalizado,
                            'correo_asignado': correo if correo else '',
                            'ruta_pdf': pdf_path
                        })
                        
                        log_func_hilo(f"  Comunidad: {nombre_comunidad}")
                        log_func_hilo(f"  Normalizado: {nombre_normalizado}")
                        log_func_hilo(f"  Correo asignado: {correo if correo else 'NO ENCONTRADO'}")
                        
                    except Exception as e_pdf: # Manejo de error para el PDF individual
                        log_func_hilo(f"  ERROR al procesar el PDF {pdf_filename}: {str(e_pdf)}")
                
                # Cambiar estado cuando termina el procesamiento de todos los PDFs
                procesando = False
                
                # Actualizar estado de botones
                for widget in button_frame.winfo_children():
                    if isinstance(widget, tk.Button):
                        if "Iniciar" in widget["text"]:
                            widget.config(state=tk.NORMAL)
                        elif widget["text"] == "Parar Proceso":
                            widget.config(state=tk.DISABLED)
                
                # Lógica después del bucle
                log_func_hilo(f"LOG_DEBUG: Post-bucle - Comunidades procesadas: {len(comunidades_procesadas)}, parada_evento.is_set(): {parada_evento.is_set()}")
                # Lógica después del bucle: decidir si mostrar ventana de asignación o mensaje
                if not comunidades_procesadas:
                    if parada_evento.is_set(): # Check if stopped by user
                        log_func_hilo("Proceso detenido por el usuario antes de procesar alguna factura.")
                        messagebox.showinfo("Proceso Detenido", "El proceso fue detenido antes de que se pudiera procesar alguna factura.")
                    else:
                        log_func_hilo("No se encontraron comunidades en los PDFs procesados o no se encontraron PDFs válidos.")
                        messagebox.showinfo("Sin Datos", "No se encontraron comunidades en los PDFs procesados o no se encontraron PDFs válidos. Verifique los logs.")
                    return # Salir si no hay datos para asignar
                    
                # Mostrar resumen del procesamiento
                total_con_correo = sum(1 for c in comunidades_procesadas if c.get('correo_asignado'))
                total_sin_correo = len(comunidades_procesadas) - total_con_correo
                
                log_func_hilo(f"\nRESUMEN DEL PROCESAMIENTO:")
                log_func_hilo(f"Total de facturas procesadas: {len(comunidades_procesadas)}")
                log_func_hilo(f"Comunidades con correo asignado: {total_con_correo}")
                log_func_hilo(f"Comunidades sin correo asignado: {total_sin_correo}\n")
                
                # Abrir ventana de asignación de correos
                log_func_hilo(f"LOG_DEBUG: Intentando abrir ventana de asignación con {len(comunidades_procesadas)} comunidades.")
                root.after(100, lambda: abrir_ventana_asignacion_correos(comunidades_procesadas))
                
            except Exception as e:
                log_func_hilo(f"ERROR inesperado durante el procesamiento: {str(e)}")
                procesando = False
                
                # Actualizar estado de botones
                for widget in button_frame.winfo_children():
                    if isinstance(widget, tk.Button):
                        if "Iniciar" in widget["text"]:
                            widget.config(state=tk.NORMAL)
                        elif widget["text"] == "Parar Proceso":
                            widget.config(state=tk.DISABLED)
                    
        def procesar_en_hilo_wrapper():
            """Inicia el procesamiento en un hilo separado para evitar bloquear la GUI"""
            nonlocal procesando
            log_func_hilo("DEBUG: procesar_en_hilo_wrapper llamado")
            if procesando:
                messagebox.showinfo("Aviso", "Ya hay un proceso en ejecución.")
                return
            log_func_hilo("DEBUG: Creando e iniciando hilo para iniciar_proceso_wrapper")
            threading.Thread(target=iniciar_proceso_wrapper, daemon=True).start()

        def detener_analisis():
            nonlocal parada_evento # Usar el evento del ámbito de run_gui
            parada_evento.set() # Establecer el evento para señalar la detención
            log_func_hilo("LOG_DEBUG: detener_analisis - parada_evento.set() llamado.")
            log_func_hilo("Solicitud de detención de análisis recibida (evento establecido).")

        btn_ejecutar = tk.Button(button_frame, text="Iniciar Proceso", command=procesar_en_hilo_wrapper, 
                               bg='#4CAF50', fg='white')
        btn_ejecutar.pack(side='left', padx=5)
        btn_parar = tk.Button(button_frame, text="Parar Proceso", command=solicitar_parada_proceso, 
                            bg='#f44336', fg='white', state='disabled')
        btn_parar.pack(side='left', padx=5)
        
        log_box = scrolledtext.ScrolledText(main_frame, height=9)
        log_box.pack(fill='both', padx=2, pady=2, expand=True)
        
        root.mainloop()
    except Exception as e_gui:
        print('ERROR EN TKINTER:', e_gui)
        import traceback
        print(traceback.format_exc())
        try:
            messagebox.showerror('Error en GUI', str(e_gui) + '\n' + traceback.format_exc())
        except Exception:
            pass # In case root is already destroyed

if __name__ == "__main__":
    run_gui()
