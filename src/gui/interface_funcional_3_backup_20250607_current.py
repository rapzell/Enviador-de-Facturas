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

# Variables globales para el estado de la aplicación
procesando = False
detener_proceso_solicitado = False
mapeo_comunidades_actual = {}

def run_gui():
    """Función principal para iniciar la GUI"""
    global procesando, detener_proceso_solicitado, mapeo_comunidades_actual, root, button_frame
    
    print('INICIANDO RUN_GUI')
    
    # Inicializar variables que pueden ser necesarias en el bloque except
    root = None
    
    try:
        # Crear la instancia principal de Tkinter
        root = tk.Tk()
        root.title("EnviadorFacturas - Envío automatizado de facturas a comunidades")
        root.geometry("800x600")
        root.minsize(600, 400)
        
        # Variables de Tkinter para formulario
        carpeta_var = tk.StringVar(root)
        mes_var = tk.StringVar(root) # Variable para el filtro de mes
        remitente_var = tk.StringVar(root)
        gmail_pass_var = tk.StringVar(root)
        
        # Contenedor principal
        main_frame = tk.Frame(root, padx=10, pady=10)
        main_frame.pack(fill='both', expand=True)
        
        # Cargar mapeo inicial desde Excel
        try:
            mapeo_comunidades_actual = cargar_mapeo_desde_excel(print)
        except Exception as e:
            print(f"Error al cargar el mapeo inicial: {e}")
            mapeo_comunidades_actual = {}
        
        # Crear área de logs primero
        log_frame = tk.Frame(main_frame)
        log_frame.pack(fill='both', expand=True, pady=5)
        
        tk.Label(log_frame, text="Registro de operaciones:").pack(anchor='w')
        log_box = scrolledtext.ScrolledText(log_frame, height=10)
        log_box.pack(fill='both', expand=True, padx=2, pady=2)
        log_box.config(state='disabled')
        log_box.config(wrap=tk.WORD)
        
        # Frame para los controles de selección de carpetas y configuración
        control_frame = tk.Frame(main_frame)
        control_frame.pack(fill='x', pady=10)
        
        # Selección de carpeta
        tk.Label(control_frame, text="Carpeta de facturas PDF:").grid(row=0, column=0, sticky='w', pady=5)
        carpeta_entry = tk.Entry(control_frame, textvariable=carpeta_var, width=50)
        carpeta_entry.grid(row=0, column=1, sticky='ew', padx=5)
        tk.Button(control_frame, text="Examinar...", command=lambda: carpeta_var.set(filedialog.askdirectory())).grid(row=0, column=2)
        
        # Datos de remitente
        tk.Label(control_frame, text="Correo Gmail:").grid(row=1, column=0, sticky='w', pady=5)
        tk.Entry(control_frame, textvariable=remitente_var, width=50).grid(row=1, column=1, sticky='ew', padx=5)
        
        tk.Label(control_frame, text="Contraseña:").grid(row=2, column=0, sticky='w', pady=5)
        tk.Entry(control_frame, textvariable=gmail_pass_var, show='*', width=50).grid(row=2, column=1, sticky='ew', padx=5)
        
        # Frame para botones de acción
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill='x', pady=10)
        
        # Declarar botones como variables globales para poder acceder a ellos desde funciones anidadas
        global btn_ejecutar, btn_parar
        
        # Añadir botones de acción
        btn_ejecutar = tk.Button(button_frame, text="Iniciar Proceso", width=15, command=lambda: procesar_en_hilo_wrapper())
        btn_ejecutar.pack(side='left', padx=5)
        
        btn_parar = tk.Button(button_frame, text="Detener Proceso", width=15, command=lambda: detener_proceso_actual(), state=tk.DISABLED)
        btn_parar.pack(side='left', padx=5)
        
        btn_demo = tk.Button(button_frame, text="Demo (buscar PDFs)", width=20, command=lambda: buscar_facturas_numeradas())
        btn_demo.pack(side='left', padx=5)
        
        # Definir la función para los logs ANTES de cualquier uso
        def log_func_hilo(mensaje):
            """Función para manejar logs desde hilos secundarios"""
            log_box.config(state='normal')
            log_box.insert(tk.END, f"{mensaje}\n")
            log_box.see(tk.END)
            log_box.config(state='disabled')
            root.update_idletasks()
            
        def detener_proceso_actual():
            """Solicita detener el proceso en ejecución"""
            global procesando, detener_proceso_solicitado
            if procesando:
                detener_proceso_solicitado = True
                log_func_hilo("Solicitando detener el proceso...")
            else:
                messagebox.showinfo("Info", "No hay ningún proceso en ejecución para detener.")
                
        def buscar_facturas_numeradas():
            """Función para demostrar la búsqueda de PDFs y detectar posibles números de factura"""
            directorio_seleccionado = carpeta_var.get().strip()
            
            # Validar que se haya seleccionado un directorio
            if not directorio_seleccionado or not os.path.isdir(directorio_seleccionado):
                messagebox.showerror("Error", "Debe seleccionar un directorio válido con facturas PDF.")
                return
                
            log_func_hilo(f"DEMO: Analizando archivos PDF en: {directorio_seleccionado}")
            
            # Buscar archivos PDF en el directorio
            try:
                archivos_pdf = [os.path.join(directorio_seleccionado, f) for f in os.listdir(directorio_seleccionado)
                             if f.lower().endswith('.pdf')]
                
                if not archivos_pdf:
                    log_func_hilo("DEMO: No se encontraron archivos PDF en el directorio seleccionado.")
                    return
                    
                log_func_hilo(f"DEMO: Se encontraron {len(archivos_pdf)} archivos PDF")
                
                # Mostrar los primeros 5 archivos como ejemplo (o todos si hay menos de 5)
                max_archivos_demo = min(5, len(archivos_pdf))
                log_func_hilo("\nPrimeros archivos encontrados:")
                
                for i in range(max_archivos_demo):
                    pdf_filename = os.path.basename(archivos_pdf[i])
                    log_func_hilo(f"  {i+1}. {pdf_filename}")
                    
                    # Buscar posibles números de factura en el nombre del archivo
                    import re
                    matches = re.findall(r'\d+', pdf_filename)
                    if matches:
                        log_func_hilo(f"     Posibles números en el nombre: {', '.join(matches)}")
                    
                log_func_hilo("\nUtilice el botón 'Iniciar Proceso' para procesar todos los archivos PDF del directorio.")
                
            except Exception as e_demo:
                log_func_hilo(f"DEMO: Error al analizar el directorio: {str(e_demo)}")
        
        def procesar_en_hilo_wrapper():
            """Inicia el procesamiento en un hilo separado para evitar bloquear la GUI"""
            global procesando
            if procesando:
                messagebox.showinfo("Aviso", "Ya hay un proceso en ejecución.")
                return
            threading.Thread(target=iniciar_proceso_wrapper, daemon=True).start()
            
        def iniciar_proceso_wrapper():
            """Función para iniciar el procesamiento de facturas"""
            global carpeta_var, remitente_var, gmail_pass_var, btn_parar, btn_ejecutar, procesando, detener_proceso_solicitado, mapeo_comunidades_actual
            
            # Variables de control del proceso
            procesando = True
            detener_proceso_solicitado = False
            
            # Obtener el directorio seleccionado de los PDFs
            directorio_seleccionado = carpeta_var.get().strip()
            
            # Validar que se haya seleccionado un directorio
            if not directorio_seleccionado or not os.path.isdir(directorio_seleccionado):
                messagebox.showerror("Error", "Debe seleccionar un directorio válido con facturas PDF.")
                procesando = False
                return
            
            # Validar que exista el mapeo de comunidades (cargado desde Excel)
            if not mapeo_comunidades_actual:
                log_func_hilo("ERROR: No se ha podido cargar el mapeo de comunidades desde Excel.")
                messagebox.showerror("Error", "No se ha podido cargar el mapeo de comunidades desde Excel.")
                procesando = False
                return
            
            # Cambiar estado a procesando
            procesando = True
            detener_proceso_solicitado = False
            
            # Actualizar estado de botones
            btn_ejecutar.config(state=tk.DISABLED)
            btn_parar.config(state=tk.NORMAL)
            
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
                    btn_ejecutar.config(state=tk.NORMAL)
                    btn_parar.config(state=tk.DISABLED)
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
                btn_ejecutar.config(state=tk.NORMAL)
                btn_parar.config(state=tk.DISABLED)
                
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
                btn_ejecutar.config(state=tk.NORMAL)
                btn_parar.config(state=tk.DISABLED)
        
        def abrir_ventana_asignacion_correos(comunidades_procesadas):
            """Abre una ventana para asignar o corregir correos"""
            global mapeo_comunidades_actual
            
            # Crear nueva ventana modal
            ventana_asignacion = tk.Toplevel(root)
            ventana_asignacion.title("Asignación de correos para envío")
            ventana_asignacion.geometry("900x600")
            ventana_asignacion.minsize(700, 500)
            ventana_asignacion.transient(root)
            ventana_asignacion.grab_set()
            
            # Frame principal
            main_frame = tk.Frame(ventana_asignacion, padx=10, pady=10)
            main_frame.pack(fill='both', expand=True)
            
            # Etiqueta informativa
            tk.Label(main_frame, text="Revise y corrija los correos electrónicos para cada comunidad", 
                    font=("Arial", 12, "bold")).pack(pady=10)
            
            tk.Label(main_frame, text="Utilice esta ventana para asignar o corregir los correos automáticos", 
                    font=("Arial", 10)).pack(pady=5)
            
            # Lista para almacenar los widgets de entrada de correo
            correo_entries = []
            
            # Frame con scroll para la lista de comunidades
            scroll_frame = tk.Frame(main_frame)
            scroll_frame.pack(fill='both', expand=True, pady=10)
            
            # Añadir scrollbar
            scrollbar = tk.Scrollbar(scroll_frame)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            canvas = tk.Canvas(scroll_frame)
            canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            canvas.configure(yscrollcommand=scrollbar.set)
            scrollbar.configure(command=canvas.yview)
            
            # Frame interno para contener los elementos
            items_frame = tk.Frame(canvas)
            canvas.create_window((0, 0), window=items_frame, anchor='nw')
            
            # Cabecera
            tk.Label(items_frame, text="#", width=3, font=("Arial", 10, "bold")).grid(row=0, column=0, padx=5, pady=5, sticky="w")
            tk.Label(items_frame, text="Nombre comunidad", width=30, font=("Arial", 10, "bold")).grid(row=0, column=1, padx=5, pady=5, sticky="w")
            tk.Label(items_frame, text="Correo asignado", width=30, font=("Arial", 10, "bold")).grid(row=0, column=2, padx=5, pady=5, sticky="w")
            tk.Label(items_frame, text="Archivo PDF", width=30, font=("Arial", 10, "bold")).grid(row=0, column=3, padx=5, pady=5, sticky="w")
            
            # Lista de comunidades con sus correos
            for idx, item in enumerate(comunidades_procesadas):
                # Índice
                tk.Label(items_frame, text=f"{idx+1}", width=3).grid(row=idx+1, column=0, padx=5, pady=2, sticky="w")
                
                # Nombre de la comunidad
                tk.Label(items_frame, text=item['nombre_comunidad_original'], width=30, anchor="w").grid(
                    row=idx+1, column=1, padx=5, pady=2, sticky="w")
                
                # Correo (editable)
                correo_var = tk.StringVar(value=item['correo_asignado'])
                correo_entry = tk.Entry(items_frame, textvariable=correo_var, width=30)
                correo_entry.grid(row=idx+1, column=2, padx=5, pady=2, sticky="w")
                correo_entries.append(correo_var)  # Guardar referencia para obtener valores luego
                
                # Archivo PDF
                pdf_name = os.path.basename(item['ruta_pdf'])
                tk.Label(items_frame, text=pdf_name, width=30, anchor="w").grid(
                    row=idx+1, column=3, padx=5, pady=2, sticky="w")
            
            # Configurar scrolling
            items_frame.update_idletasks()
            canvas.config(scrollregion=canvas.bbox("all"))
            
            # Botones de acción en la parte inferior
            button_frame = tk.Frame(main_frame)
            button_frame.pack(fill='x', pady=10)
            
            btn_cancelar = tk.Button(button_frame, text="Cancelar", width=15, 
                                  command=lambda: ventana_asignacion.destroy())
            btn_cancelar.pack(side='left', padx=10)
            
            btn_enviar = tk.Button(button_frame, text="Enviar facturas", width=15, 
                                command=lambda: _logica_envio_final())
            btn_enviar.pack(side='right', padx=10)
            
            # Función para el envío final de correos
            def _logica_envio_final():
                # Recopilar comunidades con correos actualizados
                comunidades_para_envio_final = []
                for idx, item in enumerate(comunidades_procesadas):
                    correo_actualizado = correo_entries[idx].get().strip()
                    if correo_actualizado:  # Solo incluir los que tienen correo asignado
                        comunidades_para_envio_final.append({
                            'nombre': item['nombre_comunidad_original'],
                            'correo': correo_actualizado,
                            'pdf': item['ruta_pdf']
                        })
                
                # Verificar si hay elementos para enviar
                if not comunidades_para_envio_final:
                    messagebox.showwarning("Sin destinos", "No hay facturas con correos asignados para enviar.")
                    return
                
                # Validar credenciales del remitente
                remitente_email = remitente_var.get().strip()
                remitente_password = gmail_pass_var.get().strip()
                
                if not remitente_email or not remitente_password or '@' not in remitente_email:
                    messagebox.showerror("Error", "El correo y contraseña del remitente son necesarios.")
                    return
                
                # Mostrar resumen de envío
                mensaje = f"Se enviarán {len(comunidades_para_envio_final)} facturas.\n\n"
                mensaje += f"Desde: {remitente_email}\n\n"
                mensaje += "Primeros 5 destinos:\n"
                
                for idx, item in enumerate(comunidades_para_envio_final[:5]):
                    mensaje += f"{idx+1}. {item['nombre']} -> {item['correo']}\n"
                
                if len(comunidades_para_envio_final) > 5:
                    mensaje += f"...y {len(comunidades_para_envio_final) - 5} más.\n"
                    
                mensaje += "\n¿Desea continuar con el envío?"
                
                # Confirmar el envío
                if messagebox.askyesno("Confirmar envío", mensaje):
                    # Iniciar el envío en un hilo separado para mantener la GUI responsiva
                    threading.Thread(
                        target=lambda: _logica_envio_final_hilo(comunidades_para_envio_final, remitente_email, remitente_password),
                        daemon=True
                    ).start()
                    ventana_asignacion.destroy()
            
            # Función para procesar el envío final en un hilo separado
            def _logica_envio_final_hilo(comunidades_para_envio_final, remitente_email, remitente_password):
                global detener_proceso_solicitado, procesando
                log_func_hilo("Confirmación final recibida. Iniciando envío de correos...")
                
                envios_exitosos = 0
                envios_fallidos = 0

                for idx, item in enumerate(comunidades_para_envio_final):
                    if detener_proceso_solicitado:
                        log_func_hilo("Envío de correos cancelado por el usuario.")
                        break
                    
                    nombre_comunidad = item['nombre']
                    correo_destino = item['correo']
                    ruta_pdf = item['pdf']
                    asunto = f"Factura: {nombre_comunidad}"
                    cuerpo = f"Estimado cliente,\n\nAdjuntamos su factura para {nombre_comunidad}.\n\nSaludos cordiales."
                    
                    log_func_hilo(f"Enviando ({idx+1}/{len(comunidades_para_envio_final)}): {nombre_comunidad} a {correo_destino}")
                    try:
                        try:
                            from src.procesar_envios import enviar_correo
                        except ImportError:
                            from procesar_envios import enviar_correo
                        enviar_correo(remitente_email, remitente_password, correo_destino, asunto, cuerpo, ruta_pdf)
                        envios_exitosos += 1
                    except Exception as e_envio:
                        log_func_hilo(f"ERROR al enviar a {correo_destino} para {nombre_comunidad}: {e_envio}")
                        envios_fallidos += 1
                
                mensaje_final = f"\nProceso de envío finalizado.\nFacturas procesadas: {len(comunidades_para_envio_final)}\nEnviadas con éxito: {envios_exitosos}\nFallidas: {envios_fallidos}"
                log_func_hilo(mensaje_final)
                messagebox.showinfo("Resultado del Envío", mensaje_final)
                
                # Restaurar estado
                procesando = False
                btn_ejecutar.config(state=tk.NORMAL)
                btn_parar.config(state=tk.DISABLED)
        
        # Cargar mapeo de comunidades desde Excel al iniciar
        try:
            mapeo_comunidades_actual = cargar_mapeo_desde_excel(log_func_hilo)
            if mapeo_comunidades_actual:
                log_func_hilo(f"Mapeo cargado correctamente. {len(mapeo_comunidades_actual)} comunidades mapeadas.")
            else:
                log_func_hilo("ADVERTENCIA: No se pudo cargar el mapeo desde Excel o está vacío.")
        except Exception as e_excel:
            print(f"ERROR CRÍTICO: No se pudo cargar el mapeo desde Excel: {e_excel}")
            log_func_hilo(f"ERROR al cargar mapeo desde Excel: {str(e_excel)}")
            messagebox.showerror("Error de Mapeo", f"No se pudo cargar el mapeo desde Excel: {e_excel}")

        def detener_proceso_actual():
            global detener_proceso_solicitado, procesando
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
            
        # Informar sobre estado del mapeo de comunidades
        if not mapeo_comunidades_actual:
            log_func_hilo("ADVERTENCIA: No se pudo cargar el mapeo desde Excel o está vacío.")
        else:
            log_func_hilo(f"Mapeo cargado correctamente. {len(mapeo_comunidades_actual)} comunidades mapeadas.")
        
        def _logica_envio_final(comunidades_para_envio_final, remitente_email, remitente_password, ventana_confirmacion):
            """Función interna para manejar el envío de correos una vez confirmado"""
            global detener_proceso_solicitado, procesando
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
            
            # Totales
            total_facturas = len(comunidades_a_enviar)
            remitente_info = tk.Label(confirm_win, text=f"Remitente: {remitente}")
            remitente_info.pack(pady=5)
            
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
            global mapeo_comunidades_actual, remitente_var, gmail_pass_var, carpeta_var
            
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
        
        # Iniciar el bucle principal de Tkinter para mantener la ventana abierta
        root.mainloop()
        print("GUI cerrada correctamente")
            
    except Exception as e:
        # Importar traceback aquí para asegurar que esté disponible incluso si hubo un error temprano
        import traceback
        print(f"ERROR CRÍTICO EN GUI: {e}")
        traceback.print_exc()
        try:
            # Intentar mostrar un mensaje si la GUI está parcialmente inicializada
            if root and root.winfo_exists():
                messagebox.showerror("Error fatal", f"Error crítico en la aplicación: {e}\nConsulte los logs para más detalles.")
        except Exception:
            # No podemos mostrar un mensaje en la GUI, simplemente continuamos
            pass

# Inicia la interfaz gráfica con manejo de excepciones
import traceback

# Todo el código de creación de GUI ya está dentro de la función run_gui inicial

if __name__ == "__main__":
    print("Comenzando la aplicación")
    # Iniciamos la aplicación sin try/except adicionales para evitar
    # que se capturen excepciones que deberían ser manejadas dentro de run_gui
    run_gui()
    # Esta línea no se ejecutará hasta que la ventana se cierre correctamente
    print("Programa finalizado correctamente")
