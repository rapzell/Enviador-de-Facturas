import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
import threading
import os
import sys
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from procesar_envios import procesar_envios, procesar_nombre_factura
from enviar_factura import enviar_factura
import pandas as pd # Para leer Excel
# openpyxl será usado por pandas implícitamente para .xlsx, asegúrate de que esté instalado.
import re
import time
import unicodedata



# Ruta al archivo Excel de mapeo por defecto
RUTA_EXCEL_POR_DEFECTO = ""
# lista_mapeos_global se elimina ya que los mapeos serán un diccionario cargado directamente.

# Cuerpo del correo electrónico predeterminado
CUERPO_EMAIL_PREDETERMINADO = """Saludos...
"""

_EMAIL_REGEX = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")

def _extraer_email_de_texto(valor: str):
    """Devuelve el primer email encontrado dentro del texto o None si no hay."""
    if not valor:
        return None
    m = _EMAIL_REGEX.search(str(valor))
    return m.group(0) if m else None

def _limpiar_posible_email(valor: str) -> str:
    """Elimina espacios y caracteres invisibles comunes para detectar emails escritos con espacios."""
    if not valor:
        return valor
    s = str(valor)
    # quitar espacios y tabulaciones dentro del email
    s = re.sub(r"\s+", "", s)
    # quitar caracteres de control invisibles
    s = ''.join(ch for ch in s if ch.isprintable())
    return s

def _normalizar_nombre(texto: str) -> str:
    if texto is None:
        return ""
    s = str(texto).strip().lower()
    # Quitar acentos
    s = ''.join(c for c in unicodedata.normalize('NFKD', s) if not unicodedata.combining(c))
    # Unificar separadores y espacios
    s = s.replace('\u00ba', 'º')  # por si acaso
    s = s.replace('\n', ' ').replace('\t', ' ')
    s = re.sub(r"[\s\-_/]+", " ", s)  # colapsar separadores comunes a un espacio
    s = re.sub(r"\s+", " ", s)
    s = s.strip()
    return s

def cargar_mapeo_desde_excel(ruta_excel, log_func):
    """Carga el mapeo comunidad -> email detectando automáticamente la(s) celda(s) email por fila.
    Soporta hojas sin cabecera y filas con varias columnas.
    """
    mapeo_para_procesamiento = {}
    if not os.path.exists(ruta_excel):
        log_func(f"ERROR: Archivo Excel de mapeo no encontrado en {ruta_excel}. No se cargarán mapeos.")
        return mapeo_para_procesamiento

    try:
        hojas = pd.read_excel(ruta_excel, header=None, sheet_name=None, dtype=str, engine='openpyxl')
        log_func(f"Cargando mapeo de correos desde Excel: {ruta_excel}")

        total_emails_detectados = 0
        for nombre_hoja, df in hojas.items():
            filas_total = len(df.index)
            log_func(f"Hoja '{nombre_hoja}': {filas_total} filas leídas.")
            filas_con_email = 0
            for index, row in df.iterrows():
                # Saltar filas completamente vacías
                if row.isnull().all() or (row.astype(str).str.strip() == '').all():
                    continue

                # Buscar emails en cualquier columna de la fila
                valores = [None if pd.isna(v) else str(v).strip() for v in row.tolist()]
                emails_en_fila = []
                for v in valores:
                    if not v:
                        continue
                    email = _extraer_email_de_texto(v) or _extraer_email_de_texto(_limpiar_posible_email(v))
                    if email:
                        emails_en_fila.append(email)

                if not emails_en_fila:
                    if index < 5:
                        log_func(f"Debug hoja '{nombre_hoja}' fila {index+1}: sin email en fila (primer valor='{valores[0] if valores else ''}')")
                    continue

                correo = emails_en_fila[0]
                filas_con_email += 1
                total_emails_detectados += 1
                if index < 5:
                    log_func(f"Debug hoja '{nombre_hoja}' fila {index+1}: email_detectado='{correo}'")

                # Mapear todas las celdas no-email de la fila a este correo
                for celda in valores:
                    if not celda:
                        continue
                    if _extraer_email_de_texto(celda) or _extraer_email_de_texto(_limpiar_posible_email(celda)):
                        continue
                    nombre_comunidad_original = celda
                    nombre_comunidad_normalizada = _normalizar_nombre(nombre_comunidad_original)
                    if not nombre_comunidad_normalizada:
                        continue
                    if nombre_comunidad_normalizada in mapeo_para_procesamiento and mapeo_para_procesamiento[nombre_comunidad_normalizada] != correo:
                        log_func(
                            f"ADVERTENCIA: Hoja '{nombre_hoja}', fila {index+1}: Comunidad duplicada '{nombre_comunidad_original}' (normalizada '{nombre_comunidad_normalizada}') "
                            f"mapeada previamente a '{mapeo_para_procesamiento[nombre_comunidad_normalizada]}' y ahora a '{correo}'. Se usará '{correo}'."
                        )
                    mapeo_para_procesamiento[nombre_comunidad_normalizada] = correo

            log_func(f"Hoja '{nombre_hoja}': {filas_con_email} filas con email detectadas.")

        log_func(f"Total emails detectados en todas las hojas: {total_emails_detectados}")
        # Fallback: si no se detectó ningún mapeo por filas, intentar orientación por columnas
        if len(mapeo_para_procesamiento) == 0:
            log_func("No se generó mapeo por filas. Intentando detección por columnas (emails en cabecera/columna)...")
            total_emails_col = 0
            for nombre_hoja, df in hojas.items():
                cols = df.shape[1]
                columnas_con_email = 0
                for col_idx in range(cols):
                    # extraer columna como lista de strings
                    col_vals = [None if pd.isna(v) else str(v).strip() for v in df.iloc[:, col_idx].tolist()]
                    # buscar primer email en la columna
                    email_detectado = None
                    for v in col_vals:
                        if not v:
                            continue
                        email_detectado = _extraer_email_de_texto(v) or _extraer_email_de_texto(_limpiar_posible_email(v))
                        if email_detectado:
                            break
                    if not email_detectado:
                        continue
                    columnas_con_email += 1
                    total_emails_col += 1
                    # mapear resto de celdas no-email en esta columna
                    for v in col_vals:
                        if not v:
                            continue
                        if _extraer_email_de_texto(v) or _extraer_email_de_texto(_limpiar_posible_email(v)):
                            continue
                        nombre_norm = _normalizar_nombre(v)
                        if not nombre_norm:
                            continue
                        if nombre_norm in mapeo_para_procesamiento and mapeo_para_procesamiento[nombre_norm] != email_detectado:
                            log_func(
                                f"ADVERTENCIA: Hoja '{nombre_hoja}', columna {col_idx+1}: Comunidad duplicada '{v}' (norm '{nombre_norm}') -> '{email_detectado}' (sobrescribiendo)."
                            )
                        mapeo_para_procesamiento[nombre_norm] = email_detectado
                log_func(f"Hoja '{nombre_hoja}': {columnas_con_email} columnas con email detectadas.")
            log_func(f"Total emails detectados por columnas: {total_emails_col}")

        log_func(f"Mapeo de correos cargado desde Excel: {len(mapeo_para_procesamiento)} comunidades mapeadas.")
        return mapeo_para_procesamiento

    except FileNotFoundError:
        log_func(f"ERROR CRÍTICO: Archivo Excel de mapeo NO ENCONTRADO en {RUTA_EXCEL_POR_DEFECTO}. Verifique la ruta y el nombre del archivo.")
        mapeo_para_procesamiento = {}
    except Exception as e:
        log_func(f"ERROR al leer el archivo Excel: {e}")
        mapeo_para_procesamiento = {}

def abrir_archivo(ruta_archivo):
    """Abre un archivo con la aplicación por defecto del sistema."""
    try:
        if sys.platform == "win32":
            os.startfile(os.path.abspath(ruta_archivo))
        else:
            webbrowser.open(f'file://{os.path.abspath(ruta_archivo)}')
    except Exception as e:
        print(f"No se pudo abrir el archivo {ruta_archivo}: {e}")

def run_gui():
    try:
        root = tk.Tk()
        root.title("Envío de Facturas por Correo")
        root.geometry("650x600")

        # --- 1. Variables de Estado y de la GUI ---
        procesando = False
        parada_evento = threading.Event()
        mapeo_comunidades_actual = {}
        progress_var = tk.DoubleVar()
        carpeta_var = tk.StringVar(root)
        remitente_var = tk.StringVar(root)
        gmail_pass_var = tk.StringVar(root)
        excel_file_var = tk.StringVar(root, value=RUTA_EXCEL_POR_DEFECTO)

        # --- 2. Definición de todas las Funciones Auxiliares ---

        def log_func_hilo(mensaje):
            def _log():
                if 'log_box' in globals() and log_box.winfo_exists():
                    log_box.config(state='normal')
                    log_box.insert(tk.END, mensaje + "\n")
                    log_box.config(state='disabled')
                    log_box.see(tk.END)
            root.after(0, _log)

        def actualizar_estado_botones(*args):
            # Esta función se define completamente después de crear los botones
            pass

        def finalizar_proceso():
            nonlocal procesando
            procesando = False
            parada_evento.clear()
            progress_var.set(0)
            actualizar_estado_botones()
            messagebox.showinfo("Proceso Finalizado", "El proceso de envío ha terminado. Revise los logs para ver el resumen.")

        def generar_reporte_final(exitosos, fallidos, saltados):
            log_func_hilo("\n--- RESUMEN DEL PROCESO ---")
            for tipo, lista in [("Enviados con éxito", exitosos), ("Errores de envío", fallidos), ("Facturas saltadas", saltados)]:
                log_func_hilo(f"\n>> {tipo}: {len(lista)}")
                for item in lista: log_func_hilo(f"  - {item}")
            
            try:
                if not os.path.exists('logs'): os.makedirs('logs')
                from datetime import datetime
                timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                ruta_reporte = os.path.join('logs', f'reporte_envio_{timestamp}.txt')
                with open(ruta_reporte, 'w', encoding='utf-8') as f:
                    f.write(f"--- RESUMEN DEL PROCESO ---\nFecha y Hora: {timestamp}\n")
                    for tipo, lista in [("Enviados con éxito", exitosos), ("Errores de envío", fallidos), ("Facturas saltadas", saltados)]:
                        f.write(f"\n>> {tipo}: {len(lista)}\n")
                        for item in lista: f.write(f"  - {item}\n")
                log_func_hilo(f"\nReporte guardado en: {ruta_reporte}")
                abrir_archivo(ruta_reporte)
            except Exception as e:
                log_func_hilo(f"ERROR al guardar el reporte: {e}")

        def proceso_de_envio_hilo(facturas_a_enviar, remitente, password, stop_event):
            enviados_exito, errores_envio, saltados_sin_correo = [], [], []
            try:
                log_func_hilo("--- INICIANDO PROCESO DE ENVÍO (confirmado por el usuario) ---")
                total_facturas = len(facturas_a_enviar)
                cuerpo_legal = "Este es un texto legal de prueba..."
                for i, f_info in enumerate(facturas_a_enviar):
                    if stop_event.is_set():
                        log_func_hilo("Proceso detenido por el usuario."); break
                    nombre_f, correo_d = os.path.basename(f_info['ruta_pdf']), f_info.get('correo_asignado')
                    if correo_d:
                        log_func_hilo(f"Intentando {i+1}/{total_facturas}: {nombre_f} -> {correo_d}")
                        asunto = f"Factura {f_info.get('numero_factura', 'N/A')} - {f_info.get('nombre_comunidad', 'N/A')}"
                        try:
                            exito = enviar_factura(remitente, password, correo_d, asunto, cuerpo_legal, f_info['ruta_pdf'])
                            if exito:
                                log_func_hilo(f"ÉXITO: {nombre_f} enviado.")
                                enviados_exito.append(f"{nombre_f} -> {correo_d}")
                            else:
                                raise Exception("Fallo SMTP reportado por enviar_factura")
                        except Exception as e:
                            log_func_hilo(f"ERROR al enviar {nombre_f}: {e}")
                            errores_envio.append(f"{nombre_f} -> {correo_d} (Error: {e})")
                    else:
                        log_func_hilo(f"SALTANDO: No hay correo para {nombre_f}")
                        saltados_sin_correo.append(nombre_f)
                    root.after(0, lambda p=(i + 1) * 100 / total_facturas: progress_var.set(p))
            except Exception as e:
                log_func_hilo(f"ERROR CRÍTICO en hilo de envío: {e}")
            finally:
                log_func_hilo("--- PROCESO DE ENVÍO FINALIZADO ---")
                generar_reporte_final(enviados_exito, errores_envio, saltados_sin_correo)
                root.after(0, finalizar_proceso)

        def iniciar_proceso_wrapper():
            nonlocal procesando
            if procesando: return messagebox.showwarning("Proceso en curso", "Ya hay un proceso en ejecución.")
            datos = {'carpeta': carpeta_var.get(), 'remitente': remitente_var.get(), 'pass': gmail_pass_var.get(), 'excel': excel_file_var.get()}
            if not all(datos.values()): return messagebox.showerror("Faltan datos", "Por favor, complete todos los campos antes de iniciar.")
            try:
                mapeo_comunidades_actual.clear()
                mapeo_comunidades_actual.update(cargar_mapeo_desde_excel(datos['excel'], log_func_hilo))
                if not mapeo_comunidades_actual: return log_func_hilo("No se pudo cargar el mapeo de correos. Proceso cancelado.")
            except Exception as e:
                return log_func_hilo(f"Error al cargar el archivo Excel: {e}")
            facturas_a_procesar = procesar_envios(datos['carpeta'], mapeo_comunidades_actual, log_func_hilo, lambda: False)
            if not facturas_a_procesar:
                messagebox.showinfo("Sin facturas", "No se encontraron facturas para procesar en la carpeta seleccionada.")
                return

            abrir_ventana_confirmacion(facturas_a_procesar, datos['remitente'], datos['pass'])

        def solicitar_parada_proceso():
            if procesando: log_func_hilo("Solicitando detención del proceso..."); parada_evento.set()

        def abrir_ventana_confirmacion(facturas, remitente, password):
            
            confirmation_window = tk.Toplevel(root)
            confirmation_window.title("Confirmar Envíos")
            confirmation_window.geometry("800x600")
            
            cols = ('Factura', 'Comunidad', 'Email Destino')
            tree = ttk.Treeview(confirmation_window, columns=cols, show='headings')
            for col in cols: tree.heading(col, text=col)
            tree.pack(side='left', fill='both', expand=True)

            factura_entries = {}
            for f in facturas:
                item_id = tree.insert("", "end", values=(
                    os.path.basename(f['ruta_pdf']),
                    f['nombre_comunidad'],
                    f.get('correo_asignado', '')
                ))
                factura_entries[item_id] = f

            def on_double_click(event):
                item_id = tree.selection()[0]
                column = tree.identify_column(event.x)
                if column == '#3': # Columna de Email
                    x, y, width, height = tree.bbox(item_id, column)
                    value = tree.set(item_id, column)
                    
                    entry = ttk.Entry(confirmation_window, width=width)
                    entry.place(x=x, y=y, width=width, height=height)
                    entry.insert(0, value)
                    entry.focus()
                    
                    def on_focus_out(event):
                        tree.set(item_id, column, entry.get())
                        entry.destroy()

                    entry.bind('<FocusOut>', on_focus_out)
                    entry.bind('<Return>', on_focus_out)

            tree.bind('<Double-1>', on_double_click)

            def confirmar_y_enviar():
                nonlocal procesando
                updated_facturas = []
                for item_id in tree.get_children():
                    valores = tree.item(item_id)['values']
                    factura_original = factura_entries[item_id]
                    factura_original['correo_asignado'] = valores[2]
                    updated_facturas.append(factura_original)
                
                procesando = True
                parada_evento.clear()
                actualizar_estado_botones()
                hilo = threading.Thread(target=proceso_de_envio_hilo, args=(updated_facturas, remitente, password, parada_evento))
                hilo.daemon = True
                hilo.start()
                confirmation_window.destroy()

            btn_frame = tk.Frame(confirmation_window)
            btn_frame.pack(side='right', fill='y', padx=10)
            tk.Button(btn_frame, text="Confirmar y Enviar", command=confirmar_y_enviar).pack(pady=10)
            tk.Button(btn_frame, text="Cancelar", command=confirmation_window.destroy).pack(pady=5)
            confirmation_window.transient(root)
            confirmation_window.grab_set()
            root.wait_window(confirmation_window)

        def browse_excel_file():
            f_path = filedialog.askopenfilename(title="Seleccionar archivo Excel", filetypes=(("Archivos de Excel", "*.xlsx"), ("Todos los archivos", "*.*")))
            if f_path: excel_file_var.set(f_path); log_func_hilo(f"Archivo de mapeo seleccionado: {f_path}")

        def buscar_facturas_numeradas():
            remitente_var.set("limpiezasmaylinsl@gmail.com")
            gmail_pass_var.set("lire faza liea aukc")
            log_func_hilo("Credenciales de Gmail autocompletadas.")
            carpeta = filedialog.askdirectory(title="Seleccionar carpeta de facturas")
            if not carpeta: return log_func_hilo("Selección de carpeta cancelada.")
            carpeta_var.set(carpeta)
            log_func_hilo(f"Carpeta de facturas seleccionada: {carpeta}")
            pdfs = [f for f in os.listdir(carpeta) if f.lower().endswith('.pdf')]
            if not pdfs: return messagebox.showinfo("Información", "No se encontraron archivos PDF en la carpeta seleccionada.")
            resumen = "Resumen de facturas encontradas:\n\n" + "\n".join(f"{i}. {f}" for i, f in enumerate(pdfs, 1))
            messagebox.showinfo("Facturas encontradas", resumen)

        # --- 3. Creación de la Interfaz Gráfica ---
        main_frame = tk.Frame(root, padx=10, pady=10); main_frame.pack(fill='both', expand=True)
        
        tk.Label(main_frame, text="Directorio de PDFs:").grid(row=0, column=0, sticky='w', pady=2)
        tk.Entry(main_frame, textvariable=carpeta_var, width=50).grid(row=0, column=1, padx=5, pady=2)
        tk.Button(main_frame, text="Buscar", command=lambda: carpeta_var.set(filedialog.askdirectory() or carpeta_var.get())).grid(row=0, column=2, padx=5, pady=2)
        tk.Label(main_frame, text="Remitente (Gmail):\n(Usa App Password)").grid(row=1, column=0, sticky='w', pady=2)
        tk.Entry(main_frame, textvariable=remitente_var, width=50).grid(row=1, column=1, padx=5, pady=2)
        tk.Label(main_frame, text="Contraseña App Gmail:").grid(row=2, column=0, sticky='w', pady=2)
        tk.Entry(main_frame, textvariable=gmail_pass_var, width=50, show="*").grid(row=2, column=1, padx=5, pady=2)
        tk.Label(main_frame, text="Archivo Excel de Mapeo:").grid(row=3, column=0, sticky='w', pady=2)
        tk.Entry(main_frame, textvariable=excel_file_var, width=50).grid(row=3, column=1, padx=5, pady=2)
        tk.Button(main_frame, text="Buscar Excel", command=browse_excel_file).grid(row=3, column=2, padx=5, pady=2)

        button_frame = tk.Frame(main_frame); button_frame.grid(row=4, column=0, columnspan=3, pady=10)
        autocompletar_button = tk.Button(button_frame, text="Autocompletar", command=buscar_facturas_numeradas); autocompletar_button.pack(side='left', padx=5)
        iniciar_button = tk.Button(button_frame, text="Iniciar Proceso", command=iniciar_proceso_wrapper, state='disabled'); iniciar_button.pack(side='left', padx=5)
        detener_button = tk.Button(button_frame, text="Detener Proceso", command=solicitar_parada_proceso, state='disabled'); detener_button.pack(side='left', padx=5)

        log_frame = tk.LabelFrame(main_frame, text="Logs del Proceso", padx=5, pady=5); log_frame.grid(row=5, column=0, columnspan=3, sticky='nsew', pady=10); main_frame.rowconfigure(5, weight=1)
        log_box = scrolledtext.ScrolledText(log_frame, width=70, height=15, state='disabled'); log_box.pack(fill='both', expand=True); globals()['log_box'] = log_box
        ttk.Progressbar(main_frame, variable=progress_var, maximum=100).grid(row=6, column=0, columnspan=3, sticky='ew', pady=(5,0))

        # --- 4. Lógica final de la GUI (Traces y actualización inicial) ---
        def _actualizar_estado_botones_final(*args):
            iniciar_button.config(state='normal' if all([carpeta_var.get(), remitente_var.get(), gmail_pass_var.get(), excel_file_var.get()]) and not procesando else 'disabled')
            detener_button.config(state='normal' if procesando else 'disabled')
        
        actualizar_estado_botones = _actualizar_estado_botones_final
        
        for var in [carpeta_var, remitente_var, gmail_pass_var, excel_file_var]:
            var.trace_add('write', actualizar_estado_botones)
        
        log_func_hilo("Aplicación GUI iniciada.")
        log_func_hilo("Por favor, configure los campos para habilitar el proceso.")
        actualizar_estado_botones()

        # Iniciar el bucle principal de la GUI
        root.mainloop()
    except Exception as e:
        import traceback
        error_msg = f"Error crítico en la GUI: {str(e)}\n{traceback.format_exc()}"
        print(error_msg)
        if 'log_func_hilo' in locals():
            log_func_hilo(error_msg)
        else:
            print("No se pudo registrar el error en el log de la GUI.")
        # Si la GUI ya se ha inicializado, mostrar un mensaje de error
        try:
            messagebox.showerror("Error Crítico", f"Ocurrió un error crítico: {str(e)}\n\nLa aplicación se cerrará.")
        except:
            pass  # Ignorar errores al mostrar el mensaje

if __name__ == "__main__":
    run_gui()
