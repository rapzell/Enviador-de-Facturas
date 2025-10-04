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
    
    import glob
    archivos = glob.glob(os.path.join(carpeta, '*.pdf')) + glob.glob(os.path.join(carpeta, 'factura_*.pdf'))
    
    if not archivos:
        messagebox.showinfo("Facturas", "No se encontraron facturas numeradas en la carpeta")
        return
    
    try:
        from src.extractor_comunidad_pdf import extraer_comunidad_de_pdf
    except ImportError:
        from extractor_comunidad_pdf import extraer_comunidad_de_pdf
    resumen = f"Se encontraron {len(archivos)} facturas:\n"
    # This function in the backup version doesn't directly feed abrir_asignacion_correos.
    # The procesar_en_hilo function does. So, this part is mostly for user info.
    for i, archivo in enumerate(sorted(archivos), 1):
        comunidad = extraer_comunidad_de_pdf(archivo)
        resumen += f"{i}. {os.path.basename(archivo)} | Comunidad: {comunidad if comunidad else '[NO DETECTADA]'}\n"
    messagebox.showinfo("Facturas encontradas", resumen)

def run_gui():
    import traceback
    from tkinter import messagebox
    import threading
    
    procesando = False
    detener_proceso = False
    
    def detener_analisis():
        nonlocal detener_proceso
        detener_proceso = True
        btn_parar.config(state='disabled')
        log_box.insert(tk.END, "\nDeteniendo análisis...\n")

    def mostrar_confirmacion(comunidades_a_enviar):
        confirm_win = tk.Toplevel(root)
        confirm_win.title("Confirmar Envío")
        confirm_win.geometry("600x450")
        
        tk.Label(confirm_win, text="Se enviarán correos a las siguientes comunidades:", font=('Arial', 12, 'bold')).pack(pady=10)
        
        text_frame = tk.Frame(confirm_win)
        text_frame.pack(pady=5, padx=10, fill='both', expand=True)
        
        list_text = scrolledtext.ScrolledText(text_frame, height=15, width=70)
        list_text.pack(side='left', fill='both', expand=True)
        
        scrollbar = ttk.Scrollbar(text_frame, orient='vertical', command=list_text.yview)
        scrollbar.pack(side='right', fill='y')
        list_text['yscrollcommand'] = scrollbar.set
        
        list_text.config(state='normal')
        for com in comunidades_a_enviar:
            list_text.insert(tk.END, f"Comunidad: {com['nombre']}\n")
            list_text.insert(tk.END, f"Correo: {com['correo']}\n")
            list_text.insert(tk.END, f"Archivo: {os.path.basename(com['pdf'])}\n")
            list_text.insert(tk.END, "----------------------------------------\n")
        list_text.config(state='disabled')

        def _enviar_final():
            confirm_win.destroy()
            log_box.insert(tk.END, "Iniciando proceso de envío real...\n")
            
            # Configuración SMTP para Gmail
            smtp_config = {
                'host': 'smtp.gmail.com',
                'port': 587,
                'user': remitente_var.get(),
                'password': gmail_pass_var.get(),
                'use_tls': True
            }
            
            try:
                # Llamar a procesar_envios con los datos necesarios
                resultados = procesar_envios(
                    None,  # imagen_tabla no es necesaria en el flujo de comunidades
                    remitente_var.get(),
                    smtp_config,
                    lambda msg: log_box.insert(tk.END, msg + '\n'),
                    comunidades_a_enviar
                )
                
                # Mostrar resumen de envíos
                enviados = sum(1 for r in resultados if r[2] == 'Enviado')
                errores = sum(1 for r in resultados if r[2].startswith('Error'))
                
                messagebox.showinfo("Proceso Completado", 
                    f"Se han procesado {len(resultados)} facturas.\n"
                    f"Enviados: {enviados}\n"
                    f"Errores: {errores}")
                
            except Exception as e:
                log_box.insert(tk.END, f"Error inesperado: {str(e)}\n")
                messagebox.showerror("Error", f"Ocurrió un error al enviar los correos: {str(e)}")

        btn_frame_confirm = tk.Frame(confirm_win)
        btn_frame_confirm.pack(pady=10)

        ttk.Button(btn_frame_confirm, text="Enviar Correos", command=_enviar_final).pack(side='left', padx=10)
        ttk.Button(btn_frame_confirm, text="Cancelar", command=confirm_win.destroy).pack(side='right', padx=10)

        confirm_win.transient(root)
        confirm_win.grab_set()
        root.wait_window(confirm_win)
    
    def abrir_asignacion_correos(comunidades):
        win = tk.Toplevel(root)
        win.title("Asignación de Correos")
        win.geometry("1000x600")
        
        frame_tabla = tk.Frame(win)
        frame_tabla.pack(fill='both', expand=True, padx=10, pady=10)
        
        tabla = tk.Frame(frame_tabla)
        tabla.pack(fill='both', expand=True)
        tk.Label(tabla, text='Enviar', width=8, anchor='center', font=('Arial', 9, 'bold')).grid(row=0, column=0, padx=2, pady=2)
        tk.Label(tabla, text='Comunidad', width=32, anchor='w', font=('Arial', 9, 'bold')).grid(row=0, column=1, padx=2, pady=2)
        tk.Label(tabla, text='Correo', width=36, anchor='w', font=('Arial', 9, 'bold')).grid(row=0, column=2, padx=2, pady=2)
        tk.Label(tabla, text='Archivo', width=14, anchor='center', font=('Arial', 9, 'bold')).grid(row=0, column=3, padx=2, pady=2)
        tk.Label(tabla, text='', width=8).grid(row=0, column=4)

        check_vars = []
        comunidad_nombre_vars = [] 
        correo_vars = [] 
        
        for i, com_data in enumerate(comunidades):
            var_envio = tk.BooleanVar(value=com_data.get('enviar', True))
            var_comunidad_nombre = tk.StringVar(value=com_data.get('nombre', ''))
            var_correo_email = tk.StringVar(value=com_data.get('correo', 'davidvr1994@gmail.com'))
            
            check_vars.append(var_envio)
            comunidad_nombre_vars.append(var_comunidad_nombre)
            correo_vars.append(var_correo_email)
            
            cb = tk.Checkbutton(tabla, variable=var_envio)
            cb.grid(row=i+1, column=0, padx=4, sticky='w')
            
            entry_comunidad = tk.Entry(tabla, textvariable=var_comunidad_nombre, width=32)
            entry_comunidad.grid(row=i+1, column=1, padx=2, sticky='we')
            
            entry_correo = tk.Entry(tabla, textvariable=var_correo_email, state='normal', width=36)
            entry_correo.grid(row=i+1, column=2, padx=2, sticky='we')
            
            tabla.grid_columnconfigure(1, weight=1)
            tabla.grid_columnconfigure(2, weight=1)
            
            tk.Label(tabla, text=os.path.basename(com_data['pdf']), width=14, anchor='center').grid(row=i+1, column=3)
            btn = tk.Button(tabla, text='Abrir', command=lambda f=com_data['pdf']: abrir_archivo(f))
            btn.grid(row=i+1, column=4, padx=2)

        btn_guardar = ttk.Button(win, text='Guardar y continuar', command=lambda: guardar_y_continuar_wrapper()) # Wrapped to pass correct vars
        btn_guardar.pack(pady=10)

        def guardar_y_continuar_wrapper(): # Wrapper to correctly capture loop variables for guardar_y_continuar
            nonlocal comunidades, check_vars, comunidad_nombre_vars, correo_vars, win
            comunidades_actualizadas = []
            for i, com_data_original in enumerate(comunidades):
                if check_vars[i].get() and correo_vars[i].get(): 
                    com_data_original['nombre'] = comunidad_nombre_vars[i].get()
                    com_data_original['correo'] = correo_vars[i].get()
                    com_data_original['enviar'] = True # Mark as selected for sending
                    comunidades_actualizadas.append(com_data_original)
                else:
                    com_data_original['enviar'] = False # Mark as not selected
            
            # Filter again to be sure, though appending logic should handle it
            final_comunidades_para_enviar = [c for c in comunidades_actualizadas if c.get('enviar')]

            if not final_comunidades_para_enviar:
                messagebox.showwarning("Advertencia", "No hay comunidades seleccionadas para enviar o falta algún correo.")
                return
            
            win.destroy()
            mostrar_confirmacion(final_comunidades_para_enviar)

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
        
        top_frame = tk.Frame(root)
        top_frame.pack(fill='x', padx=5, pady=5)
        # btn_buscar (command=buscar_facturas_numeradas) is mostly for info in this version
        # The main flow starts with procesar_en_hilo triggered by btn_ejecutar
        tk.Button(top_frame, text=" Info Facturas (Carpeta)", command=buscar_facturas_numeradas).pack(side='right', padx=5)

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
            # Ruta actualizada para el botón de autocompletar demo
            carpeta_var.set(r'C:\Users\David\Desktop\PROGRAMAR\ENVIADOR DE FACTURAS\FACTURAS\FACTURAS COMPLETA\05 - Mayo - 2025')
            mes_var.set('mayo')
            remitente_var.set('davidvr1994@gmail.com')
            gmail_pass_var.set('') # No poner la contraseña real aquí
        tk.Button(button_frame, text="Autocompletar demo", command=autocompletar_demo, bg='#2196F3', fg='white').pack(side='left', padx=5)
        
        def procesar_en_hilo_wrapper(): # Wrapper to avoid issues with procesando/detener_proceso scope if run_gui is called multiple times
            nonlocal procesando, detener_proceso
            if procesando:
                messagebox.showinfo("Info", "Ya hay un proceso en ejecución.")
                return
            
            procesando = True
            detener_proceso = False
            
            try:
                carpeta = carpeta_var.get()
                mes = mes_var.get()
                remitente = remitente_var.get()
                gmail_pass = gmail_pass_var.get()
                
                if not carpeta or not mes or not remitente: # Gmail pass can be empty for testing GUI
                    messagebox.showerror("Error", "Por favor complete Carpeta, Mes y Correo Remitente.")
                    procesando = False
                    return
                
                root.after(0, lambda: [
                    btn_parar.config(state='normal'),
                    btn_ejecutar.config(state='disabled')
                ])
                
                log_box.delete(1.0, tk.END)
                log_box.insert(tk.END, "Iniciando análisis de facturas...\n")
                root.update_idletasks()
                
                log_box.insert(tk.END, "Buscando facturas en la carpeta...\n")
                root.update_idletasks()
                
                import glob
                archivos = glob.glob(os.path.join(carpeta, '*.pdf')) + glob.glob(os.path.join(carpeta, 'factura_*.pdf'))
                
                if detener_proceso: procesando = False; return
                if not archivos:
                    log_box.insert(tk.END, "No se encontraron archivos PDF en la carpeta.\n")
                    messagebox.showinfo("Info", "No se encontraron facturas PDF en la carpeta especificada.")
                    procesando = False
                    root.after(0, lambda: [btn_ejecutar.config(state='normal'), btn_parar.config(state='disabled')])
                    return
                
                log_box.insert(tk.END, f"Se encontraron {len(archivos)} archivos PDF.\n")
                log_box.insert(tk.END, "Extrayendo nombres de comunidades de los PDFs...\n")
                root.update_idletasks()
                
                try:
                    from src.extractor_comunidad_pdf import extraer_comunidad_de_pdf
                except ImportError:
                    from extractor_comunidad_pdf import extraer_comunidad_de_pdf
                
                comunidades_data = []
                # Verificar si el proceso fue detenido incluso antes de comenzar la extracción
                if not detener_proceso:
                    for idx, archivo_pdf in enumerate(sorted(archivos)):
                        if detener_proceso:
                            log_box.insert(tk.END, "Proceso detenido durante la extracción de comunidades.\n")
                            root.update_idletasks()
                            break # Salir del bucle, pero continuar para verificar si hay datos parciales
                        nombre_comunidad = extraer_comunidad_de_pdf(archivo_pdf)
                        if not nombre_comunidad:
                            nombre_comunidad = f"Comunidad Desconocida {idx+1}"
                        comunidades_data.append({'nombre': nombre_comunidad, 'pdf': archivo_pdf, 'correo': 'davidvr1994@gmail.com', 'enviar': True})
                        log_box.insert(tk.END, f"  {os.path.basename(archivo_pdf)} -> {nombre_comunidad}\n")
                        root.update_idletasks()
                
                # Después del bucle (o si fue omitido/interrumpido por 'detener_proceso')
                # Abrir la ventana de asignación si tenemos algún dato.
                if comunidades_data:
                    log_box.insert(tk.END, "Abriendo ventana de asignación de correos...\n")
                    root.update_idletasks()
                    # Pasar una copia de la lista para evitar problemas si el original se modifica más tarde
                    root.after(0, lambda: abrir_asignacion_correos(list(comunidades_data)))
                elif detener_proceso: # Detenido y no se recopilaron datos
                    log_box.insert(tk.END, "Proceso detenido. No hay datos de comunidades para mostrar.\n")
                    root.update_idletasks()
                # Si no es detener_proceso y no hay comunidades_data, significa que no se encontraron PDFs,
                # lo cual se maneja antes en la función procesar_en_hilo_wrapper.
                
            except Exception as e_thread:
                log_box.insert(tk.END, f"Error en el procesamiento: {str(e_thread)}\n")
                messagebox.showerror("Error en Proceso", str(e_thread))
                print(traceback.format_exc())
            finally:
                procesando = False
                root.after(0, lambda: [
                    btn_ejecutar.config(state='normal'),
                    btn_parar.config(state='disabled')
                ])

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
