import tkinter as tk
print('>>> CARGADO src/gui/interface.py')
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
    messagebox.showinfo("Info", "Los correos se han enviado correctamente")
    # Puedes guardar comunidades_pdf en variable global o pasarla al flujo de asignación

def run_gui():
    import traceback
    from tkinter import messagebox
    import threading
    import sys
    
    # Variables globales para controlar el hilo
    procesando = False
    detener_proceso = False
    root = None
    
    try:
        # Crear la ventana principal
        root = tk.Tk()
        root.title("Envío de Facturas")
        root.geometry("800x600")
        
        # Variable para controlar el estado de la aplicación
        app_running = True
        
        # Función para cerrar la aplicación de manera segura
        def on_closing():
            nonlocal app_running
            app_running = False
            if root is not None:
                try:
                    if hasattr(root, 'destroy'):
                        root.quit()
                        root.destroy()
                except Exception as e:
                    print(f"Error al cerrar la ventana: {e}")
        
        # Configurar el manejador de cierre de la ventana
        root.protocol("WM_DELETE_WINDOW", on_closing)
        
        # Variables globales para controlar el hilo
        procesando = False
        detener_proceso = False
        
        def detener_analisis():
            nonlocal detener_proceso
            detener_proceso = True
            btn_parar.config(state='disabled')
            log_box.insert(tk.END, "\nDeteniendo análisis...\n")

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

        def browse_folder():
            d = filedialog.askdirectory()
            if d:
                carpeta_var.set(d)

        def autocompletar_demo():
            carpeta_var.set(r'C:\\Users\\David\\Desktop\\PROGRAMAR\\ENVIADOR DE FACTURAS\\FACTURAS\\FACTURAS COMPLETA\\05 - Mayo - 2025')
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
                        comunidades.append({'pdf': archivo, 'nombre': comunidad, 'correo': ''})
                        log_box.insert(tk.END, f"Procesada: {os.path.basename(archivo)} - {comunidad}\n")
                    else:
                        log_box.insert(tk.END, f"[ADVERTENCIA] No se pudo extraer comunidad de: {os.path.basename(archivo)}\n")
                    root.update_idletasks()
                    
                if not comunidades:
                    messagebox.showerror("Error", "No se pudo extraer información de ninguna factura")
                    return
                    
                # Siempre abrir la ventana de asignación si hay alguna comunidad, aunque se haya detenido el análisis
                abrir_asignacion_correos(comunidades)
                
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
        
        # Botón de parar (inicialmente deshabilitado)
        btn_parar = tk.Button(button_frame, text="[PARAR] Análisis", command=detener_analisis, 
                            bg='#f44336', fg='white', state='disabled')
        btn_parar.pack(side='left', padx=5)
        
        progress_frame = tk.Frame(main_frame)
        progress = ttk.Progressbar(progress_frame, mode='indeterminate', length=200)
        progress_label = tk.Label(progress_frame, text="Procesando, por favor espera...")
        log_box = scrolledtext.ScrolledText(main_frame, height=9)
        log_box.pack(fill='both', padx=2, pady=2, expand=True)
        
        # Ejecutar el bucle principal de la interfaz
        try:
            root.mainloop()
        except KeyboardInterrupt:
            print("\nAplicación interrumpida por el usuario")
        except Exception as e:
            print(f"Error en el bucle principal: {e}")
        finally:
            on_closing()
        
    except Exception as e:
        error_msg = f"Error inesperado: {str(e)}\n\n{traceback.format_exc()}"
        print(error_msg, file=sys.stderr)
        if root is not None:
            try:
                root.quit()
            except:
                pass
    finally:
        # Limpieza final
        if root is not None:
            try:
                root.destroy()
            except:
                pass
    
    # --- FUNCIONES INTERNAS ---
    def abrir_asignacion_correos(comunidades):
        win = tk.Toplevel(root)
        win.title("Asignación de Correos")
        win.geometry("1000x600")
        
        frame_tabla = tk.Frame(win)
        frame_tabla.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Cabecera y filas en un solo Frame usando grid
        tabla = tk.Frame(frame_tabla)
        tabla.pack(fill='both', expand=True)
        
        # Encabezados de la tabla
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
            
            # Crear fila de la tabla
            cb = tk.Checkbutton(tabla, variable=var_envio)
            cb.grid(row=i+1, column=0, padx=4, sticky='w')
            
            # Mostrar nombre de la comunidad
            tk.Label(tabla, text=com['nombre'], width=32, anchor='w').grid(
                row=i+1, column=1, sticky='w')
                
            # Campo de entrada para el correo
            entry = tk.Entry(tabla, textvariable=var_correo, state='normal')
            entry.grid(row=i+1, column=2, padx=2, sticky='we')
            tabla.grid_columnconfigure(2, weight=1)
            
            # Mostrar nombre del archivo
            tk.Label(tabla, text=os.path.basename(com['pdf']), 
                    width=14, anchor='center').grid(row=i+1, column=3)
                    
            # Botón para abrir el archivo
            btn = tk.Button(tabla, text='Abrir', 
                          command=lambda f=com['pdf']: abrir_archivo(f))
            btn.grid(row=i+1, column=4, padx=2)

        def mostrar_confirmacion(comunidades_envio):
            confirm_win = tk.Toplevel(win)
            confirm_win.title("Confirmar envio de correos")
            msg = "Se enviaran {} correos. Desea continuar?".format(len(comunidades_envio))
            tk.Label(confirm_win, text=msg, font=('Arial', 10, 'bold')).pack(padx=20, pady=20)
            
            btn_frame = tk.Frame(confirm_win)
            btn_frame.pack(pady=10)
            
            btn_enviar = ttk.Button(
                btn_frame, 
                text="Enviar correos", 
                command=lambda: [confirm_win.destroy(), enviar_correos(comunidades_envio)],
                style='Accent.TButton'
            )
            btn_enviar.pack(side='right', padx=5)
            
            btn_cancelar = ttk.Button(
                btn_frame, 
                text="Cancelar", 
                command=confirm_win.destroy,
                style='TButton'
            )
            btn_cancelar.pack(side='right', padx=5)
            
            confirm_win.transient(win)
            confirm_win.grab_set()
            win.wait_window(confirm_win)
        
        def enviar_correos(comunidades_envio):
            try:
                import yagmail
                import os
                
                print("\n=== INICIANDO ENVÍO DE CORREOS ===")
                print(f"Total de correos a enviar: {len(comunidades_envio)}")
                
                # Obtener credenciales de la ventana principal
                remitente = remitente_var.get().strip()
                password = gmail_pass_var.get().strip()
                
                print(f"Remitente: {remitente}")
                print("Contraseña:", "*" * (len(password) if password else 0))
                
                if not remitente or not password:
                    error_msg = "Faltan credenciales de correo. Por favor, complete los campos de correo y contraseña."
                    print(f"[ERROR] {error_msg}")
                    messagebox.showerror("Error", error_msg)
                    return
                
                # Verificar que los archivos PDF existan
                for com in comunidades_envio:
                    if not os.path.exists(com['pdf']):
                        error_msg = f"El archivo no existe: {com['pdf']}"
                        print(f"[ERROR] {error_msg}")
                        messagebox.showerror("Error", error_msg)
                        return
                
                # Configurar yagmail
                yag = yagmail.SMTP(remitente, password)
                
                print("\n=== DETALLES DE ENVÍO ===")
                for i, com in enumerate(comunidades_envio, 1):
                    print(f"{i}. {com['nombre']} -> {com['correo']}")
                    print(f"   PDF: {com['pdf']}")
                
                # Enviar cada correo
                exitos = 0
                errores = 0
                
                for com in comunidades_envio:
                    try:
                        destinatario = com['correo'].strip()
                        asunto = f"Factura - {com['nombre']}"
                        mensaje = f"""
                        <h2>Estimado/a,</h2>
                        <p>Adjunto encontrará la factura correspondiente.</p>
                        <p>Saludos cordiales,<br>Equipo de Facturación</p>
                        """
                        
                        print(f"\nEnviando a {destinatario}...")
                        
                        # Enviar el correo con yagmail
                        yag.send(
                            to=destinatario,
                            subject=asunto,
                            contents=mensaje,
                            attachments=com['pdf']
                        )
                        
                        print(f"  Enviado correctamente a {destinatario}")
                        exitos += 1
                        
                    except Exception as e:
                        error_msg = str(e)
                        print(f"  Error al enviar a {com['correo']}: {error_msg}")
                        errores += 1
                
                # Mostrar resumen
                print("\n=== RESUMEN DE ENVÍOS ===")
                print(f"Total de correos: {len(comunidades_envio)}")
                print(f"Enviados correctamente: {exitos}")
                print(f"Errores: {errores}")
                
                if exitos > 0 and errores == 0:
                    msg = f"Se enviaron {exitos} correos correctamente."
                    messagebox.showinfo("Éxito", msg)
                elif exitos > 0 and errores > 0:
                    msg = f"Se enviaron {exitos} correos correctamente.\nNo se pudieron enviar {errores} correos."
                    messagebox.showwarning("Enviado parcialmente", msg)
                else:
                    messagebox.showerror("Error", "No se pudo enviar ningún correo. Verifique sus credenciales o conexión a Internet.")
                
            except Exception as e:
                import traceback
                tb = traceback.format_exc()
                error_msg = f"Error inesperado al enviar correos:\n{str(e)}"
                print(f"[ERROR CRITICO] {error_msg}")
                print(f"Traceback:\n{tb}")
                messagebox.showerror("Error", error_msg)
        
        # Botón de confirmación
        btn_confirmar = tk.Button(
            win, 
            text="Confirmar y enviar correos",
            command=lambda: mostrar_confirmacion([
                {'nombre': com['nombre'], 'correo': var.get(), 'pdf': com['pdf']} 
                for com, var in zip(comunidades, entry_vars) 
                if var.get()
            ]),
            bg='#4CAF50', 
            fg='white',
            font=('Arial', 10, 'bold')
        )
        btn_confirmar.pack(pady=10)
        
        # Botón para cancelar
        btn_cancelar = tk.Button(
            win, 
            text="Cancelar",
            command=win.destroy,
            bg='#f44336', 
            fg='white'
        )
        btn_cancelar.pack(pady=5)
        
        # Hacer que la ventana sea modal
        win.transient(root)
        win.grab_set()
        root.wait_window(win)
            
# --- FIN ENVÍO DE CORREOS Y FLUJO PRINCIPAL ---

