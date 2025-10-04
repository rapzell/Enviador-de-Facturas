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

        def mostrar_confirmacion(comunidades_envio):
            confirm_win = tk.Toplevel(win)
            confirm_win.title("Confirmar envío")
            confirm_win.geometry("500x300")
            
            tk.Label(confirm_win, text="¿Desea enviar los siguientes correos?").pack(pady=10)
            
            # Frame para la lista de correos
            frame_lista = tk.Frame(confirm_win)
            frame_lista.pack(fill='both', expand=True, padx=10, pady=5)
            
            # Crear un canvas con scrollbar
            canvas = tk.Canvas(frame_lista)
            scrollbar = ttk.Scrollbar(frame_lista, orient="vertical", command=canvas.yview)
            scrollable_frame = ttk.Frame(canvas)
            
            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(
                    scrollregion=canvas.bbox("all")
                )
            )
            
            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)
            
            # Mostrar lista de correos a enviar
            for i, com in enumerate(comunidades_envio, 1):
                tk.Label(scrollable_frame, 
                         text=f"{i}. {com['nombre']} -> {com['correo']}",
                         anchor='w').pack(fill='x', pady=2)
            
            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")
            
            # Botones de confirmación
            btn_frame = tk.Frame(confirm_win)
            btn_frame.pack(fill='x', padx=10, pady=10)
            
            def on_confirm():
                confirm_win.destroy()
                enviar_correos(comunidades_envio)
            
            ttk.Button(btn_frame, text="Cancelar", command=confirm_win.destroy).pack(side='right', padx=5)
            ttk.Button(btn_frame, text="Confirmar envío", command=on_confirm).pack(side='right', padx=5)
            
            confirm_win.transient(win)
            confirm_win.grab_set()
            win.wait_window(confirm_win)
        
        def enviar_correos(comunidades_envio):
            try:
                # Configurar los parámetros del correo
                asunto = "PRUEBA"
                mensaje = "Esto es una prueba"
                
                # Obtener credenciales de la ventana principal
                remitente = remitente_var.get()
                password = gmail_pass_var.get()
                
                if not remitente or not password:
                    messagebox.showerror("Error", "Faltan credenciales de correo. Por favor, complete los campos de correo y contraseña en la ventana principal.")
                    return
                
                # Enviar cada correo
                for com in comunidades_envio:
                    try:
                        # Llamar a la función de envío
                        procesar_envios(
                            archivos_pdf=[com['pdf']],
                            destinatario=com['correo'],
                            asunto=asunto,
                            mensaje=mensaje,
                            remitente=remitente,
                            password=password
                        )
                        messagebox.showinfo("Éxito", f"Correo enviado correctamente a {com['correo']}")
                    except Exception as e:
                        messagebox.showerror("Error", f"Error al enviar a {com['correo']}: {str(e)}")
                        
            except Exception as e:
                messagebox.showerror("Error", f"Error inesperado: {str(e)}")
        
        def guardar_y_continuar():
            comunidades_actualizadas = []
            for i, com in enumerate(comunidades):
                if check_vars[i].get() and entry_vars[i].get():
                    com['correo'] = entry_vars[i].get()
                    comunidades_actualizadas.append(com)
            if not comunidades_actualizadas:
                messagebox.showwarning("Advertencia", "No hay comunidades seleccionadas para enviar")
                return
            mostrar_confirmacion(comunidades_actualizadas)

        # Botón inferior
        btn_guardar = ttk.Button(win, text='Guardar y continuar', command=guardar_y_continuar)
        btn_guardar.pack(pady=10)

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
