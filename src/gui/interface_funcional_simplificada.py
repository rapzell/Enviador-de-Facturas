import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import threading
import os
import sys
import traceback

# Variables globales para controlar el estado
procesando = False
detener_proceso_solicitado = False
mapeo_comunidades_actual = {}

def cargar_mapeo_desde_excel(log_func):
    """Simula la carga de mapeo desde Excel"""
    log_func("Cargando mapeo de correos (simulado)...")
    # En una versión real, aquí cargarías el Excel
    return {"Comunidad 1": "correo1@ejemplo.com", "Comunidad 2": "correo2@ejemplo.com"}

def run_gui():
    """Función principal para iniciar la GUI"""
    global procesando, detener_proceso_solicitado, mapeo_comunidades_actual
    
    try:
        print('INICIANDO GUI SIMPLIFICADA')
        
        # Crear instancia principal de Tkinter
        root = tk.Tk()
        root.title("Envío de Facturas a Comunidades - VERSIÓN SIMPLIFICADA")
        root.geometry("600x600")
        root.minsize(500, 500)
        
        # Variables Tkinter
        carpeta_var = tk.StringVar(root)
        remitente_var = tk.StringVar(root)
        gmail_pass_var = tk.StringVar(root)
        
        # Frame principal
        main_frame = tk.Frame(root, padx=10, pady=10)
        main_frame.pack(fill='both', expand=True)
        
        # Frame para selección de carpeta
        frame_img = tk.Frame(main_frame)
        frame_img.pack(fill='x', pady=5)
        
        tk.Label(frame_img, text="Carpeta de facturas:").pack(side='left')
        tk.Entry(frame_img, textvariable=carpeta_var, width=40).pack(side='left', padx=5, expand=True, fill='x')
        
        def browse_folder():
            d = filedialog.askdirectory()
            if d:
                carpeta_var.set(d)
        
        tk.Button(frame_img, text="Examinar", command=browse_folder).pack(side='left')
        
        # Frame para credenciales
        cred_frame = tk.Frame(main_frame)
        cred_frame.pack(fill='x', pady=5)
        
        tk.Label(cred_frame, text="Correo Gmail:").pack(anchor='w')
        tk.Entry(cred_frame, textvariable=remitente_var, width=40).pack(fill='x', padx=2)
        
        tk.Label(cred_frame, text="Contraseña de aplicación:").pack(anchor='w')
        tk.Entry(cred_frame, textvariable=gmail_pass_var, show='*', width=40).pack(fill='x')
        
        # Área de logs
        log_frame = tk.Frame(main_frame)
        log_frame.pack(fill='both', expand=True, pady=10)
        
        tk.Label(log_frame, text="Registro de operaciones:").pack(anchor='w')
        log_box = scrolledtext.ScrolledText(log_frame, height=10)
        log_box.pack(fill='both', expand=True, padx=2, pady=2)
        log_box.config(state='disabled')
        
        def log_func_hilo(mensaje):
            """Función para manejar logs desde hilos secundarios"""
            log_box.config(state='normal')
            log_box.insert(tk.END, f"{mensaje}\n")
            log_box.see(tk.END)
            log_box.config(state='disabled')
            root.update_idletasks()
        
        # Carga inicial de mapeo
        try:
            mapeo_comunidades_actual = cargar_mapeo_desde_excel(log_func_hilo)
            log_func_hilo(f"Mapeo cargado correctamente. {len(mapeo_comunidades_actual)} comunidades mapeadas.")
        except Exception as e_excel:
            log_func_hilo(f"ERROR al cargar mapeo desde Excel: {str(e_excel)}")
        
        # Función para detener el proceso
        def detener_proceso():
            global detener_proceso_solicitado
            detener_proceso_solicitado = True
            log_func_hilo("Se ha solicitado detener el proceso...")
            btn_parar.config(state=tk.DISABLED)
        
        # Función para simular el procesamiento
        def iniciar_proceso_wrapper():
            global procesando, detener_proceso_solicitado
            
            if not carpeta_var.get().strip():
                messagebox.showerror("Error", "Debe seleccionar un directorio válido")
                return
                
            procesando = True
            detener_proceso_solicitado = False
            
            # Actualizar estado de botones
            btn_ejecutar.config(state=tk.DISABLED)
            btn_parar.config(state=tk.NORMAL)
            
            log_func_hilo(f"Iniciando procesamiento en: {carpeta_var.get()}")
            log_func_hilo(f"Usando cuenta: {remitente_var.get()}")
            
            # Simular proceso
            for i in range(1, 6):
                if detener_proceso_solicitado:
                    log_func_hilo("Proceso detenido por el usuario")
                    break
                    
                log_func_hilo(f"Procesando factura {i}...")
                root.after(1000)  # Simulación de proceso
            
            # Restaurar estado
            procesando = False
            btn_ejecutar.config(state=tk.NORMAL)
            btn_parar.config(state=tk.DISABLED)
            log_func_hilo("Procesamiento finalizado")
        
        # Función para iniciar en un hilo separado
        def procesar_en_hilo_wrapper():
            if procesando:
                messagebox.showinfo("Aviso", "Ya hay un proceso en ejecución")
                return
            threading.Thread(target=iniciar_proceso_wrapper, daemon=True).start()
        
        # Frame de botones
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill='x', pady=5)
        
        # Botones
        btn_ejecutar = tk.Button(button_frame, text="Iniciar Proceso", command=procesar_en_hilo_wrapper,
                               bg='#4CAF50', fg='white')
        btn_ejecutar.pack(side='left', padx=5)
        
        btn_parar = tk.Button(button_frame, text="Detener Proceso", command=detener_proceso,
                            state=tk.DISABLED, bg='#F44336', fg='white')
        btn_parar.pack(side='left', padx=5)
        
        # Función para autocompletar datos de demostración
        def autocompletar_demo():
            carpeta_var.set(r'C:\Users\David\Desktop\FACTURAS')
            remitente_var.set('ejemplo@gmail.com')
            log_func_hilo("Datos de demostración cargados")
            
        tk.Button(button_frame, text="Demo", command=autocompletar_demo).pack(side='right', padx=5)
        
        log_func_hilo("Aplicación iniciada correctamente")
        log_func_hilo("Seleccione una carpeta con facturas para comenzar")
        
        # Iniciar bucle principal de Tkinter
        root.protocol("WM_DELETE_WINDOW", lambda: sys.exit())  # Para forzar salida al cerrar
        root.mainloop()
        
    except Exception as e:
        print(f"ERROR CRÍTICO: {e}")
        traceback.print_exc()
        messagebox.showerror("Error", f"Error crítico: {e}")

if __name__ == "__main__":
    print("Iniciando aplicación simplificada")
    run_gui()
    print("Aplicación finalizada")
