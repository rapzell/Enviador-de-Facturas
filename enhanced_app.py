import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import sys
import os
import traceback

class EnhancedApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Envío de Facturas (Versión Mejorada)")
        self.root.geometry("1000x700")
        
        # Variables de estado
        self.procesando = False
        self.detener_proceso = False
        
        # Configuración de estilo
        self.setup_styles()
        
        # Inicializar la interfaz
        self.setup_ui()
        
        # Configurar manejo de cierre
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self.log("Aplicación iniciada correctamente.")
    
    def setup_styles(self):
        """Configura los estilos de la interfaz"""
        style = ttk.Style()
        style.configure('TButton', padding=5)
        style.configure('TFrame', padding=5)
        style.configure('Header.TLabel', font=('Helvetica', 12, 'bold'))
    
    def setup_ui(self):
        """Configura los elementos de la interfaz de usuario"""
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Título
        ttk.Label(main_frame, text="Envío de Facturas", style='Header.TLabel').pack(pady=10)
        
        # Frame de controles superiores
        ctrl_frame = ttk.Frame(main_frame)
        ctrl_frame.pack(fill=tk.X, pady=5)
        
        # Botón para seleccionar directorio
        ttk.Button(ctrl_frame, text="Seleccionar directorio", 
                  command=self.seleccionar_directorio).pack(side=tk.LEFT, padx=5)
        
        # Botón para buscar facturas
        ttk.Button(ctrl_frame, text="Buscar facturas", 
                  command=self.buscar_facturas).pack(side=tk.LEFT, padx=5)
        
        # Área de logs
        log_frame = ttk.LabelFrame(main_frame, text="Registro", padding="5")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.log_area = scrolledtext.ScrolledText(log_frame, height=20)
        self.log_area.pack(fill=tk.BOTH, expand=True)
        
        # Frame de botones inferiores
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=5)
        
        # Botón para detener proceso
        self.stop_btn = ttk.Button(btn_frame, text="Detener", 
                                 command=self.detener_proceso, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)
        
        # Botón para salir
        ttk.Button(btn_frame, text="Salir", 
                  command=self.on_closing).pack(side=tk.RIGHT, padx=5)
    
    def log(self, message):
        """Agrega un mensaje al área de logs"""
        self.log_area.insert(tk.END, f"{message}\n")
        self.log_area.see(tk.END)
        self.root.update_idletasks()
    
    def seleccionar_directorio(self):
        """Abre un diálogo para seleccionar un directorio"""
        try:
            directorio = filedialog.askdirectory(title="Seleccionar directorio de facturas")
            if directorio:
                self.log(f"Directorio seleccionado: {directorio}")
                # Aquí podrías guardar el directorio seleccionado
        except Exception as e:
            self.log(f"Error al seleccionar directorio: {e}")
    
    def buscar_facturas(self):
        """Simula la búsqueda de facturas"""
        if self.procesando:
            return
            
        self.procesando = True
        self.detener_proceso = False
        self.stop_btn.config(state=tk.NORMAL)
        
        try:
            self.log("Buscando facturas...")
            # Simular un proceso largo
            for i in range(1, 11):
                if self.detener_proceso:
                    self.log("Búsqueda detenida por el usuario.")
                    break
                    
                self.log(f"Procesando factura {i}/10")
                self.root.after(500)  # Pequeña pausa
                self.root.update()
            
            if not self.detener_proceso:
                self.log("Búsqueda completada.")
                
        except Exception as e:
            self.log(f"Error durante la búsqueda: {e}")
        finally:
            self.procesando = False
            self.stop_btn.config(state=tk.DISABLED)
    
    def detener_proceso(self):
        """Detiene el proceso en ejecución"""
        self.detener_proceso = True
        self.log("Solicitando detener el proceso...")
    
    def on_closing(self):
        """Maneja el cierre de la aplicación"""
        if self.procesando:
            if messagebox.askyesno("Proceso en ejecución", 
                                 "Hay un proceso en ejecución. ¿Desea forzar la salida?"):
                self.detener_proceso = True
                self.root.quit()
        else:
            if messagebox.askokcancel("Salir", "¿Está seguro de que desea salir?"):
                self.log("Cerrando la aplicación...")
                self.root.quit()

def main():
    print("=== Iniciando aplicación mejorada ===")
    
    try:
        root = tk.Tk()
        app = EnhancedApp(root)
        root.mainloop()
        print("Aplicación cerrada correctamente.")
    except Exception as e:
        print(f"Error inesperado: {e}")
        print(traceback.format_exc())
        messagebox.showerror("Error", f"Se produjo un error inesperado: {e}")
    finally:
        if 'root' in locals():
            try:
                root.quit()
            except:
                pass

if __name__ == "__main__":
    main()
