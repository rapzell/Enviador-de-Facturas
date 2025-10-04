import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import sys
import os
import traceback

def main():
    print("=== Iniciando versión simplificada de la aplicación ===")
    
    try:
        # Crear la ventana principal
        root = tk.Tk()
        root.title("Envío de Facturas (Versión Simple)")
        root.geometry("800x600")
        
        # Frame principal
        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Título
        ttk.Label(main_frame, text="Envío de Facturas", font=('Helvetica', 16, 'bold')).pack(pady=10)
        
        # Área de logs
        log_frame = ttk.LabelFrame(main_frame, text="Registro", padding="5")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        log_area = scrolledtext.ScrolledText(log_frame, height=15)
        log_area.pack(fill=tk.BOTH, expand=True)
        
        def log_message(message):
            log_area.insert(tk.END, message + "\n")
            log_area.see(tk.END)
        
        # Botón de prueba
        def on_test_click():
            log_message("Botón de prueba presionado")
            messagebox.showinfo("Prueba", "¡La aplicación está funcionando correctamente!")
        
        test_btn = ttk.Button(main_frame, text="Probar aplicación", command=on_test_click)
        test_btn.pack(pady=10)
        
        # Manejo de cierre
        def on_closing():
            if messagebox.askokcancel("Salir", "¿Está seguro de que desea salir?"):
                log_message("Cerrando la aplicación...")
                root.quit()
        
        root.protocol("WM_DELETE_WINDOW", on_closing)
        
        log_message("Aplicación iniciada correctamente.")
        print("Aplicación iniciada en modo gráfico.")
        
        # Iniciar el bucle principal
        root.mainloop()
        
    except Exception as e:
        error_msg = f"Error inesperado: {str(e)}\n\n{traceback.format_exc()}"
        print(error_msg)
        if 'root' in locals() and hasattr(root, 'winfo_exists') and root.winfo_exists():
            messagebox.showerror("Error", "Se produjo un error inesperado.")
    finally:
        print("Aplicación finalizada.")

if __name__ == "__main__":
    main()
