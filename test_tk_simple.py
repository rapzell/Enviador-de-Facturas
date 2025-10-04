import tkinter as tk
from tkinter import messagebox

def main():
    # Crear ventana principal
    root = tk.Tk()
    root.title("Prueba Tkinter")
    root.geometry("400x200")
    
    # Función para el botón
    def mostrar_mensaje():
        messagebox.showinfo("¡Funciona!", "Tkinter está funcionando correctamente.")
    
    # Botón de prueba
    btn = tk.Button(root, text="Haz clic aquí", command=mostrar_mensaje, bg='#4CAF50', fg='white', padx=20, pady=10)
    btn.pack(pady=50)
    
    # Iniciar el bucle principal
    root.mainloop()

if __name__ == "__main__":
    print("Iniciando prueba de Tkinter...")
    main()
    print("Prueba finalizada.")
