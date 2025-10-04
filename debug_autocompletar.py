import tkinter as tk
from tkinter import messagebox
import os

def autocompletar_demo():
    try:
        print("Función autocompletar_demo() llamada")  # Debug
        carpeta = r'C:\Users\David\Desktop\PROGRAMAR\ENVIADOR DE FACTURAS\FACTURAS\FACTURAS COMPLETA\05 - Mayo - 2025'
        print(f"Intentando establecer carpeta: {carpeta}")  # Debug
        carpeta_var.set(carpeta)
        mes_var.set('mayo')
        remitente_var.set('davidvr1994@gmail.com')
        gmail_pass_var.set('')
        print("Valores establecidos correctamente")  # Debug
    except Exception as e:
        print(f"Error en autocompletar_demo: {str(e)}")  # Debug
        messagebox.showerror("Error", f"Error en autocompletar_demo: {str(e)}")

# Crear ventana principal
root = tk.Tk()
root.title("Prueba de autocompletar_demo")

# Variables
try_width = 50
carpeta_var = tk.StringVar()
mes_var = tk.StringVar()
remitente_var = tk.StringVar()
gmail_pass_var = tk.StringVar()

# Interfaz
tk.Label(root, text="Carpeta de facturas:").pack()
tk.Entry(root, textvariable=carpeta_var, width=entry_width).pack()

tk.Label(root, text="Mes:").pack()
tk.Entry(root, textvariable=mes_var, width=entry_width).pack()

tk.Label(root, text="Correo Gmail remitente:").pack()
tk.Entry(root, textvariable=remitente_var, width=entry_width).pack()

tk.Label(root, text="Contraseña de aplicación Gmail:").pack()
tk.Entry(root, textvariable=gmail_pass_var, show='*', width=entry_width).pack()

# Botón de prueba
tk.Button(root, text="Probar autocompletar_demo", command=autocompletar_demo, bg='#2196F3', fg='white').pack(pady=20)

# Botón para mostrar valores actuales
def mostrar_valores():
    valores = f"""
    Carpeta: {carpeta_var.get()}
    Mes: {mes_var.get()}
    Remitente: {remitente_var.get()}
    Contraseña: {'*' * len(gmail_pass_var.get())}
    """
    messagebox.showinfo("Valores actuales", valores)

tk.Button(root, text="Mostrar valores actuales", command=mostrar_valores).pack()

root.mainloop()
