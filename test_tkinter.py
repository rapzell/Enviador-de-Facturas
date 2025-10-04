import tkinter as tk
from tkinter import messagebox

root = tk.Tk()
root.title("Prueba de Tkinter")
root.geometry("400x200")

label = tk.Label(root, text="¡Si ves esto, Tkinter funciona correctamente!")
label.pack(pady=20)

def on_click():
    messagebox.showinfo("¡Éxito!", "El botón funciona correctamente")

button = tk.Button(root, text="Haz clic aquí", command=on_click)
button.pack()

root.mainloop()
