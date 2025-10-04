import tkinter as tk
from tkinter import messagebox

def main():
    try:
        # Create the main window
        root = tk.Tk()
        root.title("Test Tkinter")
        root.geometry("400x300")
        
        # Add a label
        label = tk.Label(root, text="¡Tkinter está funcionando correctamente!")
        label.pack(pady=20)
        
        # Add a button
        def show_message():
            messagebox.showinfo("Mensaje", "¡El botón funciona!")
            
        button = tk.Button(root, text="Haz clic aquí", command=show_message)
        button.pack(pady=10)
        
        # Run the application
        print("Aplicación iniciada correctamente.")
        root.mainloop()
        
    except Exception as e:
        print(f"Error: {e}")
        input("Presiona Enter para salir...")

if __name__ == "__main__":
    main()
