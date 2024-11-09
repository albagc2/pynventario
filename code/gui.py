import tkinter as tk
from tkinter import filedialog, messagebox
from pynvent import crear_inventario  

def run_inventario():
    # Validar que todos los campos estén completos
    if not root_folder.get() or not output_doc.get() or not location.get() or not name.get() or not dni.get() or not ref_expediente.get():
        messagebox.showerror("Error", "Por favor completa todos los campos.")
        return

    # Llamar a la función principal
    try:
        crear_inventario(
            root_folder=root_folder.get(),
            output_doc=output_doc.get(),
            location=location.get(),
            name=name.get(),
            dni=dni.get(),
            ref_expediente=ref_expediente.get(),
            scale_factor=0.5,
            rename_images=rename_images.get(),
            create_doc=True
        )
        messagebox.showinfo("Éxito", "Inventario generado exitosamente.")
    except Exception as e:
        messagebox.showerror("Error", f"Ha ocurrido un error: {e}")

# Configurar la ventana principal
window = tk.Tk()
window.title("Generador de Inventario de Daños")
window.geometry("400x500")

# Variables para almacenar los datos
root_folder = tk.StringVar()
output_doc = tk.StringVar()
location = tk.StringVar()
name = tk.StringVar()
dni = tk.StringVar()
ref_expediente = tk.StringVar()
rename_images = tk.BooleanVar()

# Crear los campos de entrada y botones
tk.Label(window, text="Carpeta raíz de imágenes:").pack(pady=5)
tk.Entry(window, textvariable=root_folder, width=40).pack()
tk.Button(window, text="Seleccionar carpeta", command=lambda: root_folder.set(filedialog.askdirectory())).pack()

tk.Label(window, text="Archivo de salida (.docx):").pack(pady=5)
tk.Entry(window, textvariable=output_doc, width=40).pack()
tk.Button(window, text="Seleccionar archivo", command=lambda: output_doc.set(filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Document", "*.docx")]))).pack()

tk.Label(window, text="Ubicación:").pack(pady=5)
tk.Entry(window, textvariable=location, width=40).pack()

tk.Label(window, text="Nombre de la persona:").pack(pady=5)
tk.Entry(window, textvariable=name, width=40).pack()

tk.Label(window, text="DNI de la persona:").pack(pady=5)
tk.Entry(window, textvariable=dni, width=40).pack()

tk.Label(window, text="Referencia de expediente:").pack(pady=5)
tk.Entry(window, textvariable=ref_expediente, width=40).pack()

tk.Checkbutton(window, text="Renombrar imágenes usando Google Vision", variable=rename_images).pack(pady=5)

tk.Button(window, text="Generar Inventario", command=run_inventario).pack(pady=20)

# Ejecutar la interfaz
window.mainloop()
