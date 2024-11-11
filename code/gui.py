import tkinter as tk
from tkinter import filedialog, messagebox
from pynvent import crear_inventario  

def run_inventario():
    # Validar que todos los campos estén completos
    if not root_folder.get() or not output_doc.get():
        messagebox.showerror("Error", "Por favor, introduce al menos la carpeta con las imágenes y el nombre del documento.")
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
            create_doc=True,
            add_images=add_images.get() 
        )
        messagebox.showinfo("Éxito", "Inventario generado exitosamente.\nGuardado en pynventario/documentos.")
        window.destroy()  # Cierra la ventana principal
    except Exception as e:
        messagebox.showerror("Error", f"Ha ocurrido un error: {e}")

    except Exception as e:
        messagebox.showerror("Error", f"Ha ocurrido un error: {e}")

# Configurar la ventana principal
window = tk.Tk()
window.title("Generador de Inventario de Daños")
window.geometry("400x500")

# Variables para almacenar los datos
root_folder = tk.StringVar()
output_doc = tk.StringVar()
add_images = tk.BooleanVar()
location = tk.StringVar()
name = tk.StringVar()
dni = tk.StringVar()
ref_expediente = tk.StringVar()

# Crear los campos de entrada y botones
tk.Label(window, text="Carpeta con las imágenes:").pack(pady=5)
tk.Entry(window, textvariable=root_folder, width=40).pack()
tk.Button(window, text="Seleccionar carpeta", command=lambda: root_folder.set(filedialog.askdirectory())).pack()

tk.Label(window, text="Nombre del archivo con el inventario (.docx guardado en documentos/):").pack(pady=5)
tk.Entry(window, textvariable=output_doc, width=40).pack()

tk.Checkbutton(window, text="Añadir imágenes directamente al documento", variable=add_images).pack(pady=5)

tk.Label(window, text="Localidad:").pack(pady=5)
tk.Entry(window, textvariable=location, width=40).pack()

tk.Label(window, text="Nombre de la persona:").pack(pady=5)
tk.Entry(window, textvariable=name, width=40).pack()

tk.Label(window, text="DNI de la persona:").pack(pady=5)
tk.Entry(window, textvariable=dni, width=40).pack()

tk.Label(window, text="Referencia de expediente:").pack(pady=5)
tk.Entry(window, textvariable=ref_expediente, width=40).pack()

#tk.Checkbutton(window, text="Renombrar imágenes usando Google Vision", variable=rename_images).pack(pady=5)

tk.Button(window, text="Generar Inventario", command=run_inventario).pack(pady=20)

# Ejecutar la interfaz
window.mainloop()
