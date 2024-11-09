import sys 
sys.path.append('.')
import argparse
from pynvent import crear_inventario

# Configuración de argumentos desde la línea de comandos
parser = argparse.ArgumentParser(description="Generador de inventario de imágenes y tabla de contenidos.")
parser.add_argument("root_folder", type=str, help="Ruta de la carpeta raíz que contiene las imágenes.")
parser.add_argument("output_doc", type=str, help="Nombre del archivo de inventario de salida.")
parser.add_argument("location", type=str, help="Nombre de la población.")
parser.add_argument("name", type=str, help="Nombre de la persona.")
parser.add_argument("dni", type=str, help="DNI de la persona.")
parser.add_argument("ref_expediente", type=str, help="Referencia del expediente.")
parser.add_argument("--scale_factor", type=float, default=0.5, help="Factor de escala para redimensionar imágenes.")
parser.add_argument("--rename_images", action="store_true", help="Renombra las imágenes usando Google Vision.")
parser.add_argument("--create_doc", action="store_true", help="Genera el documento de inventario.")

args = parser.parse_args()

# Llamada a la función principal
crear_inventario(
    root_folder=args.root_folder,
    output_doc=args.output_doc,
    location=args.location,
    name=args.name,
    dni=args.dni,
    ref_expediente=args.ref_expediente,
    scale_factor=args.scale_factor,
    rename_images=args.rename_images,
    create_doc=args.create_doc
)
