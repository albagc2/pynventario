# %%
import sys
sys.path.append("C:/Users/albag/Desktop/autoinventario")

import pynvent



# %%
# Llama a la función principal
root_folder = "C:/Users/albag/Desktop/Fotos/Inventario casa nombres"
output_doc = "inventario_sin_renombrar.docx"
location = "SPRINGFIELD\nCALLE FALSA 1,2,3"
name = "HOMER J. SIMPSON"
dni = "4815162342"
ref_expediente = "XXXXXXXXXX"

# Ejecutar la función principal del módulo
pynvent.crear_inventario(
    root_folder=root_folder,
    output_doc=output_doc,
    location=location,
    rename_images=False,
    create_doc=True,
    name=name,
    dni=dni,
    ref_expediente=ref_expediente

)



