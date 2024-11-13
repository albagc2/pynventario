# 📋 Generador de Inventarios

Hola! Este programa te permite convertir carpetas de imágenes en documentos estructurados de inventario en formato Word, ideal para tareas como reclamaciones de seguros, catalogación de objetos, o simplemente para mantener todo organizado.

### 🚀 Funcionalidades principales

- **Jerarquía respetada**: Genera un documento que refleja la estructura de carpetas y subcarpetas de tus imágenes.
- **Imágenes directamente en el documento**: Incluye las imágenes dentro del archivo Word, redimensionadas automáticamente para ajustarse al diseño.
- **Soporte HEIC**: Procesa imágenes en formato HEIC (usado frecuentemente por dispositivos Apple) gracias a la integración con `pillow-heif`.
- **Exportación de tabla de contenidos**: Genera un archivo Excel con un índice de todos los elementos incluidos en el inventario.
- **Interfaz gráfica amigable**: Diseñada para usuarios sin experiencia en programación.

---

### 💻 Cómo usar el programa en Windows

1. **Descarga y ejecuta el archivo `mi inventario.exe`**:
   - Abre la interfaz gráfica desde el archivo ejecutable `mi inventario.exe`.

2. **Selecciona la carpeta de imágenes**:
   - Usa el botón **"Seleccionar carpeta"** para elegir dónde están las imágenes y carpetas que deseas procesar.

3. **Completa los datos necesarios**:
   - Ingresa información como:
     - Nombre del archivo de inventario
     - Localidad
     - Nombre y DNI de la persona
     - Referencia del expediente

4. **Opciones adicionales**:
   - Marca la casilla **"Añadir imágenes directamente al documento"** si quieres que las fotos aparezcan en el documento.

5. **Genera tu inventario**:
   - Haz clic en **"Generar Inventario"** y espera unos momentos.

6. **Resultados**:
   - El archivo generado estará en la carpeta `documentos/` dentro de la misma carpeta donde tienes el programa.

---

### ⚙️ Requisitos

- **Sistema operativo**: Windows
- **Formato de entrada**: Imágenes en `.jpg`, `.png`, `.bmp`, `.gif`, `.jfif`, y `.heic`.

> **Nota**: La implementación para macOS está en camino. 😊

---

### 🛠️ Solución de problemas

- **¿El programa no abre?** Asegúrate de ejecutar `mi inventario.exe` en un ordenador con Windows y de tener permisos suficientes.
- **¿Errores con imágenes?** Consulta el archivo `app_log.txt` que se genera automáticamente para depurar problemas.
- **¿Formato HEIC no soportado?** Este soporte está incluido de forma predeterminada, pero asegúrate de que los archivos no estén corruptos.

---

Espero que este programa haga tu trabajo más fácil y organizado. Soy consciente de que este programa tiene muchas limitaciones, y no dudo en que habrá muchas cosas que mejorar, así que no dudes en reportar cualquier sugerencia o problema en la página de Issues del repositorio!

Ojalá que os sirva. Mucho ánimo :heartbeat: 
