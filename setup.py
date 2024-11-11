from setuptools import setup

APP = ['code/gui.py']  # Cambia a tu archivo principal
DATA_FILES = []
OPTIONS = {
    'argv_emulation': True,
    'packages': ['pillow', 'google.cloud.vision', 'openpyxl'],  # Solo los módulos necesarios
    'includes': ['tkinter']  # Incluye tkinter si usas una GUI
}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
    install_requires=[
        'google-cloud-vision',
        'pillow',
        'poenpyxl'
        # otras dependencias aquí...
    ],
)
