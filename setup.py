from setuptools import setup

APP = ['code/gui.py']  # Cambia a tu archivo principal
DATA_FILES = []
OPTIONS = {
    'argv_emulation': True,
    'packages': ['PIL', 'openpyxl'],  # Solo los módulos necesarios
    'includes': ['tkinter']  # Incluye tkinter si usas una GUI
}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
    install_requires=['pillow',
        'openpyxl'
        # otras dependencias aquí...
    ],
)
