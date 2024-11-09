@echo off
REM Ejecución de main.py con argumentos
set PYTHONPATH=%CD%
python-3.13.0-embed-amd64\python.exe main.py %1 %2 %3 %4 %5 %6 --scale_factor 0.5 --create_doc

@echo off
set GOOGLE_APPLICATION_CREDENTIALS="C:\ruta\al\archivo\de\clave.json"
set PATH=%cd%;%cd%\Scripts;%SystemRoot%\system32;%SystemRoot%;%SystemRoot%\System32\Wbem;%SystemRoot%\System32\WindowsPowerShell\v1.0\
python.exe main.py %1 %2 %3 %4 %5 %6 --rename_images
