@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================================
echo  Museo Virtual - Casa de la Cultura de Filadelfia
echo ============================================================
echo.
echo  1. Generando codigo QR con la IP actual de esta red...
python generar_qr.py
echo.
echo  2. Abriendo el codigo QR y el museo en el navegador...
start "" "qr-museo.png"
start "" "http://localhost:8000/"
echo.
echo  3. Iniciando servidor local en el puerto 8000...
echo     (deja esta ventana abierta mientras dure la presentacion)
echo     Para detenerlo, cierra esta ventana o presiona Ctrl+C.
echo ============================================================
python -m http.server 8000
pause
