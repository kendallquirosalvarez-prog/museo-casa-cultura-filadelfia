"""
Genera el codigo QR del Museo Virtual - Casa de la Cultura de Filadelfia.

Detecta automaticamente la IP local de la computadora en la red WiFi/Ethernet
actual y crea un codigo QR que apunta a http://<IP>:8000/

IMPORTANTE: la IP cambia segun la red WiFi a la que este conectada la
computadora. Por eso este script se debe volver a ejecutar:
  - Si cambias de red (ej. de tu casa al aula de la universidad)
  - Antes de cada presentacion, para asegurarte de que el QR sea correcto

Uso:
    python generar_qr.py

Requiere:
    pip install qrcode pillow   (ya estan instalados en este equipo)
"""
import socket
import qrcode
from pathlib import Path

PUERTO = 8000
CARPETA = Path(__file__).parent


def obtener_ip_local() -> str:
    """Obtiene la IP local usada para salir a la red (sin necesidad de internet real)."""
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
    except OSError:
        ip = "127.0.0.1"
    finally:
        s.close()
    return ip


def main():
    ip = obtener_ip_local()
    url = f"http://{ip}:{PUERTO}/"

    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=12,
        border=3,
    )
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="#4a3324", back_color="white")

    destino = CARPETA / "qr-museo.png"
    img.save(destino)

    print("=" * 60)
    print(" Museo Virtual - Casa de la Cultura de Filadelfia")
    print("=" * 60)
    print(f" IP local detectada : {ip}")
    print(f" URL del museo       : {url}")
    print(f" Codigo QR guardado  : {destino}")
    print()
    print(" Recuerda: para que el QR funcione, primero debes iniciar")
    print(" el servidor con 'iniciar-museo.bat' (o python -m http.server 8000)")
    print(" y los celulares deben estar conectados a la MISMA red WiFi")
    print(" que esta computadora.")
    print("=" * 60)


if __name__ == "__main__":
    main()
