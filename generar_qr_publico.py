"""
Genera el codigo QR FIJO del Museo Virtual - Casa de la Cultura de Filadelfia,
apuntando a la URL publica y permanente de GitHub Pages.

A diferencia de generar_qr.py (que apunta a la IP local y cambia segun la red),
este QR ya NO necesita regenerarse nunca: la URL es siempre la misma y
funciona desde cualquier red con internet (datos moviles, otra WiFi, etc.),
sin que tu computadora tenga que estar prendida.

Uso:
    python generar_qr_publico.py
"""
import qrcode
from pathlib import Path

URL_PUBLICA = "https://kendallquirosalvarez-prog.github.io/museo-casa-cultura-filadelfia/"
CARPETA = Path(__file__).parent


def main():
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=12,
        border=3,
    )
    qr.add_data(URL_PUBLICA)
    qr.make(fit=True)
    img = qr.make_image(fill_color="#4a3324", back_color="white")

    destino = CARPETA / "qr-museo.png"
    img.save(destino)

    print("=" * 60)
    print(" Museo Virtual - Casa de la Cultura de Filadelfia")
    print("=" * 60)
    print(f" URL publica (fija)  : {URL_PUBLICA}")
    print(f" Codigo QR guardado  : {destino}")
    print()
    print(" Este QR ya no cambia. Funciona desde cualquier red con")
    print(" internet, sin que tu computadora tenga que estar prendida.")
    print("=" * 60)


if __name__ == "__main__":
    main()
