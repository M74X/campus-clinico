"""Pega el sticker QR generado sobre el banner de referencia.

Toma el banner institucional (referencia de diseño), localiza el marco
azul del QR placeholder por color (banda inferior central), y lo
reemplaza por el sticker con la URL de producción. El sticker recibe
máscara redondeada para conservar el fondo teal-striped en esquinas.

Seguridad de datos: solo usa el banner y la URL ya embebida en el
sticker; no toca datos de alumnos.

Uso:
    python scripts/pegar_qr_banner.py [banner] [sticker] [salida]
Defaults:
    banner  = primer banner*.png/jpg del proyecto (sin ser salidas)
    sticker = img/campo-clinico-banner.png
    salida  = img/banner-campo-clinico-2026ii.png

Dependencias: pillow, zxing-cpp
"""
import sys
from pathlib import Path

import zxingcpp
from PIL import Image, ImageDraw

BASE = Path(__file__).resolve().parent.parent
AZUL = (16, 100, 180)     # RGB aprox del marco azul del banner
TOL = (60, 70, 80)        # tolerancia de color del detector

STICKER_DEF = BASE / "img" / "campo-clinico-banner.png"
SALIDA_DEF = BASE / "img" / "banner-campo-clinico-2026ii.png"

# Caja del placeholder en la referencia (medida pixel a pixel, 1600×900):
# marco azul externo (680,585)-(903,786). Fija, no auto-detectada: el
# auto-detect abarcó píxeles azules ajenos y desplazó el sticker sobre
# el texto 'INGRESA'. Margen 0: el sticker ya trae su propio marco.
CAJA_QR = (680, 585, 903, 786)


def buscar_banner():
    if len(sys.argv) > 1:
        return Path(sys.argv[1])
    for raiz in (BASE, BASE / "img"):
        for suf in ("*.png", "*.jpg", "*.jpeg"):
            for h in sorted(raiz.glob(suf)):
                if h.name.lower().startswith("banner") and \
                   "campo-clinico" not in h.name:
                    return h
    return None


def es_azul(r, g, b):
    return abs(r - AZUL[0]) < TOL[0] and abs(g - AZUL[1]) < TOL[1] \
        and abs(b - AZUL[2]) < TOL[2]


def detectar_qr(banner_rgb):
    """BBox del marco azul del QR placeholder.

    Banda acotada del banner de referencia (1600x900): el QR vive en
    x∈[0.40,0.58]W, y∈[0.62,1.0]H — fuera quedan el título azul, el
    texto 'Medicina Humana' y el pill amarillo 'AQUÍ' (falsos azules).
    """
    W, H = banner_rgb.size
    x0r, x1r = int(W * 0.40), int(W * 0.58)
    y0r, y1r = int(H * 0.62), H
    px = banner_rgb.load()
    xs, ys = [], []
    for y in range(y0r, y1r, 2):
        for x in range(x0r, x1r, 2):
            if es_azul(*px[x, y]):
                xs.append(x)
                ys.append(y)
    if len(xs) < 50:
        raise SystemExit("No hallé el marco azul del QR en la referencia "
                         "(¿cambió el diseño de fondo?).")
    return min(xs), min(ys), max(xs), max(ys)


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

    ruta_banner = buscar_banner()
    if not ruta_banner or not ruta_banner.exists():
        raise SystemExit("No encontré el banner de referencia (pásalo como argumento).")
    sticker_path = Path(sys.argv[2]) if len(sys.argv) > 2 else STICKER_DEF
    salida = Path(sys.argv[3]) if len(sys.argv) > 3 else SALIDA_DEF

    banner = Image.open(ruta_banner).convert("RGB")
    sticker = Image.open(sticker_path).convert("RGBA")
    caja = CAJA_QR if banner.size == (1600, 900) else detectar_qr(banner)
    bx0, by0, bx1, by1 = caja
    cx, cy = (bx0 + bx1) // 2, (by0 + by1) // 2

    # Cubrir exactamente el marco viejo (sin margen: el sticker trae marco)
    lado_v = max(bx1 - bx0, by1 - by0)
    px0 = (bx0 + bx1) // 2 - lado_v // 2
    py0 = (by0 + by1) // 2 - lado_v // 2

    s = sticker.resize((lado_v, lado_v), Image.Resampling.LANCZOS)
    mask = Image.new("L", (lado_v, lado_v), 0)
    ImageDraw.Draw(mask).rounded_rectangle(
        [0, 0, lado_v - 1, lado_v - 1], radius=int(lado_v * 0.10), fill=255)
    banner.paste(s, (px0, py0), mask)
    banner.save(salida, dpi=(150, 150))

    # Verificación: el recorte del banner pegado debe dar el mismo contenido
    def lee(im):
        r = zxingcpp.read_barcodes(im)
        return r[0].text if r else None

    zona = banner.crop((px0, py0, px0 + lado_v, py0 + lado_v))
    igual = lee(zona) == lee(sticker.convert("RGB"))
    full = lee(banner) == lee(sticker.convert("RGB"))
    print(f"Banner: {ruta_banner.name} {banner.size[0]}x{banner.size[1]} "
          f"| placeholder {caja}")
    print(f"Sticker pegado: {lado_v}x{lado_v} px @ ({px0},{py0}) → {salida.name}")
    print(f"  ✓ recorte escanea igual que el sticker" if igual
          else "  ✘ recorte distinto: revisar caja")
    print(f"  ✓ banner completo escanea" if full
          else "  ⚠ banner completo no escanea directo (normal si el diseño "
               "corta en WhatsApp de móvil: lo manda con recote visual aparte)")

    print("Fin — solo imágenes; cero datos de alumnos.")


if __name__ == "__main__":
    main()
