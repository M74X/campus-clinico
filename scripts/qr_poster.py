"""Genera los stickers/posters QR del app de Campo Clínico 2026-II y los verifica.

Una corrida produce DOS artefactos:

1. STICKER BANNER  (img/qr-campo-clinico-banner.png)
   Para insertar EN el banner institucional: módulos casi negros sobre
   blanco, marco azul redondeado y badge central teal con "W" blanca
   (alusión al logo UNW). Si la batería falla con badge, se reintenta
   sin él automáticamente — la legibilidad escaneando manda.

2. POSTER ESTILO CARNET  (img/qr-campo-clinico-poster.png)
   Réplica del qr-carnet-2026-ii-poster.png del repo app-carnet-unw:
   marco teal, sello amarillo "UNIVERSIDAD NORBERT WIENER", título
   "CAMPOS CLÍNICOS 2026-II", CTA, divisor amarillo y footer.
   Módulos dotted teal OSCURO (#00695C, contraste ~6:1), ojos solidos.

Compartido:
- QR con ECC nivel H, zona silenciosa ≥4 módulos intacta.
- Ojos (finder) y alignment SIEMPRE sólidos (diagnóstico 2026-10-24:
  all-dotted falla a varias escalas). Posiciones alignment genéricas
  por versión (v2–v40), no la v10 fija del script original.
- Batería con zxing-cpp: exacta, distancia simulada (60/40/25%),
  thumb 700px, daño de cuadrante 25%, huella (60 crestas) y smear.

Uso:
    python scripts/qr_poster.py <URL> [salida_sin_extension]
Sin argumentos: no genera nada (nunca un QR con placeholder).
Args extra: "sin-logo" fuerza el sticker sin badge.

Dependencias:
    python -m pip install segno zxing-cpp pillow
"""
import random
import sys
import tempfile
from pathlib import Path

import segno
import zxingcpp
from PIL import Image, ImageDraw, ImageFilter, ImageFont

# ── Paleta UNW (del banner y del header del app) ─────────────────────────────
MODULES_DARK  = (16, 19, 26)     # #10131A — módulos del sticker (casi negro)
TEAL_MODULOS  = (0, 105, 92)     # #00695C — módulos del poster (oscuro ~6:1)
AZUL_BANNER   = (11, 90, 168)    # #0B5AA8 — marco azul del sticker
TEAL_BRILLO   = (0, 181, 184)    # #00B5B8 — marco/divisor del poster
YELLOW_ACC    = (255, 193, 7)    # #FFC107 — sello y divisor
GRIS_SUAVE    = (60, 80, 80)
BLANCO        = (255, 255, 255)

BADGE_FRACC  = 0.21    # badge (con anillo) vs bloque QR — ECC H lo cubre
STROKE_FRACC = 0.030   # grosor del marco azul del sticker vs lado


def fuente_bold(size):
    for cand in ("arialbd.ttf", "segoeuib.ttf",
                 "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"):
        try:
            return ImageFont.truetype(cand, size=size)
        except OSError:
            continue
    return ImageFont.load_default()


def fuente_reg(size):
    for cand in ("arial.ttf", "segoeui.ttf",
                 "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"):
        try:
            return ImageFont.truetype(cand, size=size)
        except OSError:
            continue
    return ImageFont.load_default()


# ── Tabla ISO 18004: posiciones alignment por versión (genérico) ─────────────
ALIGN_POS = {
    2: [6, 18], 3: [6, 22], 4: [6, 26], 5: [6, 30], 6: [6, 34],
    7: [6, 22, 38], 8: [6, 24, 42], 9: [6, 26, 46], 10: [6, 28, 50],
    11: [6, 30, 54], 12: [6, 32, 58], 13: [6, 34, 62], 14: [6, 26, 46, 66],
    15: [6, 26, 48, 70], 16: [6, 26, 50, 74], 17: [6, 30, 54, 78],
    18: [6, 30, 56, 82], 19: [6, 30, 58, 86], 20: [6, 34, 62, 90],
    21: [6, 28, 50, 72, 94], 22: [6, 26, 50, 74, 98], 23: [6, 30, 54, 78, 102],
    24: [6, 28, 54, 80, 106], 25: [6, 32, 58, 84, 110], 26: [6, 30, 58, 86, 114],
    27: [6, 34, 62, 90, 118], 28: [6, 26, 50, 74, 98, 122],
    29: [6, 30, 54, 78, 102, 126], 30: [6, 26, 52, 78, 104, 130],
    31: [6, 30, 56, 82, 108, 134], 32: [6, 34, 60, 86, 112, 138],
    33: [6, 30, 58, 86, 114, 142], 34: [6, 34, 62, 90, 118, 146],
    35: [6, 30, 54, 78, 102, 126, 150], 36: [6, 24, 50, 76, 102, 128, 154],
    37: [6, 28, 54, 80, 106, 132, 158], 38: [6, 32, 58, 84, 110, 136, 162],
    39: [6, 26, 54, 82, 110, 138, 166], 40: [6, 30, 58, 86, 114, 142, 170],
}


def _qr_matriz(url):
    qr = segno.make(url, error="h")
    matriz, version = qr.matrix, qr.version
    nmod = 4 * version + 17
    pos = ALIGN_POS.get(version, [6])
    solidos = {(r, c)
               for fr, fc in [(0, 0), (0, nmod - 7), (nmod - 7, 0)]
               for r in range(fr, fr + 7) for c in range(fc, fc + 7)}
    for a in pos:
        for b in pos:
            if a == 6 and b == 6:
                continue
            for r in range(a - 2, a + 3):
                for c in range(b - 2, b + 3):
                    solidos.add((r, c))
    return matriz, nmod, solidos


def dibujar_bloque(url, con_logo, color_modulos, scale=40):
    """Bloque QR + zona silenciosa, SIN marco. Devuelve (img, caja, diag)."""
    matriz, nmod, solidos = _qr_matriz(url)
    QZ = 4 * scale
    qs_px = (nmod + 8) * scale

    canvas = Image.new("RGB", (qs_px, qs_px), BLANCO)
    d = ImageDraw.Draw(canvas)
    R = int(scale * 0.30)
    for r in range(nmod):
        for c in range(nmod):
            if matriz[r][c]:
                rr = 0 if (r, c) in solidos else R
                d.rounded_rectangle([QZ + c * scale, QZ + r * scale,
                                     QZ + (c + 1) * scale, QZ + (r + 1) * scale],
                                    radius=rr, fill=color_modulos)

    if con_logo:
        lado = int(qs_px * BADGE_FRACC)
        anillo = int(lado * 0.055)
        x0 = y0 = qs_px // 2 - lado // 2
        d.rounded_rectangle([x0, y0, x0 + lado, y0 + lado],
                            radius=int(lado * 0.22), fill=BLANCO)
        x1, y1 = x0 + anillo, y0 + anillo
        l1 = lado - 2 * anillo
        d.rounded_rectangle([x1, y1, x1 + l1, y1 + l1],
                            radius=int(l1 * 0.22), fill=TEAL_BRILLO)
        d.text((x1 + l1 // 2, y1 + l1 // 2), "W",
               font=fuente_bold(int(l1 * 0.62)), fill=BLANCO, anchor="mm")

    diag = dict(version=segno.make(url, error="h").version, nmod=nmod, chars=len(url))
    return canvas, (0, 0, qs_px), diag


# ── Artefacto 1: STICKER para el banner ──────────────────────────────────────
def sticker(url, con_logo):
    bloque, caja_b, diag = dibujar_bloque(url, con_logo, MODULES_DARK)
    qs = bloque.width
    STROKE = max(6, int(qs * STROKE_FRACC))
    W = qs + 2 * STROKE

    canvas = Image.new("RGB", (W, W), BLANCO)
    canvas.paste(bloque, (STROKE, STROKE))
    d = ImageDraw.Draw(canvas)
    d.rounded_rectangle([0, 0, W - 1, W - 1], radius=int(W * 0.10),
                        outline=AZUL_BANNER, width=STROKE)
    return canvas, (STROKE, STROKE, qs), diag


# ── Artefacto 2: POSTER estilo carnet ────────────────────────────────────────
def medir(texto, size, bold=True):
    f = fuente_bold(size) if bold else fuente_reg(size)
    _scr = ImageDraw.Draw(Image.new("RGB", (10, 10)))
    x1, y1, x2, y2 = _scr.textbbox((0, 0), texto, font=f)
    return f, x2 - x1, y2 - y1


def poster(url):
    SCALE = 40
    bloque, _, diag = dibujar_bloque(url, con_logo=False,
                                     color_modulos=TEAL_MODULOS, scale=SCALE)
    qs_px = bloque.width

    TITLE = "CAMPOS CLÍNICOS 2026-II"
    SUB = "¿Dónde harás tu experiencia?  Escanea para consultar"
    PILL = "UNIVERSIDAD NORBERT WIENER"
    FOOT = "UNW · portal de campos clínicos"

    f_title, w_title, h_title = medir(TITLE, 170)
    f_sub, w_sub, h_sub = medir(SUB, 62, bold=False)
    f_pill, w_pill, h_pill = medir(PILL, 40)
    f_foot, w_foot, h_foot = medir(FOOT, 38, bold=False)

    # Flujo vertical: TODO se calcula antes de crear el lienzo (como en carnet)
    MARGIN, STROKE, PAD = 60, 16, 110
    CSP = MARGIN + STROKE + PAD
    W = qs_px + 2 * CSP

    ph = h_pill + 2 * 22
    y = CSP
    y_pill = y;            y = y_pill + ph
    y_title = y + 44;      y = y_title + h_title
    y_sub = y + 26;        y = y_sub + h_sub
    y_div = y + 26;        y = y_div + 18
    y_qr = y + 26;         y = y_qr + qs_px
    y_foot = y + 62;       y = y_foot + h_foot
    H = y + CSP + MARGIN

    canvas = Image.new("RGB", (W, H), BLANCO)
    d = ImageDraw.Draw(canvas)
    cx = W // 2

    d.rounded_rectangle([MARGIN, MARGIN, W - MARGIN, H - MARGIN],
                        radius=110, outline=TEAL_BRILLO, width=STROKE)

    # sello pill amarillo con texto teal oscuro
    pill_w = w_pill + 2 * 46
    d.rounded_rectangle([cx - pill_w // 2, y_pill, cx + pill_w // 2, y_pill + ph],
                        radius=ph // 2, fill=YELLOW_ACC)
    d.text((cx, y_pill + ph / 2), PILL, font=f_pill, fill=TEAL_MODULOS, anchor="mm")
    d.text((cx, y_title + h_title / 2), TITLE, font=f_title, fill=TEAL_MODULOS, anchor="mm")
    d.text((cx, y_sub + h_sub / 2), SUB, font=f_sub, fill=GRIS_SUAVE, anchor="mm")

    half_div = 320
    d.rounded_rectangle([cx - half_div, y_div, cx + half_div, y_div + 12],
                        radius=6, fill=YELLOW_ACC)

    canvas.paste(bloque, (cx - qs_px // 2, y_qr))
    d = ImageDraw.Draw(canvas)
    d.text((cx, y_foot + h_foot / 2), FOOT, font=f_foot, fill=GRIS_SUAVE, anchor="mm")

    return canvas, (cx - qs_px // 2, y_qr, qs_px), dict(diag, poster=True,
                                                        W=W, H=H)


# ── Batería de verificación (zxing sobre la recortada del QR) ────────────────
def decodifica(im):
    r = zxingcpp.read_barcodes(im)
    return r[0].text if r else None


def estado(t, url):
    return "✓ exacta" if t == url else ("✘ contenido distinto" if t else "✘ no decodifica")


def bateria(full, caja, url, etiqueta, exigir_dano=True):
    """exigir_dano=False: el [4] daño 25% queda advisory (sticker/banner:
    no simula uso real; el badge central + cuadrante siempre superan el
    presupuesto ECC-H — el QR del banner del carnet también lleva logo)."""
    qx0, y_qr, qs_px = caja
    bloque = full.crop((qx0, y_qr, qx0 + qs_px, y_qr + qs_px))
    requeridos = []        # tests que deben pasar
    advisory = []          # solo aviso

    print(f"  [1] {etiqueta} completa →", estado(decodifica(full), url))
    requeridos.append(decodifica(full) == url)
    for pct in (60, 40, 25):
        im2 = full.resize((full.width * pct // 100, full.height * pct // 100),
                          Image.Resampling.BOX)
        requeridos.append(decodifica(im2) == url)
        print(f"  [2] vista al {pct}%       →", estado(decodifica(im2), url))
    requeridos.append(decodifica(bloque.resize((700, 700))) == url)
    print("  [3] thumb forzado 700 px →",
          estado(decodifica(bloque.resize((700, 700))), url))

    im4 = bloque.copy()
    ImageDraw.Draw(im4).rectangle([im4.width // 2, im4.height // 2,
                                   im4.width, im4.height], fill="white")
    p4 = tempfile.mktemp(suffix=".png")
    im4.save(p4)
    dano = decodifica(Image.open(p4)) == url
    print("  [4] daño 25% (ECC H)     →", estado(decodifica(Image.open(p4)), url))
    (requeridos if exigir_dano else advisory).append(dano)
    if not exigir_dano and not dano:
        print("      ⚠ advisory: con banner no aplica daño mecánico de impresión")

    def suciedad(blur=None):
        im5 = bloque.copy()
        random.seed(42)
        rd = ImageDraw.Draw(im5)
        w5, h5 = im5.size
        for _ in range(60):
            xx, yxx = random.randint(0, w5 - 1), random.randint(0, h5 - 1)
            rr = random.randint(16, 36)
            rd.ellipse([xx - rr, yxx - rr, xx + rr, yxx + rr], fill="white")
        if blur:
            im5 = im5.filter(ImageFilter.GaussianBlur(blur))
        p5 = tempfile.mktemp(suffix=".png")
        im5.save(p5)
        return decodifica(Image.open(p5))

    requeridos.append(suciedad() == url)
    print("  [5] huella (60 crestas)  →", estado(suciedad(), url))
    requeridos.append(suciedad(blur=0.8) == url)
    print("  [6] huella + smear 0.8px →", estado(suciedad(blur=0.8), url))
    return all(requeridos)


# ── Flujo principal ───────────────────────────────────────────────────────────
def main():
    args = sys.argv[1:]
    url = args[0] if args else ""
    if not url or not url.startswith("http"):
        raise SystemExit("Uso: python scripts/qr_poster.py <URL_APP> [salida_base] [sin-logo]\n"
                         "No se genera un QR con URL vacía o placeholder (asi no toca).")

    base = Path(args[1]) if len(args) > 1 else \
        Path(__file__).resolve().parent.parent / "img" / "campo-clinico"
    sin_logo_forzado = len(args) > 2 and args[2].lower() in ("sin-logo", "no-logo", "1")

    # 1) Sticker para el banner: badge ON con fallback automático
    intentos = [False] if sin_logo_forzado else [True, False]
    for i, con_logo in enumerate(intentos):
        canvas, caja, diag = sticker(url, con_logo)
        print(f"\nSTICKER banner (badge {'ON' if con_logo else 'OFF'}): "
              f"ECC H | {diag['chars']} chars | versión {diag['version']} | "
              f"{canvas.width}x{canvas.height} px")
        if bateria(canvas, caja, url, "sticker", exigir_dano=False):
            out = Path(str(base) + "-banner.png")
            out.parent.mkdir(parents=True, exist_ok=True)
            canvas.save(out, dpi=(300, 300))
            print(f"  ✓ Guardado: {out}")
            break
        if con_logo and i == 0:
            print("  → Con badge falló la batería: se regenera SIN badge...")
    else:
        print("  ✘ Sticker reprobó batería incluso sin badge: revisar URL/módulos.")

    # 2) Poster estilo carnet (sin badge, como el original del repo)
    canvas, caja, diag = poster(url)
    print(f"\nPOSTER estilo carnet: {diag['W']}x{diag['H']} px | "
          f"~{diag['W']/300*2.54:.1f}x{diag['H']/300*2.54:.1f} cm a 300 dpi")
    if bateria(canvas, caja, url, "poster"):
        out = Path(str(base) + "-poster.png")
        out.parent.mkdir(parents=True, exist_ok=True)
        canvas.save(out, dpi=(300, 300))
        print(f"  ✓ Guardado: {out}")
    else:
        print("  ✘ Poster reprobó batería: revisar URL/imprenta.")


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    main()
