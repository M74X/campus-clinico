"""Análisis de CRUCES DE HORARIO — Campo Clínico 2026-II.

Detecta alumnos que el MISMO DIA tienen sesiones incompatibles entre sí
y genera un reporte Excel completo (6 hojas). Clasificación:

  CRÍTICO   : las horas se SOLAPAN (clase a la vez que otra).
  INVIAJBLE : no se solapan, pero el gap es MENOR que el tiempo estimado
              de viaje entre las 2 sedes (distritos distintos).
  AJUSTADO  : sedes distintas, el viaje es factible pero el día queda
              muy cargado (gap < 120 min entre clases).

Los tiempos de viaje son ESTIMADOS por zona de Lima (tráfico real
aprox.), configurables en ZONA_SEDE / TRAVEL_ZONAS. Si una sede no está
mapeada se usa TIEMPO_VIAJE_DEFAULT.

Uso:
    python scripts/cruces_horario.py [archivo.xlsx] [salida.xlsx]
Sin argumento busca el primer *MATRICULA*.xlsx en la carpeta del
proyecto (o cwd). Salida por defecto: cruces_horario.xlsx.

Dependencias:
    python -m pip install openpyxl
"""
import re
import sys
import unicodedata
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ── CONFIG ────────────────────────────────────────────────────────────────────
COLUMNAS = {
    "codigo":  ["CODIGO ALUMNO", "CODIGO DEL ALUMNO", "CODIGO DE ALUMNO",
                "CODIGO ESTUDIANTE", "CODIGO DE ESTUDIANTE", "CODIGO"],
    "nombre":  ["APELLIDOS Y NOMBRES", "APELLIDO Y NOMBRES", "ALUMNO"],
    "curso":   ["CURSO", "ASIGNATURA"],
    "ciclo":   ["NIVEL ACTUAL", "CICLO ACTUAL", "CICLO", "NIVEL"],
    "seccion": ["SECCION"],
    "grupo":   ["GRUPO"],
    "docente": ["DOCENTE CAMPO CLINICO", "DOCENTE"],
    "sede":    ["SEDE HOSPITALARIA", "SEDE", "HOSPITAL"],
    "red":     ["RED", "RED DE SALUD"],
    "dia":     ["DIA"],
    "horario": ["HORARIO HOSPITALARIO", "HORARIO"],
}

HOJAS_DATO = ["MATRICULA"]
FILAS_ENCABEZADO = 6
GAP_AJUSTADO = 120          # min: gap < esto con sedes distintas = AJUSTADO
TIEMPO_VIAJE_DEFAULT = 75   # min: sede sin zona mapeada

# Zona aproximada de cada sede (para estimar tiempos de viaje entre distritos).
# Alias de nombres alternativos caen aquí también (ej. 'Hospital Negreiros').
ZONA_SEDE = {
    "SAENZ":            "LIMA CENTRO",
    "ALMENARA":         "LIMA CENTRO",
    "RAMON CASTILLA":   "LIMA CENTRO",
    "MONGRUT":          "LIMA NORTE",
    "NEGREIROS":        "SMP-CALLAO",      # Hospital II Luis Negreiros Vega
    "SABOGAL":          "SMP-CALLAO",
    "SJL":              "SJL",
    "MOLINA":           "SJL",             # Marino Molina Scippa (Canto Grande)
    "HUAYCAN":          "LIMA ESTE-ATE",
    "VITARTE":          "LIMA ESTE-ATE",
    "VOTO":             "SANTA ANITA",     # Jorge Voto Bernales
    "MILITAR":          "SAN BORJA",
    "DIAZ":             "LIMA SUR",        # Aurelio Díaz Ufano y Peral
    "CHOSICA":          "CHOSICA",
}

# Minutos aproximados de viaje entre zonas (tráfico de Lima, orientativo).
TRAVEL_ZONAS = {
    ("LIMA CENTRO", "LIMA CENTRO"):     40,
    ("LIMA NORTE", "LIMA NORTE"):       40,
    ("SMP-CALLAO", "SMP-CALLAO"):       35,
    ("SJL", "SJL"):                     40,
    ("LIMA ESTE-ATE", "LIMA ESTE-ATE"): 50,
    ("SANTA ANITA", "SANTA ANITA"):     35,
    ("SAN BORJA", "SAN BORJA"):         35,
    ("LIMA SUR", "LIMA SUR"):           50,
    ("CHOSICA", "CHOSICA"):             30,
    ("LIMA CENTRO", "LIMA NORTE"):      50,
    ("LIMA CENTRO", "SMP-CALLAO"):      45,
    ("LIMA CENTRO", "SJL"):             60,
    ("LIMA CENTRO", "LIMA ESTE-ATE"):   75,
    ("LIMA CENTRO", "SANTA ANITA"):     45,
    ("LIMA CENTRO", "SAN BORJA"):       35,
    ("LIMA CENTRO", "LIMA SUR"):        70,
    ("LIMA CENTRO", "CHOSICA"):        150,
    ("LIMA NORTE", "SMP-CALLAO"):       40,
    ("LIMA NORTE", "SJL"):              60,
    ("LIMA NORTE", "LIMA ESTE-ATE"):    90,
    ("LIMA NORTE", "SANTA ANITA"):      60,
    ("LIMA NORTE", "SAN BORJA"):        55,
    ("LIMA NORTE", "LIMA SUR"):         90,
    ("LIMA NORTE", "CHOSICA"):         165,
    ("SMP-CALLAO", "SJL"):              70,
    ("SMP-CALLAO", "LIMA ESTE-ATE"):    95,
    ("SMP-CALLAO", "SANTA ANITA"):      55,
    ("SMP-CALLAO", "SAN BORJA"):        55,
    ("SMP-CALLAO", "LIMA SUR"):         75,
    ("SMP-CALLAO", "CHOSICA"):         180,
    ("SJL", "LIMA ESTE-ATE"):           60,
    ("SJL", "SANTA ANITA"):             40,
    ("SJL", "SAN BORJA"):               60,
    ("SJL", "LIMA SUR"):                90,
    ("SJL", "CHOSICA"):                100,
    ("LIMA ESTE-ATE", "SANTA ANITA"):   50,
    ("LIMA ESTE-ATE", "SAN BORJA"):     70,
    ("LIMA ESTE-ATE", "LIMA SUR"):     100,
    ("LIMA ESTE-ATE", "CHOSICA"):       75,
    ("SANTA ANITA", "SAN BORJA"):       55,
    ("SANTA ANITA", "LIMA SUR"):        90,
    ("SANTA ANITA", "CHOSICA"):        150,
    ("SAN BORJA", "LIMA SUR"):          60,
    ("SAN BORJA", "CHOSICA"):          150,
    ("LIMA SUR", "CHOSICA"):           180,
}

RE_HORARIO = re.compile(
    r"(\d{1,2})[:.](\d{2})(?:\s*[ap]\.?m\.?)?\s*[-–—]\s*"
    r"(\d{1,2})[:.](\d{2})(?:\s*[ap]\.?m\.?)?", re.IGNORECASE)

# Tolerancia de escritura de días (MIRCOLES, SABADO, etc. se corrigen igual).
DIAS_ORDEN = ["LUNES", "MARTES", "MIERCOLES", "JUEVES", "VIERNES",
              "SABADO", "DOMINGO"]

# ── UTILIDADES ────────────────────────────────────────────────────────────────

def normalizar(v):
    """MAYÚSCULAS, sin tildes NI eñes (SÁENZ→SAENZ, MUÑOZ→MUNOZ)."""
    s = re.sub(r"\s+", " ", str(v or "").strip()).upper()
    s = unicodedata.normalize("NFD", s)
    return "".join(c for c in s if unicodedata.category(c) != "Mn") \
        .encode("ascii", "ignore").decode()


def zona_de(sede):
    """Zona de una sede por palabras clave (Hospital Negreiros ≈ Negreiros Vega)."""
    s = normalizar(sede)
    if not s:
        return "DESCONOCIDA"
    for clave, zona in ZONA_SEDE.items():
        if clave in s:
            return zona
    return "DESCONOCIDA"


def tiempo_viaje(sede1, sede2):
    """Minutos estimados entre dos sedes (por zonas)."""
    z1, z2 = zona_de(sede1), zona_de(sede2)
    if z1 == "DESCONOCIDA" or z2 == "DESCONOCIDA":
        return TIEMPO_VIAJE_DEFAULT
    if z1 == z2:
        return TRAVEL_ZONAS[(z1, z1)]
    return TRAVEL_ZONAS.get((z1, z2),
                            TRAVEL_ZONAS.get((z2, z1),
                                             TIEMPO_VIAJE_DEFAULT))


def parsear_horario(h):
    """'07:00 - 10:00' → (420, 600) en minutos; None si no se entiende."""
    m = RE_HORARIO.search(str(h or ""))
    if not m:
        return None
    h1, m1, h2, m2 = (int(x) for x in m.groups())
    ini, fin = h1 * 60 + m1, h2 * 60 + m2
    if fin <= ini:                       # cruza medianoche u H:MM invertido
        fin += 24 * 60
    return ini, fin


def solapamiento(a, b):
    return max(0, min(a[1], b[1]) - max(a[0], b[0]))


def gap_minutos(a, b):
    return max(0, max(a[0], b[0]) - min(a[1], b[1]))


def fmt_min(m):
    return f"{m // 60:02d}:{m % 60:02d}"


def dia_estandar(d):
    d = normalizar(d)
    for real in DIAS_ORDEN:
        if real[:4] == d[:4]:            # MIRCOLES→MIERCOLES, SABA→SABADO
            return real
    return d


def buscar_archivo():
    if len(sys.argv) > 1:
        return Path(sys.argv[1])
    raices = [Path(__file__).resolve().parent.parent, Path.cwd()]
    for raiz in raices:
        hits = sorted(raiz.glob("*MATRICULA*.xls*")) + \
               sorted(raiz.glob("*MATRICULADOS*.xls*"))
        if hits:
            return hits[0]
    return None


def columnas_de(ws):
    """Devuelve ({campo: idx}, n_fila) según la cabecera, o (None, None)."""
    for i, fila in enumerate(ws.iter_rows(
            min_row=1, max_row=FILAS_ENCABEZADO, values_only=True), start=1):
        enc = [normalizar(c) for c in fila]
        mapa = {}
        for campo, nombres in COLUMNAS.items():
            for nombre in nombres:
                if nombre in enc:
                    mapa[campo] = enc.index(nombre)
                    break
        if "codigo" in mapa and "curso" in mapa:
            return mapa, i
    return None, None


# ── LECTURA Y DETECCIÓN ───────────────────────────────────────────────────────

def leer_sesiones(ws, col, fila_enc):
    """{codigo: {nombre, ciclo, sesiones[]}} + orden de aparición."""
    alumnos, orden = {}, []
    for fila in ws.iter_rows(min_row=fila_enc + 1, values_only=True):
        if all(v is None or str(v).strip() == "" for v in fila):
            continue
        codigo = normalizar(fila[col["codigo"]])
        if not codigo:
            continue
        valor = lambda campo: (str(fila[col[campo]] or "").strip()
                               if campo in col else "")
        dia = dia_estandar(valor("dia"))
        horario_raw = valor("horario")
        rango = parsear_horario(horario_raw)
        if not dia or rango is None:
            continue                    # fila sin horario parseable: no cruza
        if codigo not in alumnos:
            alumnos[codigo] = {"nombre": valor("nombre"),
                               "ciclo": valor("ciclo"), "sesiones": []}
            orden.append(codigo)
        if valor("nombre") and not alumnos[codigo]["nombre"]:
            alumnos[codigo]["nombre"] = valor("nombre")
        alumnos[codigo]["sesiones"].append({
            "dia": dia, "ini": rango[0], "fin": rango[1],
            "horario": horario_raw, "curso": valor("curso"),
            "seccion": valor("seccion"), "sede": valor("sede"),
            "red": valor("red"), "docente": valor("docente"),
        })
    return alumnos, orden


def cruces_de(alumno):
    """Pares de sesiones del mismo día: CRITICO / INVIABLE / AJUSTADO."""
    por_dia = defaultdict(list)
    for s in alumno["sesiones"]:
        por_dia[s["dia"]].append(s)

    salida = []
    for dia, lista in sorted(por_dia.items()):
        for i in range(len(lista)):
            for j in range(i + 1, len(lista)):
                a, b = lista[i], lista[j]
                solape = solapamiento((a["ini"], a["fin"]),
                                      (b["ini"], b["fin"]))
                gap = gap_minutos((a["ini"], a["fin"]),
                                  (b["ini"], b["fin"]))
                viaje = tiempo_viaje(a["sede"], b["sede"])
                sedes_distintas = (normalizar(a["sede"]) !=
                                   normalizar(b["sede"])
                                   and bool(a["sede"]) and bool(b["sede"]))
                if solape > 0:
                    tipo = "CRITICO - SOLAPAMIENTO"
                elif sedes_distintas and gap < viaje:
                    tipo = "INVIABLE - NO ALCANZA A VIAJAR"
                elif sedes_distintas and gap < GAP_AJUSTADO:
                    tipo = "AJUSTADO - DIA CARGADO EN 2 DISTRITOS"
                else:
                    continue
                salida.append({
                    "dia": dia, "tipo": tipo, "a": a, "b": b,
                    "solape": solape, "gap": gap, "viaje": viaje,
                    "sedes_distintas": sedes_distintas,
                })
    return salida


SEV_ORDEN = {"CRITICO": 0, "INVIABLE": 1, "AJUSTADO": 2}


def sev_de(tipo):
    return next(k for k in SEV_ORDEN if tipo.startswith(k))


# ── REPORTE EXCEL ─────────────────────────────────────────────────────────────
FILL_HDR = PatternFill("solid", fgColor="0E7C86")
FILL_SEV = {
    "CRITICO":  PatternFill("solid", fgColor="F4CCCC"),
    "INVIABLE": PatternFill("solid", fgColor="FFF3CD"),
    "AJUSTADO": PatternFill("solid", fgColor="E8F0FE"),
}
BORDE = Border(*(Side(style="thin", color="D9D9D9"),) * 4)


def estilo_encabezado(ws, fila, titulos, anchos):
    for j, (titulo, ancho) in enumerate(zip(titulos, anchos), start=1):
        c = ws.cell(row=fila, column=j, value=titulo)
        c.font = Font(bold=True, color="FFFFFF", size=9)
        c.fill = FILL_HDR
        c.alignment = Alignment(horizontal="center", vertical="center",
                                wrap_text=True)
        c.border = BORDE
        ws.column_dimensions[get_column_letter(j)].width = ancho
    ws.row_dimensions[fila].height = 26


def celda(ws, r, j, v, fill=None, wrap=False):
    c = ws.cell(row=r, column=j, value=v)
    c.font = Font(size=9)
    c.border = BORDE
    c.alignment = Alignment(vertical="top", wrap_text=wrap)
    if fill:
        c.fill = fill
    return c


def hoja_detalle(wb, filas):
    ws = wb.active
    ws.title = "CRUCES DETALLE"
    ws.sheet_view.showGridLines = False
    titulos = ["SEVERIDAD", "CODIGO", "APELLIDOS Y NOMBRES", "CICLO", "DIA",
               "CURSO 1", "SECCION 1", "SEDE 1", "ZONA 1", "RED 1",
               "DOCENTE 1", "HORARIO 1",
               "CURSO 2", "SECCION 2", "SEDE 2", "ZONA 2", "RED 2",
               "DOCENTE 2", "HORARIO 2",
               "SOLAPE (MIN)", "GAP (MIN)", "VIAJE ESTIMADO (MIN)",
               "SEDES DISTINTAS"]
    anchos = [26, 12, 32, 7, 10, 30, 9, 24, 15, 20, 24, 20,
              30, 9, 24, 15, 20, 24, 20,
              13, 12, 20, 14]
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(titulos))
    t = ws.cell(row=1, column=1,
                value="CRUCES DE HORARIO — CAMPO CLÍNICO 2026-II (detalle por par de sesiones)")
    t.font = Font(bold=True, size=14, color="0E7C86")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(titulos))
    ws.cell(row=2, column=1,
            value=("CRÍTICO: horas que se solapan · INVIABLE: gap menor que el viaje "
                   f"estimado entre sedes · AJUSTADO: sedes distintas con gap < "
                   f"{GAP_AJUSTADO} min · Generado: {datetime.now():%d/%m/%Y %H:%M}")
            ).font = Font(size=9, italic=True, color="666666")
    estilo_encabezado(ws, 4, titulos, anchos)

    for r, cruce in enumerate(filas, start=5):
        a, b = cruce["a"], cruce["b"]
        sev = sev_de(cruce["tipo"])
        fill = FILL_SEV[sev]
        valores = [
            cruce["tipo"], cruce["codigo"], cruce["nombre"], cruce["ciclo"],
            cruce["dia"],
            a["curso"], a["seccion"], a["sede"], zona_de(a["sede"]), a["red"],
            a["docente"], a["horario"],
            b["curso"], b["seccion"], b["sede"], zona_de(b["sede"]), b["red"],
            b["docente"], b["horario"],
            cruce["solape"] or "", cruce["gap"], cruce["viaje"],
            "SÍ" if cruce["sedes_distintas"] else "No",
        ]
        for j, v in enumerate(valores, start=1):
            celda(ws, r, j, v, fill,
                  wrap=j in (3, 6, 8, 10, 11, 13, 15, 17, 18, 1))
    ws.freeze_panes = "C5"
    ws.auto_filter.ref = f"A4:{get_column_letter(len(titulos))}" \
                         f"{max(4, len(filas) + 4)}"


def hoja_alumnos(wb, filas):
    """Una fila por alumno afectado, ordenado por severidad y # cruces."""
    por_alumno = defaultdict(list)
    for f in filas:
        por_alumno[f["codigo"]].append(f)

    ws = wb.create_sheet("ALUMNOS AFECTADOS")
    ws.sheet_view.showGridLines = False
    titulos = ["CODIGO", "APELLIDOS Y NOMBRES", "CICLO", "N CRUCES",
               "SEVERIDAD MAXIMA", "DIAS AFECTADOS",
               "SEDES INVOLUCRADAS", "CURSOS INVOLUCRADOS", "REDES INVOLUCRADAS"]
    anchos = [12, 34, 7, 10, 24, 26, 40, 46, 28]
    ws.cell(row=1, column=1, value="ALUMNOS CON AL MENOS UN CRUCE").font = \
        Font(bold=True, size=13, color="0E7C86")
    estilo_encabezado(ws, 3, titulos, anchos)

    resumen = []
    for codigo, crs in por_alumno.items():
        sev = min((sev_de(c["tipo"]) for c in crs), key=lambda s: SEV_ORDEN[s])
        resumen.append({
            "codigo": codigo, "nombre": crs[0]["nombre"],
            "ciclo": crs[0]["ciclo"], "n": len(crs), "sev": sev,
            "dias": ", ".join(sorted({c["dia"] for c in crs},
                                     key=DIAS_ORDEN.index)),
            "sedes": ", ".join(sorted({x["sede"] for c in crs
                                       for x in (c["a"], c["b"])
                                       if x["sede"]})),
            "cursos": ", ".join(sorted({x["curso"] for c in crs
                                        for x in (c["a"], c["b"])})),
            "redes": ", ".join(sorted({x["red"] for c in crs
                                       for x in (c["a"], c["b"])
                                       if x["red"]})),
        })
    resumen.sort(key=lambda x: (SEV_ORDEN[x["sev"]], -x["n"], x["nombre"]))

    for r, a in enumerate(resumen, start=4):
        fill = FILL_SEV[a["sev"]]
        for j, v in enumerate([a["codigo"], a["nombre"], a["ciclo"], a["n"],
                               a["sev"], a["dias"], a["sedes"], a["cursos"],
                               a["redes"]], start=1):
            celda(ws, r, j, v, fill, wrap=j in (2, 7, 8, 9))
    ws.freeze_panes = "C4"
    ws.auto_filter.ref = f"A3:{get_column_letter(len(titulos))}" \
                         f"{max(3, len(resumen) + 3)}"


def hoja_tabla(wb, nombre, titulo, titulos, anchos, datos, fill_fn=None):
    ws = wb.create_sheet(nombre)
    ws.sheet_view.showGridLines = False
    ws.cell(row=1, column=1, value=titulo).font = \
        Font(bold=True, size=13, color="0E7C86")
    estilo_encabezado(ws, 3, titulos, anchos)
    for r, fila in enumerate(datos, start=4):
        for j, v in enumerate(fila, start=1):
            celda(ws, r, j, v, fill_fn(fila) if fill_fn else None)
    return ws


def hoja_dias(wb, filas):
    datos = []
    for dia in DIAS_ORDEN:
        crs = [f for f in filas if f["dia"] == dia]
        if not crs:
            continue
        datos.append([
            dia, len(crs),
            sum(1 for c in crs if c["tipo"].startswith("CRITICO")),
            sum(1 for c in crs if c["tipo"].startswith("INVIABLE")),
            sum(1 for c in crs if c["tipo"].startswith("AJUSTADO")),
        ])
    hoja_tabla(wb, "POR DIA", "CRUCES POR DÍA DE LA SEMANA",
               ["DIA", "CRUCES", "CRITICOS", "INVIABLES", "AJUSTADOS"],
               [14, 10, 12, 12, 12], datos)


def hoja_sedes(wb, filas):
    pares = Counter()
    viajes = {}
    for f in filas:
        clave = tuple(sorted((f["a"]["sede"], f["b"]["sede"])))
        pares[clave] += 1
        viajes[clave] = f["viaje"]
    datos = [[s1, zona_de(s1), s2, zona_de(s2), n, viajes[k]]
             for k, n in pares.most_common()
             for s1, s2 in [k]]
    hoja_tabla(wb, "POR SEDES", "CRUCES POR PAR DE SEDES (distritos)",
               ["SEDE 1", "ZONA 1", "SEDE 2", "ZONA 2", "CRUCES",
                "VIAJE ESTIMADO (MIN)"],
               [30, 16, 30, 16, 10, 20], datos)


def hoja_cursos(wb, filas):
    por_curso = Counter()
    for f in filas:
        for x in (f["a"], f["b"]):
            por_curso[x["curso"]] += 1
    datos = [[c, n] for c, n in por_curso.most_common()]
    hoja_tabla(wb, "POR CURSO", "CURSOS MÁS INVOLUCRADOS EN CRUCES",
               ["CURSO", "APARICIONES EN CRUCES"], [42, 24], datos)


def hoja_matriz(wb):
    zonas = sorted({z for z in ZONA_SEDE.values()})
    datos = []
    for z1 in zonas:
        fila = [z1]
        for z2 in zonas:
            v = TRAVEL_ZONAS.get((z1, z2)) or TRAVEL_ZONAS.get((z2, z1))
            fila.append(v if v else ("—" if z1 == z2 else
                                     TRAVEL_ZONAS.get((z1, z1), 40)))
        datos.append(fila)
    titulos = ["ZONA \\ ZONA"] + zonas
    ws = hoja_tabla(wb, "MATRIZ VIAJES",
                    "TIEMPO ESTIMADO DE VIAJE ENTRE ZONAS (min, orientativo)",
                    titulos, [18] + [16] * len(zonas), datos)
    for r in range(4, 4 + len(zonas)):
        for j in range(2, 2 + len(zonas)):
            ws.cell(row=r, column=j).alignment = Alignment(horizontal="center")
    ws.cell(row=5 + len(zonas), column=1,
            value="Estimados en tráfico urbano; sede sin zona mapeada usa "
                  f"{TIEMPO_VIAJE_DEFAULT} min. Editar TRAVEL_ZONAS en el "
                  "script para ajustar.").font = Font(size=8, italic=True)


def generar_excel(salida, filas, alumnos, n_sesiones):
    wb = openpyxl.Workbook()
    hoja_detalle(wb, filas)
    hoja_alumnos(wb, filas)
    hoja_dias(wb, filas)
    hoja_sedes(wb, filas)
    hoja_cursos(wb, filas)
    hoja_matriz(wb)

    # RESUMEN como primera hoja
    ws = wb.create_sheet("RESUMEN", 0)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 52
    ws.column_dimensions["B"].width = 14
    ws["A1"] = "RESUMEN — ANÁLISIS DE CRUCES DE HORARIO"
    ws["A1"].font = Font(bold=True, size=15, color="0E7C86")
    con_cruce = len({f["codigo"] for f in filas})
    criticos = [f for f in filas if f["tipo"].startswith("CRITICO")]
    inviables = [f for f in filas if f["tipo"].startswith("INVIABLE")]
    ajustados = [f for f in filas if f["tipo"].startswith("AJUSTADO")]
    datos = [
        ("Base analizada", alumnos and f"{len(alumnos)} alumnos · "
         f"{n_sesiones} sesiones"),
        ("Alumnos con al menos un cruce", con_cruce),
        ("Pares de sesiones en conflicto", len(filas)),
        ("  CRÍTICO — solapamiento de horas", len(criticos)),
        ("  INVIABLE — no alcanza a viajar entre sedes", len(inviables)),
        ("  AJUSTADO — 2 distritos, día cargado", len(ajustados)),
        ("Pares entre sedes DISTINTAS (distritos diferentes)",
         sum(1 for f in filas if f["sedes_distintas"])),
        ("Pares dentro de la MISMA sede",
         sum(1 for f in filas if not f["sedes_distintas"])),
        ("Días más afectados",
         ", ".join(f"{d} ({n})" for d, n in
                   Counter(f["dia"] for f in filas).most_common(3))),
        ("Sede más involucrada",
         Counter(x["sede"] for f in filas
                 for x in (f["a"], f["b"])).most_common(1)[0][0]),
        ("Generado", datetime.now().strftime("%d/%m/%Y %H:%M")),
    ]
    for i, (k, v) in enumerate(datos, start=3):
        ws.cell(row=i, column=1, value=k).font = Font(size=10)
        c = ws.cell(row=i, column=2, value=v)
        c.font = Font(size=10, bold=True)
        c.alignment = Alignment(horizontal="right", wrap_text=True)
    wb.save(salida)


# ── MAIN ──────────────────────────────────────────────────────────────────────

def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

    archivo = buscar_archivo()
    if not archivo:
        raise SystemExit("No encontré ningún *MATRICULA*.xlsx "
                         "(pásalo como argumento).")
    salida = Path(sys.argv[2]) if len(sys.argv) > 2 else \
        archivo.with_name("cruces_horario.xlsx")

    print(f"Archivo: {archivo.name}")
    wb = openpyxl.load_workbook(archivo, read_only=True, data_only=True)
    hoja = next((h for h in wb.worksheets
                 if any(p in normalizar(h.title) for p in HOJAS_DATO)),
                wb.worksheets[0])
    col, fila_enc = columnas_de(hoja)
    if col is None:
        raise SystemExit(f'Hoja "{hoja.title}": no hallé cabecera con '
                         'CODIGO ALUMNO + CURSO.')
    print(f'Hoja "{hoja.title}": cabecera en fila {fila_enc}')

    alumnos, orden = leer_sesiones(hoja, col, fila_enc)
    n_sesiones = sum(len(a["sesiones"]) for a in alumnos.values())

    filas = []
    for codigo in orden:
        a = alumnos[codigo]
        for cruce in cruces_de(a):
            filas.append({"codigo": codigo, "nombre": a["nombre"],
                          "ciclo": a["ciclo"], **cruce})
    filas.sort(key=lambda f: (SEV_ORDEN[sev_de(f["tipo"])],
                              DIAS_ORDEN.index(f["dia"]),
                              f["nombre"]))

    criticos = sum(1 for f in filas if f["tipo"].startswith("CRITICO"))
    inviables = sum(1 for f in filas if f["tipo"].startswith("INVIABLE"))
    ajustados = sum(1 for f in filas if f["tipo"].startswith("AJUSTADO"))
    con_cruce = len({f["codigo"] for f in filas})

    print(f"\nAlumnos evaluados: {len(alumnos)} | Sesiones: {n_sesiones}")
    print(f"Alumnos con cruce: {con_cruce}")
    print(f"Pares en conflicto: {len(filas)} "
          f"(CRITICOS {criticos} · INVIABLES {inviables} · AJUSTADOS {ajustados})")

    if filas:
        print("\nDías más afectados:",
              ", ".join(f"{d} ({n})" for d, n in
                        Counter(f["dia"] for f in filas).most_common()))
        sedes_top = Counter(x["sede"] for f in filas
                            for x in (f["a"], f["b"]))
        print("Sedes más involucradas:",
              ", ".join(f"{s} ({n})" for s, n in sedes_top.most_common(5)))
        cursos_top = Counter(x["curso"] for f in filas
                             for x in (f["a"], f["b"]))
        print("Cursos más involucrados:",
              ", ".join(f"{c} ({n})" for c, n in cursos_top.most_common(5)))

        try:
            generar_excel(salida, filas, alumnos, n_sesiones)
        except PermissionError:
            # Excel tiene el archivo abierto: guardar con sello de hora.
            salida = salida.with_name(
                salida.stem + "_" + datetime.now().strftime("%H%M%S") +
                salida.suffix)
            generar_excel(salida, filas, alumnos, n_sesiones)
            print("⚠ El destino estaba bloqueado (¿abierto en Excel?): "
                  "se guardó con sello de hora.")
        print(f"\n✓ Reporte: {salida} "
              "(hojas: RESUMEN · CRUCES DETALLE · ALUMNOS AFECTADOS · "
              "POR DIA · POR SEDES · POR CURSO · MATRIZ VIAJES)")
    else:
        print("✓ Sin cruces: no se genera reporte.")


if __name__ == "__main__":
    main()
