"""Valida la base MATRICULADOS_CON_SEDE... de Campo Clínico 2026-II.

Principio de datos institucionales: NO imprime nombres, códigos de
alumno, correos, teléfonos ni ninguna celda identificable. Solo:
  - estructura (hojas, cabeceras detectadas vs esperadas),
  - conteos agregados (filas, duplicados exactos, filas con vacíos),
  - issues a nivel CURSO/SECCION/GRUPO/DOCENTE (no personales),
  - cruzado contra la hoja RESUMEN si existe (curso vs fila).

Uso:
    python scripts/validar_base.py [archivo.xlsx]
Sin argumento busca el primer MATRICULADOS*.xlsx en la carpeta del
proyecto (campo-clinico o cwd).

Dependencias:
    python -m pip install openpyxl
"""
import re
import sys
from pathlib import Path

import openpyxl

# ── Cabeceras esperadas (orden indiferente, comparación normalizada) ─────────
ESPERADAS = [
    "FACULTAD", "CARRERA", "CODIGO ALUMNO", "APELLIDOS Y NOMBRES",
    "Correo Institucional", "Correo Personal", "Telefono", "CURSO",
    "CREDITOS", "NIVEL ACTUAL", "PERIODO", "SECCION", "GRUPO",
    "DOCENTE CAMPO CLINICO", "SEDE HOSPITALARIA", "RED", "DIA",
    "HORARIO HOSPITALARIO", "HORAS REGISTRADAS", "INICIO ROTACION",
    "FIN ROTACION",
]

CLAVE  = "CODIGO ALUMNO"
CURSO  = "CURSO"
DETALLE_MINIMO = {"DOCENTE CAMPO CLINICO", "SEDE HOSPITALARIA", "RED",
                  "DIA", "HORARIO HOSPITALARIO"}

RE_HORARIO = re.compile(r"^\d{1,2}[:.]\d{2}\s*[-–—]\s*\d{1,2}[:.]\d{2}$")


def normalizar(v):
    return re.sub(r"\s+", " ", str(v or "").strip()).upper() \
        .encode("ascii", "ignore").decode()


def buscar_archivo():
    if len(sys.argv) > 1:
        return Path(sys.argv[1])
    raices = [Path(__file__).resolve().parent.parent, Path.cwd()]
    for raiz in raices:
        hits = sorted(raiz.glob("*MATRICULADOS*.xls*"))
        if hits:
            return hits[0]
    return None


def columnas_de(ws, filas_encabezado=6):
    """Busca la fila de cabecera y devuelve ({COLUMNAS: idx}, n_fila)."""
    for i, fila in enumerate(ws.iter_rows(min_row=1, max_row=filas_encabezado,
                                          values_only=True), start=1):
        mapa = {}
        for j, celda in enumerate(fila):
            n = normalizar(celda)
            if n:
                mapa.setdefault(n, j)
        if CLAVE in mapa:
            return mapa, i
    return None, None


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

    archivo = buscar_archivo()
    if not archivo:
        raise SystemExit("No encontré ningún MATRICULADOS*.xlsx (pásalo como argumento).")
    print(f"Archivo: {archivo.name}")
    wb = openpyxl.load_workbook(archivo, read_only=True, data_only=True)
    print(f"Hojas: {wb.sheetnames}\n")

    hoja = wb.worksheets[0]
    col, fila_enc = columnas_de(hoja)
    if col is None:
        raise SystemExit('No hallé la cabecera con columna "CODIGO ALUMNO".')

    faltantes = [e for e in ESPERADAS if e.upper() not in col]
    extra = [c for c in col if c not in {e.upper() for e in ESPERADAS}]
    print(f"Hoja '{hoja.title}': cabecera en fila {fila_enc}, {len(col)} columnas")
    if faltantes:
        print(f"  ⚠ Cabeceras faltantes: {faltantes}")
    if extra:
        print(f"  ⚠ Cabeceras no en la lista: {extra}")
    if not faltantes:
        print("  ✓ Cabeceras completas")

    # ── Lectura de datos (solo importes y sello de columnas no personales) ──
    idx = {k: col[k.upper()] for k in (CLAVE, CURSO, "SECCION", "GRUPO",
                                       "DOCENTE CAMPO CLINICO",
                                       "SEDE HOSPITALARIA", "RED", "DIA",
                                       "HORARIO HOSPITALARIO", "PERIODO")
           if k.upper() in col}

    filas = dup_exactos = 0
    vistos = set()
    sin_detalle = {k: 0 for k in DETALLE_MINIMO}
    horario_malo = 0
    por_curso = {}
    por_codigo = {}
    firma_hoja = set()

    for fila in hoja.iter_rows(min_row=fila_enc + 1, values_only=True):
        if all(v is None or str(v).strip() == "" for v in fila):
            continue
        filas += 1
        firma = tuple(fila)
        if firma in firma_hoja:
            dup_exactos += 1
        firma_hoja.add(firma)

        codigo = normalizar(fila[idx[CLAVE]]) if CLAVE.upper() in col else ""
        if codigo:
            por_codigo[codigo] = por_codigo.get(codigo, 0) + 1

        cur = str(fila[idx[CURSO]]).strip() if CURSO.upper() in col else "(sin curso)"
        por_curso[cur] = por_curso.get(cur, 0) + 1
        for campo in DETALLE_MINIMO:
            v = fila[idx[campo]] if campo in idx else None
            if v is None or str(v).strip() == "":
                sin_detalle[campo] += 1
        h = fila[idx["HORARIO HOSPITALARIO"]] if "HORARIO HOSPITALARIO" in idx else None
        if h is not None and str(h).strip() and not RE_HORARIO.match(str(h).strip()):
            horario_malo += 1

    print(f"\nFilas de datos: {filas} | Duplicados exactos: {dup_exactos}")
    print(f"Alumnos (códigos únicos): {len(por_codigo)}")
    if por_codigo:
        dist = {}
        for n in por_codigo.values():
            dist[n] = dist.get(n, 0) + 1
        print(f"Dist. filas por alumno (n° filas → alumnos): {dict(sorted(dist.items()))}")

    if any(sin_detalle.values()):
        print(f"⚠ Filas con vacíos: {sin_detalle}")
    else:
        print("  ✓ Sin vacíos en DOCENTE/SEDE/RED/DIA/HORARIO")
    if horario_malo:
        print(f"⚠ HORARIO con formato inesperado: {horario_malo} fila(s) "
              "(esperado 'HH:MM - HH:MM').")

    # ── Cruzado con RESUMEN (conteos por curso, nivel agregado) ──
    resumen = next((n for n in wb.sheetnames if n.strip().upper().startswith("RESUMEN")), None)
    if resumen:
        print(f"\nCruce vs hoja '{resumen}':")
        res = {normalizar(c): (v or 0) for c, v in
               wb[resumen].iter_rows(min_row=2, values_only=True)
               if c and normalizar(c) != "CURSO" and "TOTAL" not in str(c).upper()}
        for cur, n in sorted(por_curso.items()):
            esperado = res.get(cur)
            if esperado is not None and esperado != n:
                print(f"  ⚠ '{cur}': base {n} vs resumen {esperado}")
        print(f"  Cursos en base: {len(por_curso)} | en resumen: {len(res)}")

    print("\nFin — solo estructura y conteos; cero datos personales impresos.")


if __name__ == "__main__":
    main()
