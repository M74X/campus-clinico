# aplicar_fix.py - aplica el fix de cruces de horario a scripts/cruces_horario.py
# Uso (desde la raiz del repo campus-clinico):
#   python aplicar_fix.py
# Opcional: python aplicar_fix.py ruta/otro/cruces_horario.py
# Hace backup en cruces_horario.py.bak y verifica que el resultado compile.
import sys
from pathlib import Path

RUTA = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("scripts/cruces_horario.py")

# (nombre, texto_viejo, texto_nuevo, veces_esperadas)  - todo ASCII puro
EDITS = [
    ("contador de duplicadas en leer_sesiones",
     "    alumnos, orden = {}, []\n",
     "    alumnos, orden = {}, []\n    duplicadas = 0\n", 1),

    ("firmas por alumno (dedup)",
     '"ciclo": valor("ciclo"), "sesiones": []}',
     '"ciclo": valor("ciclo"), "sesiones": [],\n'
     '                               "firmas": set()}', 1),

    ("firma + descarte de fila duplicada",
     '        alumnos[codigo]["sesiones"].append({',
     '        firma = (curso_canonico(valor("curso")),\n'
     '                 normalizar(valor("seccion")), normalizar(valor("sede")),\n'
     '                 normalizar(valor("docente")), dia, rango[0], rango[1])\n'
     '        if firma in alumnos[codigo]["firmas"]:\n'
     '            duplicadas += 1             # fila repetida: misma asignacion\n'
     '            continue\n'
     '        alumnos[codigo]["firmas"].add(firma)\n'
     '        alumnos[codigo]["sesiones"].append({', 1),

    ("retorno de duplicadas",
     "    return alumnos, orden\n",
     "    return alumnos, orden, duplicadas\n", 1),

    ("funcion curso_canonico",
     "def buscar_archivo():",
     "def curso_canonico(curso):\n"
     "    \"\"\"Quita el codigo de curso: 'MH3103 NEUROLOGIA' y 'NEUROLOGIA'\n"
     "    (mismo curso con o sin codigo en el Excel) se comparan igual.\"\"\"\n"
     "    s = normalizar(curso)\n"
     "    return re.sub(r\"^[A-Z]{2,4}\\s?\\d{3,4}\\s+\", \"\", s).strip()\n"
     "\n"
     "\n"
     "def buscar_archivo():", 1),

    ("guard mismo curso + misma sede",
     "                a, b = lista[i], lista[j]\n",
     "                a, b = lista[i], lista[j]\n"
     '                if (curso_canonico(a["curso"]) == curso_canonico(b["curso"])\n'
     '                        and normalizar(a["sede"]) == normalizar(b["sede"])):\n'
     "                    descartados += 1   # misma clase, no un cruce\n"
     "                    continue\n", 1),

    ("contador de descartados",
     "    salida = []\n",
     "    salida, descartados = [], 0\n", 1),

    ("retorno de cruces y descartados",
     "    return salida\n",
     "    return salida, descartados\n", 1),

    ("firma de generar_excel",
     "def generar_excel(salida, filas, alumnos, n_sesiones):",
     "def generar_excel(salida, filas, alumnos, n_sesiones,\n"
     "                  duplicadas=0, descartados=0):", 1),

    ("fila RESUMEN de duplicadas",
     '        ("Alumnos con al menos un cruce", con_cruce),\n',
     '        ("Filas duplicadas colapsadas (misma asignacion)", duplicadas),\n'
     '        ("Alumnos con al menos un cruce", con_cruce),\n', 1),

    ("fila RESUMEN de descartados",
     '        ("Pares entre sedes DISTINTAS (distritos diferentes)",\n',
     '        ("Pares descartados - mismo curso y misma sede", descartados),\n'
     '        ("Pares entre sedes DISTINTAS (distritos diferentes)",\n', 1),

    ("desempaquetado en main",
     "    alumnos, orden = leer_sesiones(hoja, col, fila_enc)\n",
     "    alumnos, orden, duplicadas = leer_sesiones(hoja, col, fila_enc)\n", 1),

    ("acumular descartados en main",
     "    filas = []\n    for codigo in orden:\n        a = alumnos[codigo]\n"
     "        for cruce in cruces_de(a):\n",
     "    filas = []\n    descartados = 0\n    for codigo in orden:\n"
     "        a = alumnos[codigo]\n        cruces, desc = cruces_de(a)\n"
     "        descartados += desc\n        for cruce in cruces:\n", 1),

    ("reporte de duplicadas en consola",
     '    print(f"\\nAlumnos evaluados: {len(alumnos)} | '
     'Sesiones: {n_sesiones}")\n',
     '    print(f"\\nAlumnos evaluados: {len(alumnos)} | '
     'Sesiones: {n_sesiones}"\n'
     '          + (f" ({duplicadas} duplicadas colapsadas)" '
     'if duplicadas else ""))\n', 1),

    ("reporte de descartados en consola",
     "    if filas:\n",
     "    if descartados:\n"
     '        print(f"Pares descartados (mismo curso y misma sede): '
     '{descartados}")\n'
     "    if filas:\n", 1),

    ("paso de contadores al reporte (2 llamadas)",
     "            generar_excel(salida, filas, alumnos, n_sesiones)",
     "            generar_excel(salida, filas, alumnos, n_sesiones,\n"
     "                          duplicadas, descartados)", 2),
]


def main():
    if not RUTA.exists():
        raise SystemExit(f"[ERROR] No encontre {RUTA} (corre desde la raiz del repo)")
    texto = RUTA.read_text(encoding="utf-8-sig")
    crlf = "\r\n" in texto
    norm = texto.replace("\r\n", "\n")

    if '"firmas": set()' in norm and "descartados" in norm:
        print("El fix ya estaba aplicado. Nada por hacer.")
        return

    # 1) Verificar que cada ancla exista exactamente las veces esperadas
    fallas = []
    for nombre, viejo, nuevo, n in EDITS:
        c = norm.count(viejo)
        if c != n:
            fallas.append(f"  '{nombre}': encontrado {c} vez/veces, esperaba {n}")
    if fallas:
        print("[ERROR] Tu cruces_horario.py no coincide con lo esperado; "
              "NO se modifico nada. Detalle:")
        print("\n".join(fallas))
        print("(repo desactualizado? prueba 'git pull' y vuelve a intentar)")
        raise SystemExit(1)

    # 2) Aplicar
    for nombre, viejo, nuevo, n in EDITS:
        norm = norm.replace(viejo, nuevo)
        print(f"  ok - {nombre}")

    # 3) Validar que el resultado sea Python valido
    try:
        compile(norm, str(RUTA), "exec")
    except SyntaxError as e:
        raise SystemExit(f"[ERROR] El resultado no compila ({e}); no se guardo nada.")

    # 4) Guardar (conservando CRLF si el archivo original lo usaba)
    backup = Path(str(RUTA) + ".bak")
    backup.write_bytes(RUTA.read_bytes())
    with open(RUTA, "w", encoding="utf-8", newline="") as f:
        f.write(norm.replace("\n", "\r\n") if crlf else norm)

    print(f"\nFix aplicado a {RUTA} (backup: {backup.name}). "
          "Ahora: git add + commit + push")


if __name__ == "__main__":
    main()