# Portal de Campos Clínicos 2026-II

Consulta pública de asignación de campo clínico (curso, ciclo, sede
hospitalaria, horario, docente y red de salud) para estudiantes de la
Universidad Norbert Wiener. Web App de Google Apps Script sobre la
Spreadsheet de matrícula.

Basado en el front de [app-carnet-unw](https://github.com/) (mismo
estilo institucional: header teal, acento amarillo, tarjetas blancas).

## Estructura

```
Campo-Clinico/
├── code.js               → backend Apps Script (doGet + buscarAlumno)
├── index.html            → front Apps Script (consulta + tarjetas)
├── scripts/
│   ├── qr_poster.py      → genera sticker QR (banner) + poster impreso
│   ├── pegar_qr_banner.py→ pega el sticker en el banner de diseño
│   └── validar_base.py   → auditoría del Excel (agregados, sin PII)
└── img/
    ├── banner-campo-clinico-2026ii.png   ★ banner final con QR real
    ├── campo-clinico-banner.png          → sticker QR solo (marco azul + W)
    └── campo-clinico-poster.png          → poster impreso estilo carnet
```

## Despliegue en Apps Script

1. Importar la base (xlsx) a una Google Spreadsheet.
   - Hoja de datos: `MATRICULA CON HOSPITAL` (se reconoce por patrón
     `/MATRICULA/`; la hoja `RESUMEN` se ignora).
2. Crear proyecto en [script.google.com](https://script.google.com)
   y vincularlo a esa Spreadsheet.
3. Pegar `code.js` en un archivo de comandos y `index.html` en un
   archivo HTML llamado exactamente `index`.
4. Implementar → Web app (ejecutar como el propietario, acceso según
   lo que defina PAE).

La URL resultante queda embebida en los QR generados (ver Scripts).

### Cabeceras esperadas

Solo se leen las cabeceras relevantes; el mapeo es tolerante a
variaciones (ver `COLUMNAS` en `code.js`):

```
CODIGO ALUMNO · APELLIDOS Y NOMBRES · CURSO · NIVEL ACTUAL (ciclo)
SECCION · DOCENTE CAMPO CLINICO · SEDE HOSPITALARIA · RED · DIA
HORARIO HOSPITALARIO · INICIO ROTACION
```

### Regla de fecha

`INICIO ROTACION` es global (ej. 26/10/2026, lunes) y `DIA` es la
sesión semanal del grupo (Martes, Jueves…). La fecha mostrada es el
**primer DIA del grupo on/after la fecha de inicio**:

| DIA del grupo | Fecha mostrada |
|---|---|
| Lunes      | 26 de octubre |
| Martes     | 27 de octubre |
| Miércoles  | 28 de octubre |
| Jueves     | 29 de octubre |
| Viernes    | 30 de octubre |
| Sábado     | 31 de octubre |

## Scripts Python

```bash
python -m pip install segno zxing-cpp pillow openpyxl

# Auditoría de la base (agregados; no imprime datos personales)
python scripts/validar_base.py ruta/OPTIONAL.xlsx

# QRs: sticker banner + poster (batería zxing integrada, ECC H)
python scripts/qr_poster.py <URL_APP> [salida_base] [sin-logo]

# Componer el banner final con el QR de producción
python scripts/pegar_qr_banner.py [banner] [sticker] [salida]
```

## Seguridad de datos

- El repo **no contiene datos de alumnos**: solo código, diseño y
  agregados. La base vive en Google Sheets (vinculada al proyecto).
- `code.js` atiende búsquedas por código; nunca expone listas,
  correos ni teléfonos.
- `validar_base.py` imprime exclusivamente conteos y estructura.
- Si el xlsx se copia aquí para trabajar, el `.gitignore` ya lo
  bloquea (`*.xlsx`).

## Créditos

Estilo front: repositorio `app-carnet-unw` (PAE, Universidad
Norbert Wiener). Banner institucional: diseño UNW 2026-II.
