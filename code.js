/** ────────────────────────────────────────────────────────────────────────────
 *  CAMPO CLÍNICO 2026-II — Consulta de asignación
 *  Universidad Norbert Wiener
 *  Basado en app-carnet-unw/codigo.js (misma estructura de front + backend).
 *
 *  ARCHIVOS EN APPS SCRIPT:
 *    - Archivo de secuencia de comandos: pega este contenido (ej. "Code.gs").
 *    - Archivo HTML: pega index.html y NÓMBRALO exactamente "index"
 *      (doGet lo busca como 'index').
 *    - Vincula el proyecto a la Spreadsheet donde pegaste la hoja
 *      "MATRICULA CON HOSPITAL" (Archivo → Importar del Excel).
 *
 *  DATOS: solo se leen cabeceras y se busca el código enviado por el
 *  usuario; nunca se devuelven datos masivos ni correos/teléfonos.
 * ────────────────────────────────────────────────────────────────────────────*/

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Consulta de Campo Clínico - UNW')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─── CONFIGURACIÓN DE HOJA Y COLUMNAS ─────────────────────────────────────────
// Si cambian los nombres en el Excel, solo hay que tocar esta sección.
// Todo se compara normalizado: sin tildes, en mayúsculas y sin espacios sobrantes.

// Se reconocen las hojas por palabra clave (V2, V3, futuras versiones: igual).
var HOJAS_DATO = [/MATRICULA/];

// Nombres aceptados para cada columna. Se usa la primera columna que coincida.
var COLUMNAS = {
  codigo:  ['CODIGO ALUMNO', 'CODIGO DEL ALUMNO', 'CODIGO DE ALUMNO',
            'CODIGO ESTUDIANTE', 'CODIGO DE ESTUDIANTE', 'CODIGO'],
  nombre:  ['APELLIDOS Y NOMBRES', 'APELLIDO Y NOMBRES', 'ALUMNO'],
  curso:   ['CURSO', 'ASIGNATURA'],
  ciclo:   ['NIVEL ACTUAL', 'CICLO ACTUAL', 'CICLO', 'NIVEL'],
  seccion: ['SECCION'],
  grupo:   ['GRUPO'],
  docente: ['DOCENTE CAMPO CLINICO', 'DOCENTE'],
  sede:    ['SEDE HOSPITALARIA', 'SEDE', 'HOSPITAL'],
  red:     ['RED', 'RED DE SALUD'],
  dia:     ['DIA'],
  horario: ['HORARIO HOSPITALARIO', 'HORARIO'],
  inicio:  ['INICIO ROTACION', 'FECHA INICIO ROTACION', 'INICIO DE ROTACION'],
};

// Filas donde se busca el encabezado (por si el Excel trae filas vacías arriba)
var FILAS_ENCABEZADO = 5;

// ─── UTILIDADES ───────────────────────────────────────────────────────────────

function normalizar(str) {
  return (str || '').toString().trim().toUpperCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function limpiarCodigo(codigo) {
  return normalizar(codigo).replace(/\s+/g, '');
}

// "MH3102 CIRUGÍA GENERAL" → { nombre: 'CIRUGÍA GENERAL', seccion: false }
// Si no trae código de curso delante, se usa el texto completo.
function partirCurso(curso) {
  var t = (curso || '').toString().trim();
  var m = t.match(/^([A-Z]{2,4}\s?\d{3,4})\s+(.+)$/);
  return m ? { nombre: m[2].trim() } : { nombre: t };
}

// ─── FECHA DE INICIO DE ROTACIÓN ─────────────────────────────────────────────
// INICIO ROTACION es global (26/10/2026, lunes) y el DIA de cada grupo
// puede no coincidir con el día de esa fecha. La fecha que muestra el
// app es el PRIMER DIA del grupo on/after la fecha de inicio:
//   'Martes' + 26/10 → 'Martes 27 de octubre'
//   'Jueves' + 26/10 → 'Jueves 29 de octubre'

var MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio',
             'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];

// getDay(): 0=domingo … 6=sábado (normalizado: sin eñes ni tildes)
var DIAS_SEMANA = { DOMINGO: 0, LUNES: 1, MARTES: 2, MIERCOLES: 3,
                    JUEVES: 4, VIERNES: 5, SABADO: 6 };

// Date de Sheets o '26/10/2026' (DD/MM es el formato del Excel) → Date/null
function parsearFecha(inicio) {
  if (inicio instanceof Date && !isNaN(inicio)) return inicio;
  var m = (inicio || '').toString().trim()
    .match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  return m ? new Date(+m[3], +m[2] - 1, +m[1]) : null;
}

function formatearFecha(d) {
  var tz = Session.getScriptTimeZone();
  var mes = parseInt(Utilities.formatDate(d, tz, 'M'), 10);
  return Utilities.formatDate(d, tz, "d") + ' de ' + MESES[mes - 1];
}

// Primera fecha que cae en 'dia' a partir de 'inicio' (inclusive)
function primerSesion(dia, inicio) {
  var d = parsearFecha(inicio);
  var objetivo = DIAS_SEMANA[normalizar(dia)];
  if (!d || objetivo === undefined) return '';
  var delta = (objetivo - d.getDay() + 7) % 7;
  return formatearFecha(new Date(d.getFullYear(), d.getMonth(),
                                 d.getDate() + delta));
}

// 'Martes' + '26/10/2026' + '07:00 - 10:00'
//   → 'Martes 27 de octubre 07:00 hrs. - 10:00 hrs.'
// Si DIA no es un día válido o no hay fecha en el Excel, se muestra sin
// fecha ('Martes de 07:00 hrs. - 10:00 hrs.') para no contradecir el calendario.
function armarHorario(dia, horario, inicio) {
  var d = (dia || '').toString().trim();
  var h = (horario || '').toString().trim();
  var fecha = primerSesion(d, inicio);
  var rango = '';
  if (/^\d{1,2}[:.]\d{2}(.*)[-–—](.*)\d{1,2}[:.]\d{2}/.test(h)) {
    rango = h.split(/\s*[-–—]\s*/).map(function (p) {
      p = p.trim().replace('.', ':');
      return p + ' hrs.';
    }).join(' - ');
  } else if (h) {
    rango = h;   // formato no estándar: se muestra tal cual
  }

  if (!d && !fecha && !rango) return '';
  if (!rango) return fecha ? (d ? d + ' ' + fecha : fecha) : d;
  if (d && fecha) return d + ' ' + fecha + ' ' + rango;
  if (d) return d + ' de ' + rango;
  return fecha ? fecha + ' ' + rango : rango;
}

// ─── MAPEO DE COLUMNAS ────────────────────────────────────────────────────────

function mapearColumnas(encabezado) {
  var enc = encabezado.map(normalizar);
  var col = {};
  for (var campo in COLUMNAS) {
    col[campo] = -1;
    for (var i = 0; i < enc.length && col[campo] === -1; i++) {
      if (COLUMNAS[campo].indexOf(enc[i]) !== -1) col[campo] = i;
    }
  }
  return col;
}

function validarColumnas(hoja, col) {
  var faltan = [];
  if (col.codigo === -1) faltan.push('código (' + COLUMNAS.codigo.join(' / ') + ')');
  if (col.curso === -1)  faltan.push('curso (' + COLUMNAS.curso.join(' / ') + ')');
  if (col.nombre === -1) faltan.push('nombre (' + COLUMNAS.nombre.join(' / ') + ')');
  if (faltan.length) {
    throw new Error('Hoja "' + hoja.getName() + '": no se encontró la columna de ' + faltan.join(', '));
  }
}

function leerFila(fila, col) {
  var f = {};
  for (var campo in col) {
    f[campo] = col[campo] !== -1 ? fila[col[campo]].toString().trim() : '';
  }
  return f;
}

// ─── BÚSQUEDA ─────────────────────────────────────────────────────────────────

// Cualquier fila con el código del alumno es un curso asignado. Si la misma
// asignación aparece repetida (días/fechas duplicadas), se deduplica.
function buscarEnHoja(hoja, codigoBuscar) {
  var datos = hoja.getDataRange().getValues();

  // Ubicar la fila de encabezado: la primera que tenga columna de código
  var filaEnc = -1, col;
  for (var e = 0; e < Math.min(FILAS_ENCABEZADO, datos.length) && filaEnc === -1; e++) {
    col = mapearColumnas(datos[e]);
    if (col.codigo !== -1) filaEnc = e;
  }
  if (filaEnc === -1) col = mapearColumnas(datos[0] || []);
  validarColumnas(hoja, col);

  var nombre = '';
  var cursos = [];
  var vistos = {};

  for (var j = filaEnc + 1; j < datos.length; j++) {
    if (limpiarCodigo(datos[j][col.codigo]) !== codigoBuscar) continue;
    var f = leerFila(datos[j], col);

    if (!nombre) nombre = f.nombre;
    if (!f.curso) continue;

    var curso = partirCurso(f.curso);
    // Clave de deduplicación: la asignación completa de la fila
    var clave = [f.curso, f.seccion, f.grupo, f.docente, f.sede, f.red,
                 f.dia, f.horario, f.inicio].join('|');
    if (vistos[clave]) continue;
    vistos[clave] = true;

    cursos.push({
      curso:   curso.nombre,
      ciclo:   f.ciclo,
      seccion: f.seccion,
      docente: f.docente,
      sede:    f.sede,
      red:     f.red,
      horario: armarHorario(f.dia, f.horario, f.inicio),
    });
  }

  if (!nombre || !cursos.length) return null;
  return { encontrado: true, nombre: nombre, cursos: cursos };
}

// ─── ENTRY POINT ──────────────────────────────────────────────────────────────

function buscarAlumno(codigo) {
  var codigoBuscar = limpiarCodigo(codigo);
  if (!codigoBuscar) return { encontrado: false };

  var hojas = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var candidatas = hojas.filter(function (h) {
    var nombre = normalizar(h.getName());
    return HOJAS_DATO.some(function (patron) { return patron.test(nombre); });
  });
  if (!candidatas.length) candidatas = [hojas[0]];  // respaldo: primera hoja

  for (var i = 0; i < candidatas.length; i++) {
    var r = buscarEnHoja(candidatas[i], codigoBuscar);
    if (r) return r;
  }

  return { encontrado: false };
}
