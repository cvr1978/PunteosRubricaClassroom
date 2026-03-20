// ============================================================
// PunteosRubricaClassroom — Google Apps Script
// Se conecta directamente a Google Classroom API y extrae
// todos los punteos con desglose de rúbricas automáticamente.
// ============================================================
//
// REQUISITO (una sola vez):
//   En el editor de Apps Script → "Services" (+) →
//   busca "Google Classroom API" → clic "Add"
//
// ============================================================

// Nombres de hojas internas
var H_DATOS      = '_DATOS';       // Hoja oculta con datos brutos
var H_REPORTE    = 'REPORTE';      // Reporte final

// Colores
var C_AZUL_OSC   = '#1565C0';
var C_VERDE_OSC  = '#1B5E20';
var C_VERDE_MED  = '#2E7D32';
var C_GRIS       = '#F5F5F5';
var C_TOTAL_BG   = '#E8F5E9';

// ============================================================
// MENÚ
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Rúbricas Classroom')
    .addItem('1.  Seleccionar curso y sincronizar', 'mostrarSelectorCurso')
    .addItem('2.  Generar Reporte Final',            'generarReporte')
    .addSeparator()
    .addItem('Sincronizar de nuevo (mismo curso)',   'sincronizarDenuevo')
    .addSeparator()
    .addItem('[Debug] Ver JSON de la API',            'debugVerAPI')
    .addItem('[Debug] Ver contenido de _DATOS',       'debugVerDATOS')
    .addItem('Instrucciones',                        'mostrarInstrucciones')
    .addToUi();
}

// ============================================================
// PASO 1 — SELECCIONAR CURSO
// ============================================================

function mostrarSelectorCurso() {
  var ui = SpreadsheetApp.getUi();

  var cursos;
  try {
    cursos = listarCursos();
  } catch (e) {
    ui.alert(
      'No se pudo conectar con Classroom.\n\n' +
      'Asegúrate de haber agregado "Google Classroom API" en:\n' +
      'Apps Script → Services (+) → Google Classroom API\n\n' +
      'Error técnico: ' + e.message
    );
    return;
  }

  if (!cursos || cursos.length === 0) {
    ui.alert('No se encontraron cursos activos en tu cuenta de Google Classroom.');
    return;
  }

  var opciones = cursos.map(function (c) {
    var etiqueta = c.name + (c.section ? '  —  ' + c.section : '');
    return '<option value="' + c.id + '">' + escHtml(etiqueta) + '</option>';
  }).join('');

  var html = HtmlService.createHtmlOutput(
    '<style>' +
    'body{font-family:Arial,sans-serif;padding:18px;font-size:13px;color:#333}' +
    'h3{color:#1565C0;margin:0 0 10px}' +
    'p{margin:4px 0 10px;color:#555;font-size:12px}' +
    'select{width:100%;padding:8px;font-size:13px;border:1px solid #ccc;border-radius:3px;margin-bottom:12px}' +
    '.btn{background:#1565C0;color:#fff;padding:9px 24px;border:none;border-radius:4px;cursor:pointer;font-size:13px}' +
    '.btn:hover{background:#0d47a1}.btn:disabled{background:#90A4AE;cursor:not-allowed}' +
    '#msg{margin-top:12px;color:#2E7D32;font-size:12px;display:none}' +
    '</style>' +
    '<h3>Seleccionar Curso</h3>' +
    '<p>Elige el curso cuyos punteos quieres importar.<br>' +
    'Se descargarán estudiantes, actividades, rúbricas y calificaciones automáticamente.</p>' +
    '<select id="sel">' + opciones + '</select>' +
    '<button class="btn" id="btn" onclick="go()">Importar datos</button>' +
    '<div id="msg">Conectando con Classroom, espera un momento…</div>' +
    '<script>' +
    'function go(){' +
    '  var s=document.getElementById("sel");' +
    '  document.getElementById("btn").disabled=true;' +
    '  document.getElementById("msg").style.display="block";' +
    '  google.script.run' +
    '    .withSuccessHandler(function(r){alert(r);google.script.host.close();})' +
    '    .withFailureHandler(function(e){alert("Error: "+e.message);' +
    '      document.getElementById("btn").disabled=false;' +
    '      document.getElementById("msg").style.display="none";})' +
    '    .importarCurso(s.value,s.options[s.selectedIndex].text);' +
    '}' +
    '</script>'
  ).setWidth(500).setHeight(300).setTitle('Seleccionar Curso');

  ui.showModalDialog(html, 'Seleccionar Curso');
}

// ============================================================
// IMPORTAR CURSO (corre en servidor)
// ============================================================

function importarCurso(courseId, courseLabel) {
  // ── 1. Estudiantes ──────────────────────────────────────
  var rawStudents = obtenerEstudiantes(courseId);
  var estudiantes = rawStudents.map(function (s) {
    var p = s.profile || {};
    return {
      id:     s.userId,
      nombre: (p.name && p.name.fullName) ? p.name.fullName : '',
      email:  p.emailAddress || ''
    };
  });

  // ── 2. Actividades ──────────────────────────────────────
  var rawWorks = obtenerActividades(courseId);

  // ── 3. Para cada actividad: rúbrica + submissions ───────
  var actividades = [];      // metadata
  var tablaPunteos = [];     // filas planas para la hoja oculta

  for (var i = 0; i < rawWorks.length; i++) {
    var cw = rawWorks[i];

    // Rúbrica (puede ser null si la actividad no tiene)
    var rubrica   = obtenerRubrica(courseId, cw.id);
    var criterios = parsearCriterios(rubrica);

    // Submissions de todos los estudiantes
    var subs = obtenerSubmissions(courseId, cw.id);

    // Mapa de fallback: si rubricGrades.points es null, buscar puntos por levelId
    var mapaNiveles = construirMapaNiveles(rubrica);

    // Guardar metadata de la actividad
    actividades.push({
      id:           cw.id,
      nombre:       cw.title || '',
      fecha:        formatFecha(cw.dueDate),
      maxPts:       cw.maxPoints || 0,
      tieneRubrica: criterios.length > 0,
      criterios:    criterios          // [{id, nombre, max}]
    });

    // Guardar punteos por estudiante
    for (var j = 0; j < subs.length; j++) {
      var sub   = subs[j];
      var total = sub.assignedGrade !== undefined && sub.assignedGrade !== null
                    ? sub.assignedGrade
                    : (sub.draftGrade !== undefined && sub.draftGrade !== null
                        ? sub.draftGrade : '');

      var filaPunteo = [cw.id, sub.userId, total];

      // Desglosar criterios — la API devuelve assignedRubricGrades (objeto por criterionId)
      if (criterios.length > 0) {
        var mapaGrades = {};
        // Preferir assignedRubricGrades; si no, draftRubricGrades
        var rubObj = sub.assignedRubricGrades || sub.draftRubricGrades || null;
        if (rubObj) {
          for (var crId in rubObj) {
            var rg  = rubObj[crId];
            var pts = rg.points;
            if ((pts === null || pts === undefined) && rg.levelId && mapaNiveles[rg.criterionId]) {
              pts = mapaNiveles[rg.criterionId][rg.levelId];
            }
            mapaGrades[rg.criterionId] = (pts !== null && pts !== undefined) ? pts : '';
          }
        }
        for (var k = 0; k < criterios.length; k++) {
          var val = mapaGrades.hasOwnProperty(criterios[k].id) ? mapaGrades[criterios[k].id] : '';
          filaPunteo.push(val);
        }
      }

      tablaPunteos.push(filaPunteo);
    }
  }

  // ── 4. Persistir en hoja oculta ─────────────────────────
  guardarEnHojaOculta(courseId, courseLabel, estudiantes, actividades, tablaPunteos);

  return (
    'Datos importados correctamente.\n\n' +
    'Curso:        ' + courseLabel + '\n' +
    'Estudiantes:  ' + estudiantes.length + '\n' +
    'Actividades:  ' + actividades.length + '\n' +
    'Con rúbrica:  ' + actividades.filter(function(a){ return a.tieneRubrica; }).length + '\n\n' +
    'Ahora ejecuta el Paso 2: Generar Reporte Final.'
  );
}

// ============================================================
// FUNCIONES DE CLASSROOM API
// ============================================================

function listarCursos() {
  var cursos = [], pageToken;
  do {
    var p   = { teacherId: 'me', pageSize: 50, courseStates: ['ACTIVE'] };
    if (pageToken) p.pageToken = pageToken;
    var res = Classroom.Courses.list(p);
    if (res.courses) cursos = cursos.concat(res.courses);
    pageToken = res.nextPageToken;
  } while (pageToken);
  return cursos;
}

function obtenerEstudiantes(courseId) {
  var lista = [], pageToken;
  do {
    var p   = { pageSize: 200 };
    if (pageToken) p.pageToken = pageToken;
    var res = Classroom.Courses.Students.list(courseId, p);
    if (res.students) lista = lista.concat(res.students);
    pageToken = res.nextPageToken;
  } while (pageToken);
  return lista;
}

function obtenerActividades(courseId) {
  var lista = [], pageToken;
  do {
    var p   = { pageSize: 100, orderBy: 'dueDate asc' };
    if (pageToken) p.pageToken = pageToken;
    var res = Classroom.Courses.CourseWork.list(courseId, p);
    if (res.courseWork) lista = lista.concat(res.courseWork);
    pageToken = res.nextPageToken;
  } while (pageToken);
  return lista;
}

function obtenerRubrica(courseId, courseWorkId) {
  try {
    // La API de Rúbricas fue añadida en 2023; puede no estar disponible en todos los entornos
    var res = Classroom.Courses.CourseWork.Rubrics.list(courseId, courseWorkId, { pageSize: 1 });
    var rubrics = (res && res.rubrics) ? res.rubrics : [];
    return rubrics.length > 0 ? rubrics[0] : null;
  } catch (e) {
    return null; // Si no está disponible o no hay rúbrica, retorna null sin error
  }
}

// Usamos UrlFetchApp para obtener rubricGrades, que el servicio avanzado
// de Apps Script no siempre devuelve. Sin fields filter para evitar 400.
function obtenerSubmissions(courseId, courseWorkId) {
  var lista     = [];
  var pageToken = null;
  var token     = ScriptApp.getOAuthToken();
  var baseUrl   = 'https://classroom.googleapis.com/v1/courses/' +
                  encodeURIComponent(courseId) + '/courseWork/' +
                  encodeURIComponent(courseWorkId) + '/studentSubmissions?pageSize=100';

  do {
    var url = baseUrl + (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');
    var res = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });
    var http = res.getResponseCode();
    var data = JSON.parse(res.getContentText());
    if (http !== 200) throw new Error((data.error && data.error.message) || 'HTTP ' + http);
    if (data.studentSubmissions) lista = lista.concat(data.studentSubmissions);
    pageToken = data.nextPageToken || null;
  } while (pageToken);

  return lista;
}

// Extrae criterios de una rúbrica → [{id, nombre, max}]
function parsearCriterios(rubrica) {
  if (!rubrica || !rubrica.criteria) return [];
  return rubrica.criteria.map(function (c) {
    var maxPts = 0;
    if (c.levels) {
      for (var i = 0; i < c.levels.length; i++) {
        var pts = parseFloat(c.levels[i].points) || 0;
        if (pts > maxPts) maxPts = pts;
      }
    }
    return { id: c.id, nombre: c.title || '', max: maxPts };
  });
}

// Construye mapa { criterionId → { levelId → points } }
// Se usa como fallback cuando rubricGrades.points es null pero levelId sí tiene valor.
function construirMapaNiveles(rubrica) {
  var mapa = {};
  if (!rubrica || !rubrica.criteria) return mapa;
  for (var i = 0; i < rubrica.criteria.length; i++) {
    var c = rubrica.criteria[i];
    mapa[c.id] = {};
    if (c.levels) {
      for (var j = 0; j < c.levels.length; j++) {
        mapa[c.id][c.levels[j].id] = parseFloat(c.levels[j].points) || 0;
      }
    }
  }
  return mapa;
}

// ============================================================
// HOJA OCULTA  (_DATOS)
// Estructura:
//   Fila 1 : CURSO | courseId | courseLabel
//   Fila 2 : ESTUDIANTES (marcador)
//   Filas  : userId | nombre | email
//   Fila   : ACTIVIDADES (marcador)
//   Filas  : actId | nombre | fecha | maxPts | tieneRubrica |
//             crit1Id | crit1Nombre | crit1Max |
//             crit2Id | crit2Nombre | crit2Max | ...
//   Fila   : PUNTEOS (marcador)
//   Filas  : actId | userId | total | crit1Pts | crit2Pts | ...
// ============================================================

function guardarEnHojaOculta(courseId, courseLabel, estudiantes, actividades, tablaPunteos) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var h  = ss.getSheetByName(H_DATOS);
  if (h) ss.deleteSheet(h);
  h = ss.insertSheet(H_DATOS);
  h.hideSheet();

  var filas = [];

  // Cabecera del curso
  filas.push(['CURSO', courseId, courseLabel]);
  filas.push([]);

  // Bloque de estudiantes
  filas.push(['ESTUDIANTES']);
  for (var i = 0; i < estudiantes.length; i++) {
    var e = estudiantes[i];
    filas.push([e.id, e.nombre, e.email]);
  }
  filas.push([]);

  // Bloque de actividades
  filas.push(['ACTIVIDADES']);
  for (var i = 0; i < actividades.length; i++) {
    var a   = actividades[i];
    var fila = [a.id, a.nombre, a.fecha, a.maxPts, a.tieneRubrica ? 'SI' : 'NO'];
    for (var k = 0; k < a.criterios.length; k++) {
      fila.push(a.criterios[k].id, a.criterios[k].nombre, a.criterios[k].max);
    }
    filas.push(fila);
  }
  filas.push([]);

  // Bloque de punteos
  filas.push(['PUNTEOS']);
  for (var i = 0; i < tablaPunteos.length; i++) {
    filas.push(tablaPunteos[i]);
  }

  // Normalizar ancho de columnas
  var maxCols = 0;
  for (var i = 0; i < filas.length; i++) {
    if (filas[i].length > maxCols) maxCols = filas[i].length;
  }
  for (var i = 0; i < filas.length; i++) {
    while (filas[i].length < maxCols) filas[i].push('');
  }

  if (filas.length > 0 && maxCols > 0) {
    h.getRange(1, 1, filas.length, maxCols).setValues(filas);
  }
}

// Lee la hoja oculta y devuelve { courseLabel, estudiantes, actividades, punteos }
function leerHojaOculta() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var h  = ss.getSheetByName(H_DATOS);
  if (!h) throw new Error('No hay datos importados. Ejecuta primero el Paso 1.');

  var datos = h.getDataRange().getValues();
  var resultado = {
    courseLabel: '',
    estudiantes:  [],   // [{id, nombre, email}]
    actividades:  [],   // [{id, nombre, fecha, maxPts, tieneRubrica, criterios:[]}]
    punteos:      {}    // { actId: { userId: { total, criterios:[] } } }
  };

  var modo = null;
  for (var r = 0; r < datos.length; r++) {
    var f = datos[r];
    var tag = String(f[0]).trim().toUpperCase();

    if (tag === 'CURSO') {
      resultado.courseLabel = String(f[2]).trim();
      continue;
    }
    if (tag === 'ESTUDIANTES') { modo = 'EST';  continue; }
    if (tag === 'ACTIVIDADES') { modo = 'ACT';  continue; }
    if (tag === 'PUNTEOS')     { modo = 'PUN';  continue; }
    if (!f[0] && !f[1])        { continue; }   // fila vacía

    if (modo === 'EST') {
      resultado.estudiantes.push({
        id:     String(f[0]).trim(),
        nombre: String(f[1]).trim(),
        email:  String(f[2]).trim()
      });

    } else if (modo === 'ACT') {
      var criterios = [];
      var col = 5;
      while (col < f.length && f[col] !== '') {
        criterios.push({
          id:     String(f[col]).trim(),
          nombre: String(f[col + 1]).trim(),
          max:    parseFloat(f[col + 2]) || 0
        });
        col += 3;
      }
      resultado.actividades.push({
        id:          String(f[0]).trim(),
        nombre:      String(f[1]).trim(),
        fecha:       String(f[2]).trim(),
        maxPts:      parseFloat(f[3]) || 0,
        tieneRubrica: String(f[4]).trim().toUpperCase() === 'SI',
        criterios:   criterios
      });

    } else if (modo === 'PUN') {
      var actId = String(f[0]).trim();
      var estId = String(f[1]).trim();
      var total = f[2] !== '' ? f[2] : null;
      var crits = [];
      for (var c = 3; c < f.length; c++) {
        crits.push(f[c] !== '' ? f[c] : null);
      }
      if (!resultado.punteos[actId]) resultado.punteos[actId] = {};
      resultado.punteos[actId][estId] = { total: total, criterios: crits };
    }
  }

  return resultado;
}

// ============================================================
// PASO 2 — GENERAR REPORTE FINAL
// ============================================================

function generarReporte() {
  var ui = SpreadsheetApp.getUi();
  var datos;
  try { datos = leerHojaOculta(); }
  catch (e) { ui.alert('Error: ' + e.message); return; }

  if (datos.estudiantes.length === 0 || datos.actividades.length === 0) {
    ui.alert('No hay datos suficientes. Ejecuta el Paso 1 primero.');
    return;
  }

  // ── Construir lista de columnas del reporte ──────────────
  // tipo: 'directa' | 'criterio' | 'total'
  var cols = [];
  for (var ai = 0; ai < datos.actividades.length; ai++) {
    var act = datos.actividades[ai];
    if (act.tieneRubrica && act.criterios.length > 0) {
      for (var k = 0; k < act.criterios.length; k++) {
        cols.push({
          cabecera:  act.nombre + '\n' + act.criterios[k].nombre + '\n(/' + act.criterios[k].max + ')',
          actId:     act.id,
          tipo:      'criterio',
          critIdx:   k
        });
      }
      cols.push({
        cabecera: act.nombre + '\nTOTAL\n(/' + act.maxPts + ')',
        actId:    act.id,
        tipo:     'total'
      });
    } else {
      cols.push({
        cabecera: act.nombre + '\n(/' + act.maxPts + ')',
        actId:    act.id,
        tipo:     'directa'
      });
    }
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaRep = ss.getSheetByName(H_REPORTE);
  if (hojaRep) ss.deleteSheet(hojaRep);
  hojaRep = ss.insertSheet(H_REPORTE);
  ss.setActiveSheet(hojaRep);
  ss.moveActiveSheet(1);

  var FIJOS    = 3;   // Nombre | Email | (espacio promedio futuro)
  var FILA_H   = 4;
  var totalCols = FIJOS + cols.length;

  // ── Filas 1-2: título (sin merge para evitar conflicto con freeze) ──────────
  // Pintamos toda la fila pero no fusionamos para que setFrozenColumns funcione.
  hojaRep.getRange(1, 1, 1, totalCols)
    .setBackground(C_AZUL_OSC).setFontColor('#FFFFFF');
  hojaRep.getRange(1, 1)
    .setValue('Reporte de Calificaciones — ' + datos.courseLabel)
    .setFontSize(13).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  hojaRep.setRowHeight(1, 30);

  hojaRep.getRange(2, 1, 1, totalCols).setBackground('#FAFAFA');
  hojaRep.getRange(2, 1)
    .setValue('Sincronizado desde Google Classroom: ' + new Date().toLocaleString('es-GT'))
    .setFontSize(10).setFontStyle('italic').setFontColor('#888888');
  hojaRep.setRowHeight(2, 18);

  // ── Fila 4: encabezados ───────────────────────────────────
  var encFijos = ['Estudiante', 'Email', 'Promedio'];
  var encActs  = cols.map(function (c) { return c.cabecera; });
  hojaRep.getRange(FILA_H, 1, 1, totalCols).setValues([encFijos.concat(encActs)]);
  hojaRep.getRange(FILA_H, 1, 1, FIJOS)
    .setBackground('#37474F').setFontColor('#FFFFFF').setFontWeight('bold');

  for (var ci = 0; ci < cols.length; ci++) {
    var bgH = cols[ci].tipo === 'directa'  ? C_AZUL_OSC :
              cols[ci].tipo === 'criterio' ? C_VERDE_MED : C_VERDE_OSC;
    hojaRep.getRange(FILA_H, FIJOS + 1 + ci)
      .setBackground(bgH).setFontColor('#FFFFFF').setFontWeight('bold');
  }

  hojaRep.getRange(FILA_H, 1, 1, totalCols)
    .setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('middle');
  hojaRep.setRowHeight(FILA_H, 58);

  // ── Filas de estudiantes ─────────────────────────────────
  var ests = datos.estudiantes;
  for (var ei = 0; ei < ests.length; ei++) {
    var filaN = FILA_H + 1 + ei;
    var bg    = ei % 2 === 0 ? '#FFFFFF' : C_GRIS;
    var est   = ests[ei];

    hojaRep.getRange(filaN, 1).setValue(est.nombre)
      .setBackground(bg).setFontColor('#000000').setHorizontalAlignment('left');
    hojaRep.getRange(filaN, 2).setValue(est.email)
      .setBackground(bg).setFontColor('#000000').setHorizontalAlignment('left');
    hojaRep.getRange(filaN, 3).setValue('')
      .setBackground(bg);

    for (var ci = 0; ci < cols.length; ci++) {
      var colN    = FIJOS + 1 + ci;
      var col     = cols[ci];
      var actPts  = datos.punteos[col.actId] || {};
      var estPts  = actPts[est.id] || null;
      var valor   = null;

      if (col.tipo === 'directa') {
        valor = estPts ? estPts.total : null;
        hojaRep.getRange(filaN, colN)
          .setBackground(ei % 2 === 0 ? '#E3F2FD' : '#BBDEFB');

      } else if (col.tipo === 'criterio') {
        valor = (estPts && estPts.criterios[col.critIdx] !== undefined)
                ? estPts.criterios[col.critIdx] : null;
        hojaRep.getRange(filaN, colN)
          .setBackground(ei % 2 === 0 ? '#E8F5E9' : '#C8E6C9');

      } else { // total
        valor = estPts ? estPts.total : null;
        hojaRep.getRange(filaN, colN)
          .setBackground(ei % 2 === 0 ? C_TOTAL_BG : '#DCEDC8')
          .setFontWeight('bold');
      }

      if (valor !== null && valor !== '') {
        hojaRep.getRange(filaN, colN)
          .setValue(valor).setHorizontalAlignment('center');
      }
    }
  }

  // ── Fila de máximos ──────────────────────────────────────
  var filaTot = FILA_H + ests.length + 2;
  hojaRep.getRange(filaTot, 1).setValue('PUNTAJE MÁXIMO')
    .setFontWeight('bold').setBackground('#E3F2FD').setFontColor('#000000');
  for (var ci = 0; ci < cols.length; ci++) {
    var act = datos.actividades.filter(function(a){ return a.id === cols[ci].actId; })[0];
    var maxVal = 0;
    if (cols[ci].tipo === 'criterio' && act) {
      maxVal = act.criterios[cols[ci].critIdx] ? act.criterios[cols[ci].critIdx].max : 0;
    } else if (act) {
      maxVal = act.maxPts;
    }
    hojaRep.getRange(filaTot, FIJOS + 1 + ci)
      .setValue(maxVal).setFontWeight('bold').setBackground('#E3F2FD')
      .setHorizontalAlignment('center').setFontColor('#000000');
  }

  // ── Formato final ────────────────────────────────────────
  hojaRep.setColumnWidth(1, 200);
  hojaRep.setColumnWidth(2, 215);
  hojaRep.setColumnWidth(3, 75);
  for (var c = FIJOS + 1; c <= totalCols; c++) hojaRep.setColumnWidth(c, 105);

  hojaRep.getRange(FILA_H, 1, ests.length + 1, totalCols)
    .setBorder(true, true, true, true, true, true, '#BDBDBD', SpreadsheetApp.BorderStyle.SOLID);

  hojaRep.setFrozenRows(FILA_H);
  hojaRep.setFrozenColumns(FIJOS);

  ui.alert(
    'Reporte generado correctamente.\n\n' +
    'Columnas azules    = puntaje directo de Classroom\n' +
    'Columnas verdes    = criterios de rúbrica (desde Classroom)\n' +
    'Columnas verde osc = TOTAL de la rúbrica\n\n' +
    'Si una celda está vacía significa que aún no hay calificación registrada.'
  );
}

// ============================================================
// SINCRONIZAR DE NUEVO (mismo curso)
// ============================================================

function sincronizarDenuevo() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var h  = ss.getSheetByName(H_DATOS);
  if (!h) { ui.alert('No hay un curso importado. Usa el Paso 1 primero.'); return; }

  var datos = h.getRange(1, 1, 1, 3).getValues()[0];
  var courseId    = String(datos[1]).trim();
  var courseLabel = String(datos[2]).trim();
  if (!courseId) { ui.alert('No se encontró el ID del curso. Usa el Paso 1 para seleccionarlo.'); return; }

  try {
    var msg = importarCurso(courseId, courseLabel);
    ui.alert('Datos actualizados.\n\n' + msg);
  } catch (e) {
    ui.alert('Error al sincronizar: ' + e.message);
  }
}

// ============================================================
// DEBUG — Ver JSON crudo de la API
// ============================================================

function debugVerAPI() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var hojaD = ss.getSheetByName(H_DATOS);
  if (!hojaD) { ui.alert('Primero importa un curso (Paso 1).'); return; }

  var fila1    = hojaD.getRange(1, 1, 1, 3).getValues()[0];
  var courseId = String(fila1[1]).trim();
  if (!courseId) { ui.alert('No se encontró el ID del curso en _DATOS.'); return; }

  // Buscar la primera actividad que tenga rúbrica
  var works = obtenerActividades(courseId);
  var cwTarget = null;
  var rubricaTarget = null;
  for (var i = 0; i < works.length; i++) {
    var r = obtenerRubrica(courseId, works[i].id);
    if (r) { cwTarget = works[i]; rubricaTarget = r; break; }
  }
  if (!cwTarget) { ui.alert('No se encontró ninguna actividad con rúbrica en este curso.'); return; }

  var token = ScriptApp.getOAuthToken();

  // 1. Obtener submissions via LIST (sin fields filter)
  var urlList = 'https://classroom.googleapis.com/v1/courses/' +
                encodeURIComponent(courseId) + '/courseWork/' +
                encodeURIComponent(cwTarget.id) + '/studentSubmissions?pageSize=3';
  var resList = UrlFetchApp.fetch(urlList, {
    headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true
  });
  var dataList = JSON.parse(resList.getContentText());
  var primeros = (dataList.studentSubmissions || []).slice(0, 2);

  // 2. Obtener el primer submission via GET individual
  var getIndividual = '';
  if (primeros.length > 0) {
    var urlGet = 'https://classroom.googleapis.com/v1/courses/' +
                 encodeURIComponent(courseId) + '/courseWork/' +
                 encodeURIComponent(cwTarget.id) + '/studentSubmissions/' +
                 encodeURIComponent(primeros[0].id);
    var resGet = UrlFetchApp.fetch(urlGet, {
      headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true
    });
    getIndividual = resGet.getContentText();
  }

  // Escribir todo en hoja _DEBUG
  var hojaDbg = ss.getSheetByName('_DEBUG');
  if (hojaDbg) ss.deleteSheet(hojaDbg);
  hojaDbg = ss.insertSheet('_DEBUG');

  var filas = [
    ['=== DIAGNÓSTICO rubricGrades ==='],
    ['Actividad:', cwTarget.title],
    ['CourseWork ID:', cwTarget.id],
    [''],
    ['--- Rúbrica (criteria) ---'],
    [JSON.stringify(rubricaTarget.criteria || [], null, 2)],
    [''],
    ['--- LIST /studentSubmissions (primeras 2) ---'],
  ];
  for (var i = 0; i < primeros.length; i++) {
    filas.push(['Submission ' + (i+1) + ':']);
    filas.push([JSON.stringify(primeros[i], null, 2)]);
    filas.push(['']);
  }
  filas.push(['--- GET individual (submission[0]) ---']);
  filas.push([getIndividual]);

  for (var r = 0; r < filas.length; r++) {
    hojaDbg.getRange(r + 1, 1).setValue(filas[r][0] || '');
  }
  hojaDbg.setColumnWidth(1, 900);

  ss.setActiveSheet(hojaDbg);
  ui.alert(
    'Diagnóstico completado.\n\n' +
    'Revisa la hoja "_DEBUG" que acaba de aparecer.\n' +
    'Busca el campo "rubricGrades" en el texto.\n\n' +
    'Comparte lo que ves para poder corregir el script.'
  );
}

function debugVerDATOS() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var hojaD = ss.getSheetByName(H_DATOS);
  if (!hojaD) { ui.alert('No hay hoja _DATOS. Ejecuta el Paso 1 primero.'); return; }

  var datos = hojaD.getDataRange().getValues();
  var hojaDbg = ss.getSheetByName('_DEBUG2');
  if (hojaDbg) ss.deleteSheet(hojaDbg);
  hojaDbg = ss.insertSheet('_DEBUG2');

  // Mostrar primeras filas de PUNTEOS con header
  var modo = null;
  var salida = [['=== Contenido de _DATOS (sección PUNTEOS) ===']];
  var contPun = 0;
  for (var r = 0; r < datos.length; r++) {
    var tag = String(datos[r][0]).trim().toUpperCase();
    if (tag === 'PUNTEOS') { modo = 'PUN'; salida.push(['--- PUNTEOS ---']); continue; }
    if (tag === 'ACTIVIDADES') { modo = 'ACT'; salida.push(['--- ACTIVIDADES ---']); continue; }
    if (modo === 'ACT' && datos[r][0] !== '') {
      salida.push([JSON.stringify(datos[r].slice(0, 15))]);
    }
    if (modo === 'PUN' && datos[r][0] !== '' && contPun < 5) {
      salida.push([JSON.stringify(datos[r].slice(0, 10))]);
      contPun++;
    }
  }

  for (var i = 0; i < salida.length; i++) {
    hojaDbg.getRange(i + 1, 1).setValue(salida[i][0]);
  }
  hojaDbg.setColumnWidth(1, 1000);
  ss.setActiveSheet(hojaDbg);
  ui.alert('Revisa la hoja "_DEBUG2" para ver el contenido de _DATOS.');
}

// ============================================================
// UTILIDADES
// ============================================================

function formatFecha(dueDate) {
  if (!dueDate) return '';
  var meses = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];
  return (dueDate.day || '') + '-' + meses[((dueDate.month || 1) - 1)];
}

function escHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// ============================================================
// INSTRUCCIONES
// ============================================================

function mostrarInstrucciones() {
  var html = HtmlService.createHtmlOutput(
    '<style>' +
    'body{font-family:Arial,sans-serif;padding:18px;font-size:13px;line-height:1.6;color:#333}' +
    'h2{color:#1565C0;margin:0 0 12px;font-size:15px}' +
    '.box{background:#F5F5F5;padding:10px 14px;border-radius:4px;margin:8px 0;border-left:4px solid #1565C0}' +
    '.box.verde{border-color:#2E7D32}' +
    '.box.naranja{border-color:#E65100}' +
    'b{color:#1565C0}.box.verde b{color:#1B5E20}' +
    '</style>' +
    '<h2>Rúbricas Classroom — Instrucciones</h2>' +

    '<div class="box naranja"><b>Requisito (una sola vez)</b><br>' +
    'En el editor de Apps Script → <b>Services (+)</b> → busca<br>' +
    '<b>Google Classroom API</b> → clic en <b>Add</b>.</div>' +

    '<div class="box"><b>Paso 1: Seleccionar curso y sincronizar</b><br>' +
    'El script se conecta a Classroom y descarga automáticamente:<br>' +
    '• Lista completa de estudiantes<br>' +
    '• Todas las actividades con fechas y puntajes máximos<br>' +
    '• Rúbricas configuradas en Classroom (criterios y niveles)<br>' +
    '• Calificaciones por estudiante, con desglose por criterio si hay rúbrica</div>' +

    '<div class="box verde"><b>Paso 2: Generar Reporte Final</b><br>' +
    'Crea la hoja <b>REPORTE</b> con todo consolidado:<br>' +
    '• Columnas <b>azules</b> = puntaje directo<br>' +
    '• Columnas <b>verde claro</b> = criterio de rúbrica<br>' +
    '• Columnas <b>verde oscuro</b> = TOTAL de la rúbrica<br>' +
    '• Celda vacía = actividad sin calificar aún</div>' +

    '<div class="box"><b>Actualizar datos</b><br>' +
    'Usa "Sincronizar de nuevo" para refrescar las calificaciones<br>' +
    'del mismo curso sin tener que seleccionarlo otra vez.</div>'
  ).setWidth(480).setHeight(420).setTitle('Instrucciones');
  SpreadsheetApp.getUi().showModalDialog(html, 'Instrucciones');
}
