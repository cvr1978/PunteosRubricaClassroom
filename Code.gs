// ============================================================
// PunteosRubricaClassroom — Google Apps Script
// ============================================================
// Pega este código en Extensions > Apps Script de tu
// Google Sheets y recarga la hoja para ver el menú.
// ============================================================

var HOJA_DATOS    = '_DATOS_CLASSROOM';
var HOJA_PLANTILLA = 'PLANTILLA';
var HOJA_REPORTE  = 'REPORTE';
var PREFIJO_RUB   = 'RUB_';

// Colores
var C_AZUL_OSC   = '#1565C0';
var C_AZUL_MED   = '#1976D2';
var C_VERDE_OSC  = '#1B5E20';
var C_VERDE_MED  = '#2E7D32';
var C_AMARILLO   = '#FFF9C4';
var C_TOTAL_BG   = '#E8F5E9';
var C_GRIS       = '#F5F5F5';
var C_GRIS_MED   = '#E0E0E0';

// ============================================================
// MENÚ
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Rúbricas Classroom')
    .addItem('Paso 1: Importar CSV de Classroom', 'mostrarDialogoImportar')
    .addSeparator()
    .addItem('Paso 2: Generar Plantilla de Rúbricas', 'generarPlantilla')
    .addItem('Paso 3: Crear Hojas de Punteos', 'crearHojasPunteos')
    .addItem('Paso 4: Generar Reporte Final', 'generarReporte')
    .addSeparator()
    .addItem('Instrucciones', 'mostrarInstrucciones')
    .addToUi();
}

// ============================================================
// PASO 1: IMPORTAR CSV
// ============================================================

function mostrarDialogoImportar() {
  var html = HtmlService.createHtmlOutput(
    '<style>' +
    'body{font-family:Arial,sans-serif;padding:16px;font-size:13px}' +
    'h3{color:#1565C0;margin-top:0}' +
    'p{color:#555;margin:4px 0 10px}' +
    'textarea{width:100%;height:180px;font-size:11px;border:1px solid #ccc;border-radius:3px;padding:6px;box-sizing:border-box}' +
    'button{background:#1565C0;color:#fff;padding:9px 22px;border:none;border-radius:4px;cursor:pointer;font-size:13px;margin-top:10px}' +
    'button:hover{background:#0d47a1}' +
    '.note{font-size:11px;color:#888;margin-top:8px}' +
    '</style>' +
    '<h3>Importar CSV de Google Classroom</h3>' +
    '<p>Abre el archivo CSV en un editor de texto o en el Bloc de notas,<br>' +
    'selecciona todo (Ctrl+A), copia (Ctrl+C) y pega aquí:</p>' +
    '<textarea id="csv" placeholder="Pega aquí el contenido completo del CSV..."></textarea>' +
    '<p class="note">El CSV debe ser el exportado directamente desde Google Classroom > Calificaciones.</p>' +
    '<button onclick="enviar()">Importar</button>' +
    '<script>' +
    'function enviar(){' +
    '  var csv=document.getElementById("csv").value;' +
    '  if(!csv.trim()){alert("Por favor pega el CSV primero");return;}' +
    '  google.script.run' +
    '    .withSuccessHandler(function(){alert("CSV importado correctamente.");google.script.host.close();})' +
    '    .withFailureHandler(function(e){alert("Error: "+e.message);})' +
    '    .importarCSVDesdeTexto(csv);' +
    '}' +
    '</script>'
  ).setWidth(560).setHeight(360).setTitle('Paso 1: Importar CSV');
  SpreadsheetApp.getUi().showModalDialog(html, 'Paso 1: Importar CSV');
}

function importarCSVDesdeTexto(csvTexto) {
  var filas = Utilities.parseCsv(csvTexto);
  if (!filas || filas.length < 6) {
    throw new Error('El CSV no tiene el formato esperado de Google Classroom (mínimo 6 filas).');
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(HOJA_DATOS);
  if (hoja) ss.deleteSheet(hoja);
  hoja = ss.insertSheet(HOJA_DATOS);
  hoja.hideSheet();

  var maxCols = 0;
  for (var i = 0; i < filas.length; i++) {
    if (filas[i].length > maxCols) maxCols = filas[i].length;
  }
  for (var i = 0; i < filas.length; i++) {
    while (filas[i].length < maxCols) filas[i].push('');
    hoja.getRange(i + 1, 1, 1, maxCols).setValues([filas[i]]);
  }
}

// ============================================================
// LEER DATOS DE CLASSROOM
// ============================================================

function leerDatosClassroom() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(HOJA_DATOS);
  if (!hoja) throw new Error('Primero importa el CSV usando el Paso 1.');

  var datos = hoja.getDataRange().getValues();
  if (datos.length < 6) throw new Error('El CSV importado tiene muy pocas filas.');

  // Fila 0: curso + fechas | Fila 1: sección + nombres de actividades
  // Fila 2: "ABRIR CLASSROOM" + puntos máximos
  // Fila 3: vacía | Fila 4: promedio | Fila 5+: estudiantes
  var curso   = String(datos[0][0]).trim();
  var seccion = String(datos[1][0]).trim();

  // Actividades desde columna índice 4
  var actividades = [];
  for (var c = 4; c < datos[1].length; c++) {
    var nombre = String(datos[1][c]).trim();
    if (!nombre || nombre === '' || nombre.toLowerCase() === 'nan') continue;
    var fecha  = String(datos[0][c]).trim();
    var maxRaw = String(datos[2][c]).replace(',', '.');
    var maxPts = parseFloat(maxRaw) || 0;
    actividades.push({ nombre: nombre, fecha: fecha, maxPts: maxPts, colIdx: c });
  }

  // Estudiantes desde fila índice 5
  var estudiantes = [];
  for (var r = 5; r < datos.length; r++) {
    var apellido = String(datos[r][0]).trim();
    if (!apellido || apellido === '' || apellido.toLowerCase() === 'nan') continue;
    if (apellido.toLowerCase().indexOf('promedio') !== -1) continue;

    var est = {
      apellido:   apellido,
      nombre:     String(datos[r][1]).trim(),
      email:      String(datos[r][2]).trim(),
      porcentaje: String(datos[r][3]).trim(),
      punteos:    {}
    };
    for (var ai = 0; ai < actividades.length; ai++) {
      var raw = datos[r][actividades[ai].colIdx];
      var val = raw !== '' && raw !== null && raw !== undefined ? parseFloat(String(raw).replace(',', '.')) : null;
      est.punteos[actividades[ai].nombre] = isNaN(val) ? null : val;
    }
    estudiantes.push(est);
  }

  return { curso: curso, seccion: seccion, actividades: actividades, estudiantes: estudiantes };
}

// ============================================================
// PASO 2: GENERAR PLANTILLA
// ============================================================

function generarPlantilla() {
  var ui = SpreadsheetApp.getUi();
  var datosClassroom;
  try { datosClassroom = leerDatosClassroom(); }
  catch(e) { ui.alert('Error: ' + e.message); return; }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaExist = ss.getSheetByName(HOJA_PLANTILLA);
  if (hojaExist) {
    var resp = ui.alert(
      'La hoja PLANTILLA ya existe.',
      '¿Deseas recrearla? Perderás la configuración actual.',
      ui.ButtonSet.YES_NO
    );
    if (resp !== ui.Button.YES) return;
    ss.deleteSheet(hojaExist);
  }

  var h = ss.insertSheet(HOJA_PLANTILLA);
  ss.setActiveSheet(h);
  ss.moveActiveSheet(1);

  // Título
  h.getRange('A1').setValue('Configuración de Rúbricas — ' + datosClassroom.curso + '  |  ' + datosClassroom.seccion);
  h.getRange('A1').setFontSize(13).setFontWeight('bold').setBackground(C_AZUL_OSC).setFontColor('#FFFFFF');
  h.getRange('A1:L1').merge();
  h.setRowHeight(1, 30);

  // Subtítulo instrucción
  h.getRange('A2').setValue(
    'Instrucciones: En la columna "¿Tiene Rúbrica?" escribe SI o NO. ' +
    'Para las actividades con rúbrica, define los criterios y sus puntos máximos (celdas amarillas).'
  );
  h.getRange('A2:L2').merge();
  h.getRange('A2').setFontSize(10).setFontStyle('italic').setFontColor('#555555').setBackground('#E3F2FD').setWrap(true);
  h.setRowHeight(2, 32);

  // Encabezados tabla (fila 4)
  var FILA_H = 4;
  var headers = [
    'Actividad', 'Fecha', 'Máx. Pts', '¿Tiene Rúbrica?\n(SI / NO)',
    'Criterio 1', 'Máx C1', 'Criterio 2', 'Máx C2',
    'Criterio 3', 'Máx C3', 'Criterio 4', 'Máx C4'
  ];
  var rngH = h.getRange(FILA_H, 1, 1, headers.length);
  rngH.setValues([headers]);
  rngH.setBackground(C_AZUL_OSC).setFontColor('#FFFFFF').setFontWeight('bold')
      .setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('middle');
  h.setRowHeight(FILA_H, 46);

  // Validación desplegable SI/NO
  var validacion = SpreadsheetApp.newDataValidation()
    .requireValueInList(['SI', 'NO'], true).build();

  // Filas de actividades
  var acts = datosClassroom.actividades;
  for (var i = 0; i < acts.length; i++) {
    var fila = FILA_H + 1 + i;
    var bg   = i % 2 === 0 ? '#FFFFFF' : C_GRIS;

    h.getRange(fila, 1).setValue(acts[i].nombre).setBackground(bg).setFontColor('#000000').setHorizontalAlignment('left');
    h.getRange(fila, 2).setValue(acts[i].fecha).setBackground(bg).setFontColor('#000000').setHorizontalAlignment('center');
    h.getRange(fila, 3).setValue(acts[i].maxPts).setBackground(bg).setFontColor('#000000').setHorizontalAlignment('center');
    h.getRange(fila, 4).setValue('NO').setBackground(C_GRIS_MED).setFontColor('#000000')
                       .setHorizontalAlignment('center').setDataValidation(validacion);
    // Celdas de criterios (amarillo = editables)
    h.getRange(fila, 5, 1, 8).setBackground(C_AMARILLO).setFontColor('#000000');
  }

  // Anchos
  h.setColumnWidth(1, 240);
  h.setColumnWidth(2, 80);
  h.setColumnWidth(3, 75);
  h.setColumnWidth(4, 100);
  for (var c = 5; c <= 12; c++) h.setColumnWidth(c, 115);

  // Bordes
  var numActas = acts.length;
  h.getRange(FILA_H, 1, numActas + 1, headers.length)
   .setBorder(true, true, true, true, true, true, '#BDBDBD', SpreadsheetApp.BorderStyle.SOLID);

  h.setFrozenRows(FILA_H);

  ui.alert(
    'Plantilla generada.\n\n' +
    'Siguiente paso:\n' +
    '1. En la columna "¿Tiene Rúbrica?" cambia a SI las actividades evaluadas con rúbrica.\n' +
    '2. Para esas actividades define los criterios y sus puntos máximos en las celdas amarillas.\n' +
    '3. Ejecuta el Paso 3: Crear Hojas de Punteos.'
  );
}

// ============================================================
// LEER CONFIGURACIÓN DE LA PLANTILLA
// ============================================================

function leerConfigPlantilla(hojaPlantilla, actividades) {
  var datos = hojaPlantilla.getDataRange().getValues();

  // Encontrar fila de encabezados buscando "Actividad"
  var filaH = -1;
  for (var r = 0; r < datos.length; r++) {
    if (String(datos[r][0]).trim() === 'Actividad') { filaH = r; break; }
  }
  if (filaH === -1) throw new Error('No se encontró la tabla en la hoja PLANTILLA.');

  var config = [];
  for (var r = filaH + 1; r < datos.length; r++) {
    var nombreAct = String(datos[r][0]).trim();
    if (!nombreAct || nombreAct === '') continue;

    var tieneRubrica = String(datos[r][3]).trim().toUpperCase() === 'SI';
    var criterios = [];
    if (tieneRubrica) {
      for (var ci = 0; ci < 4; ci++) {
        var nomCrit = datos[r][4 + ci * 2] ? String(datos[r][4 + ci * 2]).trim() : '';
        var maxCrit = datos[r][5 + ci * 2];
        if (nomCrit && nomCrit !== '') {
          var maxVal = parseFloat(String(maxCrit).replace(',', '.')) || 0;
          criterios.push({ nombre: nomCrit, max: maxVal });
        }
      }
    }

    var actOrig = null;
    for (var ai = 0; ai < actividades.length; ai++) {
      if (actividades[ai].nombre === nombreAct) { actOrig = actividades[ai]; break; }
    }

    config.push({
      nombre:       nombreAct,
      fecha:        actOrig ? actOrig.fecha : '',
      maxPts:       actOrig ? actOrig.maxPts : 0,
      tieneRubrica: tieneRubrica,
      criterios:    criterios
    });
  }
  return config;
}

// ============================================================
// PASO 3: CREAR HOJAS DE PUNTEOS
// ============================================================

function crearHojasPunteos() {
  var ui = SpreadsheetApp.getUi();
  var datosClassroom;
  try { datosClassroom = leerDatosClassroom(); }
  catch(e) { ui.alert('Error: ' + e.message); return; }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaPlantilla = ss.getSheetByName(HOJA_PLANTILLA);
  if (!hojaPlantilla) {
    ui.alert('Primero ejecuta el Paso 2: Generar Plantilla.');
    return;
  }

  var config;
  try { config = leerConfigPlantilla(hojaPlantilla, datosClassroom.actividades); }
  catch(e) { ui.alert('Error al leer la plantilla: ' + e.message); return; }

  var actConRubrica = config.filter(function(c) { return c.tieneRubrica && c.criterios.length > 0; });
  if (actConRubrica.length === 0) {
    ui.alert('No hay actividades marcadas con SI en la plantilla, o no tienen criterios definidos.');
    return;
  }

  var creadas = 0;
  for (var i = 0; i < actConRubrica.length; i++) {
    var act = actConRubrica[i];
    var nombreHoja = (PREFIJO_RUB + act.nombre).substring(0, 100).replace(/[\/\\?\*\[\]:]/g, '_');
    var hojaRub = ss.getSheetByName(nombreHoja);
    if (hojaRub) ss.deleteSheet(hojaRub);
    hojaRub = ss.insertSheet(nombreHoja);
    crearHojaRubrica(hojaRub, act, datosClassroom.estudiantes);
    creadas++;
  }

  ui.alert(
    'Se crearon ' + creadas + ' hoja(s) de punteos.\n\n' +
    'Siguiente paso:\n' +
    '1. Ve a cada hoja que empieza con RUB_\n' +
    '2. Ingresa los puntajes en las celdas amarillas\n' +
    '3. La columna TOTAL se calcula automáticamente\n' +
    '4. Cuando termines, ejecuta el Paso 4: Generar Reporte Final.'
  );
}

function crearHojaRubrica(h, act, estudiantes) {
  // Título
  h.getRange('A1').setValue('Rúbrica: ' + act.nombre);
  h.getRange('A1').setFontSize(12).setFontWeight('bold').setBackground(C_VERDE_OSC).setFontColor('#FFFFFF');
  h.getRange(1, 1, 1, 3 + act.criterios.length + 1).merge();
  h.setRowHeight(1, 26);

  h.getRange('A2').setValue('Fecha: ' + act.fecha + '   |   Puntos máximos: ' + act.maxPts);
  h.getRange('A2').setFontSize(10).setFontStyle('italic').setFontColor('#555555');
  h.getRange(2, 1, 1, 3 + act.criterios.length + 1).merge();
  h.setRowHeight(2, 20);

  // Encabezados (fila 4)
  var FILA_H = 4;
  var headers = ['Apellido', 'Nombre', 'Email'];
  for (var ci = 0; ci < act.criterios.length; ci++) {
    headers.push(act.criterios[ci].nombre + '\n(Máx: ' + act.criterios[ci].max + ')');
  }
  headers.push('TOTAL');

  h.getRange(FILA_H, 1, 1, headers.length).setValues([headers]);
  h.getRange(FILA_H, 1, 1, 3).setBackground('#37474F').setFontColor('#FFFFFF').setFontWeight('bold');
  h.getRange(FILA_H, 4, 1, act.criterios.length).setBackground(C_VERDE_MED).setFontColor('#FFFFFF').setFontWeight('bold');
  h.getRange(FILA_H, 4 + act.criterios.length, 1, 1)
   .setBackground(C_VERDE_OSC).setFontColor('#FFFFFF').setFontWeight('bold');
  h.getRange(FILA_H, 1, 1, headers.length)
   .setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('middle');
  h.setRowHeight(FILA_H, 46);

  // Filas de estudiantes
  for (var i = 0; i < estudiantes.length; i++) {
    var fila = FILA_H + 1 + i;
    var bg   = i % 2 === 0 ? '#FFFFFF' : C_GRIS;
    var est  = estudiantes[i];

    h.getRange(fila, 1).setValue(est.apellido).setBackground(bg).setFontColor('#000000').setHorizontalAlignment('left');
    h.getRange(fila, 2).setValue(est.nombre).setBackground(bg).setFontColor('#000000').setHorizontalAlignment('left');
    h.getRange(fila, 3).setValue(est.email).setBackground(bg).setFontColor('#000000').setHorizontalAlignment('left');

    // Celdas de criterios (amarillo)
    var numCrit = act.criterios.length;
    if (numCrit > 0) {
      h.getRange(fila, 4, 1, numCrit).setBackground(C_AMARILLO).setFontColor('#000000').setHorizontalAlignment('center');
    }

    // Fórmula TOTAL
    var colTotal = 4 + numCrit;
    if (numCrit > 0) {
      var iniLetra = columnToLetter(4);
      var finLetra = columnToLetter(3 + numCrit);
      h.getRange(fila, colTotal).setFormula('=SUM(' + iniLetra + fila + ':' + finLetra + fila + ')')
       .setBackground(C_TOTAL_BG).setFontWeight('bold').setHorizontalAlignment('center');
    }
  }

  // Fila de máximos
  var filaTot = FILA_H + estudiantes.length + 2;
  h.getRange(filaTot, 1).setValue('PUNTAJE MÁXIMO').setFontWeight('bold').setBackground('#E3F2FD');
  for (var ci = 0; ci < act.criterios.length; ci++) {
    h.getRange(filaTot, 4 + ci).setValue(act.criterios[ci].max)
     .setFontWeight('bold').setBackground('#E3F2FD').setHorizontalAlignment('center');
  }
  h.getRange(filaTot, 4 + act.criterios.length).setValue(act.maxPts)
   .setFontWeight('bold').setBackground('#E3F2FD').setHorizontalAlignment('center');

  // Anchos
  h.setColumnWidth(1, 190);
  h.setColumnWidth(2, 190);
  h.setColumnWidth(3, 210);
  for (var c = 4; c <= headers.length; c++) h.setColumnWidth(c, 115);

  // Bordes
  h.getRange(FILA_H, 1, estudiantes.length + 1, headers.length)
   .setBorder(true, true, true, true, true, true, '#BDBDBD', SpreadsheetApp.BorderStyle.SOLID);

  h.setFrozenRows(FILA_H);
  h.setFrozenColumns(3);
}

// ============================================================
// PASO 4: GENERAR REPORTE FINAL
// ============================================================

function generarReporte() {
  var ui = SpreadsheetApp.getUi();
  var datosClassroom;
  try { datosClassroom = leerDatosClassroom(); }
  catch(e) { ui.alert('Error: ' + e.message); return; }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaPlantilla = ss.getSheetByName(HOJA_PLANTILLA);
  if (!hojaPlantilla) { ui.alert('Primero ejecuta el Paso 2: Generar Plantilla.'); return; }

  var config;
  try { config = leerConfigPlantilla(hojaPlantilla, datosClassroom.actividades); }
  catch(e) { ui.alert('Error al leer la plantilla: ' + e.message); return; }

  // Crear/recrear hoja REPORTE
  var hojaRep = ss.getSheetByName(HOJA_REPORTE);
  if (hojaRep) ss.deleteSheet(hojaRep);
  hojaRep = ss.insertSheet(HOJA_REPORTE);
  ss.setActiveSheet(hojaRep);
  ss.moveActiveSheet(1);

  // ── Leer puntajes de cada hoja de rúbrica ──
  // estructura: { actNombre: { 'APELLIDO|NOMBRE': { criterios: [], total: num } } }
  var punteosRub = {};
  for (var ci = 0; ci < config.length; ci++) {
    var actCfg = config[ci];
    if (!actCfg.tieneRubrica || actCfg.criterios.length === 0) continue;
    var nombreHoja = (PREFIJO_RUB + actCfg.nombre).substring(0, 100).replace(/[\/\\?\*\[\]:]/g, '_');
    var hojaRub = ss.getSheetByName(nombreHoja);
    if (!hojaRub) continue;

    var datosRub = hojaRub.getDataRange().getValues();
    var FILA_H_RUB = 3; // 0-indexed (fila 4)
    punteosRub[actCfg.nombre] = {};

    for (var r = FILA_H_RUB + 1; r < datosRub.length; r++) {
      var fila = datosRub[r];
      var apellido = String(fila[0]).trim();
      var nombre   = String(fila[1]).trim();
      if (!apellido && !nombre) continue;

      var key = apellido + '|' + nombre;
      var criVals = [];
      for (var k = 0; k < actCfg.criterios.length; k++) {
        var rawVal = fila[3 + k];
        criVals.push(rawVal !== '' && rawVal !== null ? (parseFloat(rawVal) || 0) : null);
      }
      var totalIdx = 3 + actCfg.criterios.length;
      var totalVal = fila[totalIdx] !== '' && fila[totalIdx] !== null ? (parseFloat(fila[totalIdx]) || 0) : null;
      punteosRub[actCfg.nombre][key] = { criterios: criVals, total: totalVal };
    }
  }

  // ── Construir lista de columnas del reporte ──
  var colsInfo = []; // { label, actNombre, tipo ('directa'|'criterio'|'total'), criterioIdx? }
  for (var ci = 0; ci < config.length; ci++) {
    var act = config[ci];
    if (!act.tieneRubrica || act.criterios.length === 0) {
      colsInfo.push({ label: act.nombre, actNombre: act.nombre, tipo: 'directa' });
    } else {
      for (var k = 0; k < act.criterios.length; k++) {
        colsInfo.push({
          label: act.nombre + '\n' + act.criterios[k].nombre,
          actNombre: act.nombre, tipo: 'criterio', criterioIdx: k
        });
      }
      colsInfo.push({ label: act.nombre + '\nTOTAL', actNombre: act.nombre, tipo: 'total' });
    }
  }

  var FILA_H = 4;
  var numFijos = 4; // Apellido, Nombre, Email, Promedio
  var totalCols = numFijos + colsInfo.length;

  // ── Fila 1: Título ──
  hojaRep.getRange(1, 1, 1, totalCols).merge()
    .setValue('Reporte Final — ' + datosClassroom.curso + '   ' + datosClassroom.seccion)
    .setFontSize(14).setFontWeight('bold')
    .setBackground(C_AZUL_OSC).setFontColor('#FFFFFF')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  hojaRep.setRowHeight(1, 32);

  hojaRep.getRange(2, 1, 1, totalCols).merge()
    .setValue('Generado: ' + new Date().toLocaleString('es-GT'))
    .setFontSize(10).setFontStyle('italic').setFontColor('#888888');
  hojaRep.setRowHeight(2, 18);

  // ── Fila 4: Encabezados ──
  var encFijos = ['Apellido', 'Nombre', 'Email', 'Promedio'];
  var encActs  = colsInfo.map(function(c) { return c.label; });
  var todosEnc = encFijos.concat(encActs);

  hojaRep.getRange(FILA_H, 1, 1, todosEnc.length).setValues([todosEnc]);
  hojaRep.getRange(FILA_H, 1, 1, 4)
    .setBackground('#37474F').setFontColor('#FFFFFF').setFontWeight('bold');

  for (var ci = 0; ci < colsInfo.length; ci++) {
    var col = ci + 5;
    var tipo = colsInfo[ci].tipo;
    var bgH = tipo === 'directa' ? C_AZUL_OSC : (tipo === 'criterio' ? C_VERDE_MED : C_VERDE_OSC);
    hojaRep.getRange(FILA_H, col).setBackground(bgH).setFontColor('#FFFFFF').setFontWeight('bold');
  }

  hojaRep.getRange(FILA_H, 1, 1, totalCols)
    .setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('middle');
  hojaRep.setRowHeight(FILA_H, 52);

  // ── Filas de estudiantes ──
  var ests = datosClassroom.estudiantes;
  for (var ei = 0; ei < ests.length; ei++) {
    var fila = FILA_H + 1 + ei;
    var bg   = ei % 2 === 0 ? '#FFFFFF' : C_GRIS;
    var est  = ests[ei];

    hojaRep.getRange(fila, 1).setValue(est.apellido).setBackground(bg).setFontColor('#000000').setHorizontalAlignment('left');
    hojaRep.getRange(fila, 2).setValue(est.nombre).setBackground(bg).setFontColor('#000000').setHorizontalAlignment('left');
    hojaRep.getRange(fila, 3).setValue(est.email).setBackground(bg).setFontColor('#000000').setHorizontalAlignment('left');
    hojaRep.getRange(fila, 4).setValue(est.porcentaje).setBackground(bg).setFontColor('#000000').setHorizontalAlignment('center');

    for (var ci = 0; ci < colsInfo.length; ci++) {
      var col  = ci + 5;
      var info = colsInfo[ci];
      var key  = est.apellido + '|' + est.nombre;
      var valor = null;

      if (info.tipo === 'directa') {
        valor = est.punteos[info.actNombre];
        hojaRep.getRange(fila, col).setBackground(ei % 2 === 0 ? '#E3F2FD' : '#BBDEFB');
      } else if (info.tipo === 'criterio') {
        var rubData = punteosRub[info.actNombre];
        if (rubData && rubData[key]) valor = rubData[key].criterios[info.criterioIdx];
        hojaRep.getRange(fila, col).setBackground(ei % 2 === 0 ? '#E8F5E9' : '#C8E6C9');
      } else { // total
        var rubData = punteosRub[info.actNombre];
        if (rubData && rubData[key]) valor = rubData[key].total;
        hojaRep.getRange(fila, col).setBackground(ei % 2 === 0 ? '#F1F8E9' : '#DCEDC8').setFontWeight('bold');
      }

      if (valor !== null && valor !== undefined && valor !== '') {
        hojaRep.getRange(fila, col).setValue(valor).setHorizontalAlignment('center');
      }
    }
  }

  // ── Anchos ──
  hojaRep.setColumnWidth(1, 185);
  hojaRep.setColumnWidth(2, 185);
  hojaRep.setColumnWidth(3, 210);
  hojaRep.setColumnWidth(4, 80);
  for (var c = 5; c <= totalCols; c++) hojaRep.setColumnWidth(c, 110);

  // ── Bordes ──
  hojaRep.getRange(FILA_H, 1, ests.length + 1, totalCols)
    .setBorder(true, true, true, true, true, true, '#BDBDBD', SpreadsheetApp.BorderStyle.SOLID);

  hojaRep.setFrozenRows(FILA_H);
  hojaRep.setFrozenColumns(4);

  ui.alert(
    'Reporte Final generado.\n\n' +
    'Columnas azules = puntaje directo de Classroom\n' +
    'Columnas verde claro = criterios de rúbrica\n' +
    'Columnas verde oscuro = TOTAL de rúbrica'
  );
}

// ============================================================
// INSTRUCCIONES
// ============================================================

function mostrarInstrucciones() {
  var html = HtmlService.createHtmlOutput(
    '<style>body{font-family:Arial,sans-serif;padding:18px;font-size:13px;line-height:1.5}' +
    'h2{color:#1565C0;margin-top:0;font-size:16px}' +
    '.paso{background:#F5F5F5;padding:10px 14px;border-radius:4px;margin:10px 0;border-left:4px solid #1565C0}' +
    '.paso.verde{border-left-color:#2E7D32}' +
    '.paso h3{margin:0 0 4px;font-size:13px;color:#1565C0}' +
    '.paso.verde h3{color:#2E7D32}' +
    'p{margin:4px 0}' +
    '</style>' +
    '<h2>Rúbricas Classroom — Instrucciones</h2>' +
    '<div class="paso"><h3>Paso 1: Importar CSV</h3>' +
    '<p>Exporta las calificaciones desde Google Classroom como CSV.<br>' +
    'Abre el archivo en el Bloc de notas, copia todo y usa el menú para importarlo.</p></div>' +
    '<div class="paso"><h3>Paso 2: Generar Plantilla</h3>' +
    '<p>Se crea la hoja <b>PLANTILLA</b>.<br>' +
    'Escribe <b>SI</b> o <b>NO</b> para cada actividad en la columna "¿Tiene Rúbrica?".<br>' +
    'Para las que tienen rúbrica, ingresa los criterios y sus puntos máximos en las celdas amarillas.</p></div>' +
    '<div class="paso verde"><h3>Paso 3: Crear Hojas de Punteos</h3>' +
    '<p>Se crean hojas <b>RUB_[nombre]</b> para cada actividad con rúbrica.<br>' +
    'Ingresa los puntajes por criterio en las celdas amarillas.<br>' +
    'La columna TOTAL se calcula sola.</p></div>' +
    '<div class="paso verde"><h3>Paso 4: Generar Reporte Final</h3>' +
    '<p>Se crea la hoja <b>REPORTE</b> con todo consolidado.<br>' +
    'Columnas azules = puntaje directo · Columnas verdes = criterios de rúbrica.</p></div>'
  ).setWidth(480).setHeight(400).setTitle('Instrucciones');
  SpreadsheetApp.getUi().showModalDialog(html, 'Instrucciones');
}

// ============================================================
// UTILIDAD
// ============================================================

function columnToLetter(col) {
  var letter = '';
  while (col > 0) {
    var rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}
