// ============================================================
// CAMPEONATO DE FÚTBOL — Web App + Google Sheets
// Backend de Apps Script para la tabla de posiciones interactiva
// ============================================================

const SHEET_EQUIPOS = 'Equipos';
const SHEET_PARTIDOS = 'Partidos';
const SHEET_CONFIG = 'Config';

const PUNTOS_VICTORIA = 3;
const PUNTOS_EMPATE = 1;
const PUNTOS_DERROTA = 0;

// ============================================================
// MENÚ
// ============================================================

function onOpenCampeonato() {
  SpreadsheetApp.getUi()
    .createMenu('⚽ Campeonato')
    .addItem('1. Crear hojas base', 'crearHojasCampeonato')
    .addItem('2. Cargar datos de ejemplo', 'cargarDatosEjemplo')
    .addSeparator()
    .addItem('Abrir vista previa', 'abrirVistaPrevia')
    .addItem('Instrucciones para publicar', 'mostrarInstruccionesPublicar')
    .addToUi();
}

function crearHojasCampeonato() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let hojaEquipos = ss.getSheetByName(SHEET_EQUIPOS);
  if (!hojaEquipos) {
    hojaEquipos = ss.insertSheet(SHEET_EQUIPOS);
    hojaEquipos.getRange('A1:D1').setValues([['ID', 'Nombre', 'Grado/Sección', 'Escudo (URL)']]);
    hojaEquipos.getRange('A1:D1')
      .setFontWeight('bold')
      .setBackground('#1e40af')
      .setFontColor('#ffffff')
      .setHorizontalAlignment('center');
    hojaEquipos.setFrozenRows(1);
    hojaEquipos.setColumnWidth(1, 60);
    hojaEquipos.setColumnWidth(2, 220);
    hojaEquipos.setColumnWidth(3, 160);
    hojaEquipos.setColumnWidth(4, 320);
  }

  let hojaPartidos = ss.getSheetByName(SHEET_PARTIDOS);
  if (!hojaPartidos) {
    hojaPartidos = ss.insertSheet(SHEET_PARTIDOS);
    hojaPartidos.getRange('A1:G1').setValues([[
      'Jornada', 'Fecha', 'Equipo Local', 'Goles Local',
      'Equipo Visitante', 'Goles Visitante', 'Estado'
    ]]);
    hojaPartidos.getRange('A1:G1')
      .setFontWeight('bold')
      .setBackground('#059669')
      .setFontColor('#ffffff')
      .setHorizontalAlignment('center');
    hojaPartidos.setFrozenRows(1);
    hojaPartidos.setColumnWidth(1, 80);
    hojaPartidos.setColumnWidth(2, 110);
    hojaPartidos.setColumnWidth(3, 180);
    hojaPartidos.setColumnWidth(4, 100);
    hojaPartidos.setColumnWidth(5, 180);
    hojaPartidos.setColumnWidth(6, 110);
    hojaPartidos.setColumnWidth(7, 110);

    // Validación de datos en Estado
    const reglaEstado = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Jugado', 'Pendiente', 'Suspendido'], true)
      .build();
    hojaPartidos.getRange('G2:G1000').setDataValidation(reglaEstado);
  }

  let hojaConfig = ss.getSheetByName(SHEET_CONFIG);
  if (!hojaConfig) {
    hojaConfig = ss.insertSheet(SHEET_CONFIG);
    hojaConfig.getRange('A1:B1').setValues([['Parámetro', 'Valor']]);
    hojaConfig.getRange('A2:B7').setValues([
      ['Nombre del Campeonato', 'Campeonato Escolar de Fútbol'],
      ['Subtítulo', 'Temporada 2026'],
      ['Color Primario', '#1e40af'],
      ['Color Secundario', '#059669'],
      ['Logo (URL)', ''],
      ['Clasifican (primeros N)', 4]
    ]);
    hojaConfig.getRange('A1:B1')
      .setFontWeight('bold')
      .setBackground('#374151')
      .setFontColor('#ffffff');
    hojaConfig.setColumnWidth(1, 220);
    hojaConfig.setColumnWidth(2, 320);
  }

  SpreadsheetApp.getUi().alert(
    '✅ Hojas creadas/verificadas',
    'Se prepararon las hojas:\n\n• Equipos\n• Partidos\n• Config\n\n' +
    'Registra tus equipos y partidos y usa la Web App para visualizar la tabla.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function cargarDatosEjemplo() {
  crearHojasCampeonato();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaEquipos = ss.getSheetByName(SHEET_EQUIPOS);
  const hojaPartidos = ss.getSheetByName(SHEET_PARTIDOS);

  hojaEquipos.getRange('A2:D100').clearContent();
  hojaPartidos.getRange('A2:G200').clearContent();

  const equipos = [
    [1, 'Leones FC',   '6to A', ''],
    [2, 'Tigres Rojos','6to B', ''],
    [3, 'Águilas',     '5to A', ''],
    [4, 'Halcones',    '5to B', ''],
    [5, 'Panteras',    '4to A', ''],
    [6, 'Cóndores',    '4to B', '']
  ];
  hojaEquipos.getRange(2, 1, equipos.length, 4).setValues(equipos);

  const partidos = [
    [1, new Date(2026, 3, 5),  'Leones FC',    3, 'Tigres Rojos', 1, 'Jugado'],
    [1, new Date(2026, 3, 5),  'Águilas',      2, 'Halcones',     2, 'Jugado'],
    [1, new Date(2026, 3, 5),  'Panteras',     1, 'Cóndores',     0, 'Jugado'],
    [2, new Date(2026, 3, 12), 'Tigres Rojos', 2, 'Águilas',      0, 'Jugado'],
    [2, new Date(2026, 3, 12), 'Halcones',     1, 'Panteras',     3, 'Jugado'],
    [2, new Date(2026, 3, 12), 'Cóndores',     1, 'Leones FC',    2, 'Jugado'],
    [3, new Date(2026, 3, 19), 'Leones FC',    4, 'Águilas',      1, 'Jugado'],
    [3, new Date(2026, 3, 19), 'Tigres Rojos', 3, 'Panteras',     3, 'Jugado'],
    [3, new Date(2026, 3, 19), 'Halcones',     0, 'Cóndores',     2, 'Jugado'],
    [4, new Date(2026, 3, 26), 'Leones FC',    '', 'Halcones',    '', 'Pendiente'],
    [4, new Date(2026, 3, 26), 'Tigres Rojos', '', 'Cóndores',    '', 'Pendiente'],
    [4, new Date(2026, 3, 26), 'Águilas',      '', 'Panteras',    '', 'Pendiente']
  ];
  hojaPartidos.getRange(2, 1, partidos.length, 7).setValues(partidos);
  hojaPartidos.getRange(2, 2, partidos.length, 1).setNumberFormat('dd/MM/yyyy');

  SpreadsheetApp.getUi().alert('✅ Datos de ejemplo cargados en Equipos y Partidos.');
}

function abrirVistaPrevia() {
  const html = HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setWidth(1100)
    .setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, 'Tabla de Posiciones');
}

function mostrarInstruccionesPublicar() {
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family: Arial, sans-serif; padding: 20px; line-height: 1.6;">' +
      '<h2 style="margin-top:0;color:#1e40af;">Publicar como Web App</h2>' +
      '<ol>' +
        '<li>En el editor de Apps Script haz clic en <b>Implementar → Nueva implementación</b>.</li>' +
        '<li>Selecciona tipo <b>Aplicación web</b>.</li>' +
        '<li>Ejecutar como: <b>Tu cuenta</b>.</li>' +
        '<li>Quién tiene acceso: <b>Cualquier usuario</b>.</li>' +
        '<li>Copia la <b>URL de la Web App</b>.</li>' +
      '</ol>' +
      '<h3 style="color:#059669;">Insertar en Google Sites</h3>' +
      '<ol>' +
        '<li>Abre tu sitio en <b>sites.google.com</b>.</li>' +
        '<li><b>Insertar → Insertar → URL</b> y pega la URL de la Web App.</li>' +
        '<li>Guarda y publica.</li>' +
      '</ol>' +
      '<p style="background:#fef3c7;padding:10px;border-radius:6px;">' +
        '<b>Tip:</b> cuando agregues o edites partidos en la hoja, la tabla se actualiza ' +
        'al recargar la Web App o tocar el botón <b>Actualizar</b>.' +
      '</p>' +
    '</div>'
  ).setWidth(520).setHeight(460);
  SpreadsheetApp.getUi().showModalDialog(html, 'Instrucciones');
}

// ============================================================
// WEB APP
// ============================================================

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Tabla de Posiciones — Campeonato')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
// API (consumida por el cliente vía google.script.run)
// ============================================================

function getDatosCampeonato() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const equipos = leerEquipos(ss.getSheetByName(SHEET_EQUIPOS));
  const partidos = leerPartidos(ss.getSheetByName(SHEET_PARTIDOS));
  const config = leerConfig(ss.getSheetByName(SHEET_CONFIG));
  const tabla = calcularTabla(equipos, partidos);
  const jornadas = [...new Set(partidos.map(p => p.jornada))]
    .filter(j => j > 0)
    .sort((a, b) => a - b);

  return {
    config,
    equipos,
    partidos,
    tabla,
    jornadas,
    actualizado: new Date().toISOString()
  };
}

function leerEquipos(hoja) {
  if (!hoja || hoja.getLastRow() < 2) return [];
  const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 4).getValues();
  return datos
    .filter(r => r[1])
    .map(r => ({
      id: r[0],
      nombre: String(r[1]).trim(),
      grado: String(r[2] || '').trim(),
      escudo: String(r[3] || '').trim()
    }));
}

function leerPartidos(hoja) {
  if (!hoja || hoja.getLastRow() < 2) return [];
  const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 7).getValues();
  return datos
    .filter(r => r[2] && r[4])
    .map((r, i) => {
      const gl = r[3] === '' || r[3] === null ? null : Number(r[3]);
      const gv = r[5] === '' || r[5] === null ? null : Number(r[5]);
      const estadoRaw = String(r[6] || '').trim();
      const estado = estadoRaw || (gl !== null && gv !== null ? 'Jugado' : 'Pendiente');
      return {
        id: i + 1,
        jornada: Number(r[0]) || 0,
        fecha: r[1] instanceof Date ? r[1].toISOString() : String(r[1] || ''),
        local: String(r[2]).trim(),
        golesLocal: gl,
        visitante: String(r[4]).trim(),
        golesVisitante: gv,
        estado: estado
      };
    });
}

function leerConfig(hoja) {
  const defaults = {
    nombre: 'Campeonato de Fútbol',
    subtitulo: '',
    colorPrimario: '#1e40af',
    colorSecundario: '#059669',
    logo: '',
    clasifican: 4
  };
  if (!hoja || hoja.getLastRow() < 2) return defaults;
  const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 2).getValues();
  const map = {};
  datos.forEach(r => { if (r[0]) map[String(r[0]).trim()] = r[1]; });
  return {
    nombre: map['Nombre del Campeonato'] || defaults.nombre,
    subtitulo: map['Subtítulo'] || defaults.subtitulo,
    colorPrimario: map['Color Primario'] || defaults.colorPrimario,
    colorSecundario: map['Color Secundario'] || defaults.colorSecundario,
    logo: map['Logo (URL)'] || defaults.logo,
    clasifican: Number(map['Clasifican (primeros N)']) || defaults.clasifican
  };
}

function calcularTabla(equipos, partidos) {
  const tabla = {};
  equipos.forEach(e => {
    tabla[e.nombre] = {
      equipo: e.nombre,
      grado: e.grado,
      escudo: e.escudo,
      pj: 0, pg: 0, pe: 0, pp: 0,
      gf: 0, gc: 0, dg: 0, pts: 0,
      forma: []
    };
  });

  partidos
    .slice()
    .sort((a, b) => a.jornada - b.jornada)
    .forEach(p => {
      if (p.estado !== 'Jugado') return;
      if (p.golesLocal === null || p.golesVisitante === null) return;
      const L = tabla[p.local];
      const V = tabla[p.visitante];
      if (!L || !V) return;

      L.pj++; V.pj++;
      L.gf += p.golesLocal; L.gc += p.golesVisitante;
      V.gf += p.golesVisitante; V.gc += p.golesLocal;

      if (p.golesLocal > p.golesVisitante) {
        L.pg++; L.pts += PUNTOS_VICTORIA;
        V.pp++; V.pts += PUNTOS_DERROTA;
        L.forma.push('G'); V.forma.push('P');
      } else if (p.golesLocal < p.golesVisitante) {
        V.pg++; V.pts += PUNTOS_VICTORIA;
        L.pp++; L.pts += PUNTOS_DERROTA;
        L.forma.push('P'); V.forma.push('G');
      } else {
        L.pe++; L.pts += PUNTOS_EMPATE;
        V.pe++; V.pts += PUNTOS_EMPATE;
        L.forma.push('E'); V.forma.push('E');
      }
    });

  return Object.values(tabla)
    .map(t => ({ ...t, dg: t.gf - t.gc, forma: t.forma.slice(-5) }))
    .sort((a, b) => {
      if (b.pts !== a.pts) return b.pts - a.pts;
      if (b.dg !== a.dg) return b.dg - a.dg;
      if (b.gf !== a.gf) return b.gf - a.gf;
      return a.equipo.localeCompare(b.equipo);
    })
    .map((t, i) => ({ ...t, posicion: i + 1 }));
}
