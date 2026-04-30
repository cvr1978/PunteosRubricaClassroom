// ============================================================
// CAMPEONATO DE FÚTBOL — Web App + Google Sheets
// Backend de Apps Script para la tabla de posiciones interactiva
// ============================================================

const SHEET_EQUIPOS = 'Equipos';
const SHEET_PARTIDOS = 'Partidos';
const SHEET_CONFIG = 'Config';
const SHEET_GOLEADORES = 'Goleadores';
const SHEET_ELIMINATORIAS = 'Eliminatorias';

const PUNTOS_VICTORIA = 3;
const PUNTOS_EMPATE = 1;
const PUNTOS_DERROTA = 0;

const FASES_ORDEN = ['Octavos', 'Cuartos', 'Semifinal', 'Final', 'Tercer Lugar'];

// ============================================================
// MENÚ
// ============================================================

function onOpenCampeonato() {
  SpreadsheetApp.getUi()
    .createMenu('⚽ Campeonato')
    .addItem('1. Crear hojas base', 'crearHojasCampeonato')
    .addItem('2. Cargar datos de ejemplo', 'cargarDatosEjemplo')
    .addSeparator()
    .addItem('Refrescar dropdown de equipos', 'aplicarValidacionEquipos')
    .addItem('Abrir vista previa', 'abrirVistaPrevia')
    .addItem('Instrucciones para publicar', 'mostrarInstruccionesPublicar')
    .addToUi();
}

function crearHojasCampeonato() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let hojaEquipos = ss.getSheetByName(SHEET_EQUIPOS);
  if (!hojaEquipos) {
    hojaEquipos = ss.insertSheet(SHEET_EQUIPOS);
    hojaEquipos.getRange('A1:E1').setValues([[
      'ID', 'Nombre', 'Grado/Sección', 'Grupo', 'Escudo (URL)'
    ]]);
    hojaEquipos.getRange('A1:E1')
      .setFontWeight('bold')
      .setBackground('#1e40af')
      .setFontColor('#ffffff')
      .setHorizontalAlignment('center');
    hojaEquipos.setFrozenRows(1);
    hojaEquipos.setColumnWidth(1, 60);
    hojaEquipos.setColumnWidth(2, 220);
    hojaEquipos.setColumnWidth(3, 140);
    hojaEquipos.setColumnWidth(4, 130);
    hojaEquipos.setColumnWidth(5, 300);

    const reglaGrupo = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Grupo A', 'Grupo B', 'Femenino'], true)
      .setAllowInvalid(true)
      .build();
    hojaEquipos.getRange('D2:D1000').setDataValidation(reglaGrupo);
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
    hojaConfig.getRange('A2:B8').setValues([
      ['Nombre del Campeonato', 'Campeonato Escolar de Fútbol'],
      ['Subtítulo', 'Temporada 2026'],
      ['Color Primario', '#1e40af'],
      ['Color Secundario', '#059669'],
      ['Logo (URL)', ''],
      ['Clasifican (primeros N)', 2],
      ['Orden de grupos', 'Grupo A, Grupo B, Femenino']
    ]);
    hojaConfig.getRange('A1:B1')
      .setFontWeight('bold')
      .setBackground('#374151')
      .setFontColor('#ffffff');
    hojaConfig.setColumnWidth(1, 220);
    hojaConfig.setColumnWidth(2, 320);
  }

  let hojaGoleadores = ss.getSheetByName(SHEET_GOLEADORES);
  if (!hojaGoleadores) {
    hojaGoleadores = ss.insertSheet(SHEET_GOLEADORES);
    hojaGoleadores.getRange('A1:C1').setValues([['Jugador', 'Equipo', 'Goles']]);
    hojaGoleadores.getRange('A1:C1')
      .setFontWeight('bold')
      .setBackground('#b45309')
      .setFontColor('#ffffff')
      .setHorizontalAlignment('center');
    hojaGoleadores.setFrozenRows(1);
    hojaGoleadores.setColumnWidth(1, 240);
    hojaGoleadores.setColumnWidth(2, 180);
    hojaGoleadores.setColumnWidth(3, 90);
  }

  let hojaElim = ss.getSheetByName(SHEET_ELIMINATORIAS);
  if (!hojaElim) {
    hojaElim = ss.insertSheet(SHEET_ELIMINATORIAS);
    hojaElim.getRange('A1:J1').setValues([[
      'Fase', 'Rama', 'Fecha',
      'Equipo 1', 'Goles 1',
      'Equipo 2', 'Goles 2',
      'Penales 1', 'Penales 2',
      'Estado'
    ]]);
    hojaElim.getRange('A1:J1')
      .setFontWeight('bold')
      .setBackground('#7c2d12')
      .setFontColor('#ffffff')
      .setHorizontalAlignment('center');
    hojaElim.setFrozenRows(1);
    hojaElim.setColumnWidth(1, 110);
    hojaElim.setColumnWidth(2, 120);
    hojaElim.setColumnWidth(3, 110);
    hojaElim.setColumnWidth(4, 170);
    hojaElim.setColumnWidth(5, 90);
    hojaElim.setColumnWidth(6, 170);
    hojaElim.setColumnWidth(7, 90);
    hojaElim.setColumnWidth(8, 90);
    hojaElim.setColumnWidth(9, 90);
    hojaElim.setColumnWidth(10, 110);

    const reglaFase = SpreadsheetApp.newDataValidation()
      .requireValueInList(FASES_ORDEN, true)
      .setAllowInvalid(false)
      .build();
    hojaElim.getRange('A2:A1000').setDataValidation(reglaFase);

    const reglaEstado = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Jugado', 'Pendiente', 'Suspendido'], true)
      .build();
    hojaElim.getRange('J2:J1000').setDataValidation(reglaEstado);
  }

  aplicarValidacionEquipos();

  SpreadsheetApp.getUi().alert(
    '✅ Hojas creadas/verificadas',
    'Se prepararon las hojas:\n\n• Equipos\n• Partidos\n• Config\n• Goleadores\n• Eliminatorias\n\n' +
    'Registra tus equipos y partidos y usa la Web App para visualizar la tabla.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// Aplica validación de datos (dropdown) con la lista de equipos en
// las columnas relevantes de Partidos, Goleadores y Eliminatorias.
function aplicarValidacionEquipos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaEquipos = ss.getSheetByName(SHEET_EQUIPOS);
  if (!hojaEquipos) return;

  const rangoEquipos = hojaEquipos.getRange('B2:B1000');
  const regla = SpreadsheetApp.newDataValidation()
    .requireValueInRange(rangoEquipos, true)
    .setAllowInvalid(true)
    .build();

  const hojaPartidos = ss.getSheetByName(SHEET_PARTIDOS);
  if (hojaPartidos) {
    hojaPartidos.getRange('C2:C1000').setDataValidation(regla);
    hojaPartidos.getRange('E2:E1000').setDataValidation(regla);
  }

  const hojaGoleadores = ss.getSheetByName(SHEET_GOLEADORES);
  if (hojaGoleadores) {
    hojaGoleadores.getRange('B2:B1000').setDataValidation(regla);
  }

  const hojaElim = ss.getSheetByName(SHEET_ELIMINATORIAS);
  if (hojaElim) {
    hojaElim.getRange('D2:D1000').setDataValidation(regla);
    hojaElim.getRange('F2:F1000').setDataValidation(regla);
  }
}

function cargarDatosEjemplo() {
  crearHojasCampeonato();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaEquipos = ss.getSheetByName(SHEET_EQUIPOS);
  const hojaPartidos = ss.getSheetByName(SHEET_PARTIDOS);

  hojaEquipos.getRange('A2:E100').clearContent();
  hojaPartidos.getRange('A2:G200').clearContent();

  const equipos = [
    // Grupo A (masculino)
    [1, 'Leones FC',    '6to A', 'Grupo A',  ''],
    [2, 'Tigres Rojos', '6to B', 'Grupo A',  ''],
    [3, 'Águilas',      '5to A', 'Grupo A',  ''],
    [4, 'Halcones',     '5to B', 'Grupo A',  ''],
    // Grupo B (masculino)
    [5, 'Panteras',     '4to A', 'Grupo B',  ''],
    [6, 'Cóndores',     '4to B', 'Grupo B',  ''],
    [7, 'Lobos',        '3ro A', 'Grupo B',  ''],
    [8, 'Toros',        '3ro B', 'Grupo B',  ''],
    // Femenino
    [9,  'Estrellas',   '6to A', 'Femenino', ''],
    [10, 'Cometas',     '6to B', 'Femenino', ''],
    [11, 'Gacelas',     '5to A', 'Femenino', ''],
    [12, 'Fénix',       '5to B', 'Femenino', '']
  ];
  hojaEquipos.getRange(2, 1, equipos.length, 5).setValues(equipos);

  const partidos = [
    // Jornada 1
    [1, new Date(2026, 3, 5),  'Leones FC',    3, 'Tigres Rojos', 1, 'Jugado'],
    [1, new Date(2026, 3, 5),  'Águilas',      2, 'Halcones',     2, 'Jugado'],
    [1, new Date(2026, 3, 5),  'Panteras',     1, 'Cóndores',     0, 'Jugado'],
    [1, new Date(2026, 3, 5),  'Lobos',        2, 'Toros',        2, 'Jugado'],
    [1, new Date(2026, 3, 5),  'Estrellas',    3, 'Cometas',      2, 'Jugado'],
    [1, new Date(2026, 3, 5),  'Gacelas',      1, 'Fénix',        1, 'Jugado'],
    // Jornada 2
    [2, new Date(2026, 3, 12), 'Leones FC',    2, 'Águilas',      0, 'Jugado'],
    [2, new Date(2026, 3, 12), 'Tigres Rojos', 1, 'Halcones',     3, 'Jugado'],
    [2, new Date(2026, 3, 12), 'Panteras',     2, 'Lobos',        1, 'Jugado'],
    [2, new Date(2026, 3, 12), 'Cóndores',     0, 'Toros',        2, 'Jugado'],
    [2, new Date(2026, 3, 12), 'Estrellas',    4, 'Gacelas',      1, 'Jugado'],
    [2, new Date(2026, 3, 12), 'Cometas',      2, 'Fénix',        0, 'Jugado'],
    // Jornada 3
    [3, new Date(2026, 3, 19), 'Leones FC',    1, 'Halcones',     1, 'Jugado'],
    [3, new Date(2026, 3, 19), 'Tigres Rojos', 2, 'Águilas',      2, 'Jugado'],
    [3, new Date(2026, 3, 19), 'Panteras',     3, 'Toros',        1, 'Jugado'],
    [3, new Date(2026, 3, 19), 'Cóndores',     1, 'Lobos',        1, 'Jugado'],
    [3, new Date(2026, 3, 19), 'Estrellas',    2, 'Fénix',        2, 'Jugado'],
    [3, new Date(2026, 3, 19), 'Cometas',      1, 'Gacelas',      3, 'Jugado'],
    // Jornada 4 (pendiente — vueltas)
    [4, new Date(2026, 3, 26), 'Tigres Rojos', '', 'Leones FC',   '', 'Pendiente'],
    [4, new Date(2026, 3, 26), 'Halcones',     '', 'Águilas',     '', 'Pendiente'],
    [4, new Date(2026, 3, 26), 'Cóndores',     '', 'Panteras',    '', 'Pendiente'],
    [4, new Date(2026, 3, 26), 'Toros',        '', 'Lobos',       '', 'Pendiente'],
    [4, new Date(2026, 3, 26), 'Cometas',      '', 'Estrellas',   '', 'Pendiente'],
    [4, new Date(2026, 3, 26), 'Fénix',        '', 'Gacelas',     '', 'Pendiente']
  ];
  hojaPartidos.getRange(2, 1, partidos.length, 7).setValues(partidos);
  hojaPartidos.getRange(2, 2, partidos.length, 1).setNumberFormat('dd/MM/yyyy');

  const hojaGoleadores = ss.getSheetByName(SHEET_GOLEADORES);
  if (hojaGoleadores) {
    hojaGoleadores.getRange('A2:C200').clearContent();
    const goleadores = [
      ['Carlos Méndez',    'Leones FC',    5],
      ['Luis Hernández',   'Leones FC',    3],
      ['Diego Pérez',      'Tigres Rojos', 2],
      ['Andrés López',     'Águilas',      4],
      ['Pablo Ramírez',    'Halcones',     3],
      ['José Castillo',    'Halcones',     2],
      ['Mario García',     'Panteras',     4],
      ['Javier Morales',   'Cóndores',     1],
      ['Kevin Torres',     'Lobos',        3],
      ['Daniel Cruz',      'Toros',        2],
      ['María Gómez',      'Estrellas',    6],
      ['Ana Ruiz',         'Estrellas',    3],
      ['Sofía Álvarez',    'Cometas',      3],
      ['Carolina Díaz',    'Gacelas',      4],
      ['Valentina Rojas',  'Fénix',        3]
    ];
    hojaGoleadores.getRange(2, 1, goleadores.length, 3).setValues(goleadores);
  }

  const hojaElim = ss.getSheetByName(SHEET_ELIMINATORIAS);
  if (hojaElim) {
    hojaElim.getRange('A2:J200').clearContent();
    const elim = [
      ['Semifinal',    'Masculino', new Date(2026, 4, 3),  'Leones FC',  '', 'Toros',     '', '', '', 'Pendiente'],
      ['Semifinal',    'Masculino', new Date(2026, 4, 3),  'Halcones',   '', 'Panteras',  '', '', '', 'Pendiente'],
      ['Semifinal',    'Femenino',  new Date(2026, 4, 3),  'Estrellas',  '', 'Fénix',     '', '', '', 'Pendiente'],
      ['Semifinal',    'Femenino',  new Date(2026, 4, 3),  'Gacelas',    '', 'Cometas',   '', '', '', 'Pendiente'],
      ['Final',        'Masculino', new Date(2026, 4, 10), '',           '', '',          '', '', '', 'Pendiente'],
      ['Tercer Lugar', 'Masculino', new Date(2026, 4, 10), '',           '', '',          '', '', '', 'Pendiente'],
      ['Final',        'Femenino',  new Date(2026, 4, 10), '',           '', '',          '', '', '', 'Pendiente'],
      ['Tercer Lugar', 'Femenino',  new Date(2026, 4, 10), '',           '', '',          '', '', '', 'Pendiente']
    ];
    hojaElim.getRange(2, 1, elim.length, 10).setValues(elim);
    hojaElim.getRange(2, 3, elim.length, 1).setNumberFormat('dd/MM/yyyy');
  }

  SpreadsheetApp.getUi().alert('✅ Datos de ejemplo cargados en Equipos, Partidos, Goleadores y Eliminatorias.');
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
  const config = leerConfig(ss.getSheetByName(SHEET_CONFIG));
  const partidos = leerPartidos(ss.getSheetByName(SHEET_PARTIDOS), equipos);
  const goleadores = leerGoleadores(ss.getSheetByName(SHEET_GOLEADORES), equipos);
  const eliminatorias = leerEliminatorias(ss.getSheetByName(SHEET_ELIMINATORIAS), equipos);
  const tablas = calcularTablasPorGrupo(equipos, partidos);
  const grupos = ordenarGrupos(Object.keys(tablas), config.ordenGrupos);
  const jornadas = [...new Set(partidos.map(p => p.jornada))]
    .filter(j => j > 0)
    .sort((a, b) => a - b);

  return {
    config,
    equipos,
    partidos,
    tablas,
    grupos,
    jornadas,
    goleadores,
    eliminatorias,
    actualizado: new Date().toISOString()
  };
}

function ordenarGrupos(detectados, orden) {
  const result = [];
  (orden || []).forEach(g => {
    if (detectados.indexOf(g) !== -1 && result.indexOf(g) === -1) result.push(g);
  });
  detectados.sort().forEach(g => {
    if (result.indexOf(g) === -1) result.push(g);
  });
  return result;
}

function leerEquipos(hoja) {
  if (!hoja || hoja.getLastRow() < 2) return [];
  const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 5).getValues();
  return datos
    .filter(r => r[1])
    .map(r => ({
      id: r[0],
      nombre: String(r[1]).trim(),
      grado: String(r[2] || '').trim(),
      grupo: String(r[3] || '').trim() || 'General',
      escudo: String(r[4] || '').trim()
    }));
}

function leerGoleadores(hoja, equipos) {
  if (!hoja || hoja.getLastRow() < 2) return [];
  const equipoInfo = {};
  (equipos || []).forEach(e => {
    const key = String(e.nombre).trim().toLowerCase();
    if (key) equipoInfo[key] = {
      nombre: e.nombre, grupo: e.grupo, escudo: e.escudo, grado: e.grado
    };
  });
  const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 3).getValues();

  // Acumular goles por jugador+equipo (insensible a mayúsculas/espacios)
  const acumulado = {};
  datos
    .filter(r => r[0] && r[1])
    .forEach(r => {
      const equipoRaw = String(r[1]).trim();
      const info = equipoInfo[equipoRaw.toLowerCase()] || {};
      const jugador = String(r[0]).trim();
      const equipoCanonico = info.nombre || equipoRaw;
      const key = jugador.toLowerCase() + '||' + equipoCanonico.toLowerCase();
      const goles = Number(r[2]) || 0;
      if (!acumulado[key]) {
        acumulado[key] = {
          jugador: jugador,
          equipo: equipoCanonico,
          grupo: info.grupo || null,
          grado: info.grado || '',
          escudo: info.escudo || '',
          goles: 0
        };
      }
      acumulado[key].goles += goles;
    });

  return Object.values(acumulado)
    .filter(r => r.goles > 0)
    .sort((a, b) => b.goles - a.goles || a.jugador.localeCompare(b.jugador));
}

function leerEliminatorias(hoja, equipos) {
  if (!hoja || hoja.getLastRow() < 2) return [];
  const infoEquipo = {};
  (equipos || []).forEach(e => {
    const key = String(e.nombre).trim().toLowerCase();
    if (key) infoEquipo[key] = {
      nombre: e.nombre, escudo: e.escudo, grado: e.grado, grupo: e.grupo
    };
  });
  const normalizar = (s) => {
    const raw = String(s || '').trim();
    const found = infoEquipo[raw.toLowerCase()];
    return found ? found.nombre : raw;
  };
  const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 10).getValues();
  return datos
    .filter(r => r[0])
    .map((r, i) => {
      const g1 = r[4] === '' || r[4] === null ? null : Number(r[4]);
      const g2 = r[6] === '' || r[6] === null ? null : Number(r[6]);
      const p1 = r[7] === '' || r[7] === null ? null : Number(r[7]);
      const p2 = r[8] === '' || r[8] === null ? null : Number(r[8]);
      const estadoRaw = String(r[9] || '').trim();
      const equipo1 = normalizar(r[3]);
      const equipo2 = normalizar(r[5]);
      const info1 = infoEquipo[equipo1.toLowerCase()] || {};
      const info2 = infoEquipo[equipo2.toLowerCase()] || {};
      const jugado = g1 !== null && g2 !== null && equipo1 && equipo2;
      const estado = estadoRaw || (jugado ? 'Jugado' : 'Pendiente');
      let ganador = null;
      if (jugado && estado === 'Jugado') {
        if (g1 > g2) ganador = equipo1;
        else if (g2 > g1) ganador = equipo2;
        else if (p1 !== null && p2 !== null) {
          if (p1 > p2) ganador = equipo1;
          else if (p2 > p1) ganador = equipo2;
        }
      }
      return {
        id: i + 1,
        fase: String(r[0]).trim(),
        rama: String(r[1] || '').trim() || 'General',
        fecha: r[2] instanceof Date ? r[2].toISOString() : String(r[2] || ''),
        equipo1: equipo1,
        escudo1: info1.escudo || '',
        grado1: info1.grado || '',
        goles1: g1,
        equipo2: equipo2,
        escudo2: info2.escudo || '',
        grado2: info2.grado || '',
        goles2: g2,
        penales1: p1,
        penales2: p2,
        estado: estado,
        ganador: ganador
      };
    });
}

function leerPartidos(hoja, equipos) {
  if (!hoja || hoja.getLastRow() < 2) return [];
  const lookup = {};
  (equipos || []).forEach(e => {
    const key = String(e.nombre).trim().toLowerCase();
    if (key) lookup[key] = { nombre: e.nombre, grupo: e.grupo };
  });
  const normalizar = (s) => {
    const raw = String(s || '').trim();
    const found = lookup[raw.toLowerCase()];
    return found ? found.nombre : raw;
  };
  const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 7).getValues();
  return datos
    .filter(r => r[2] && r[4])
    .map((r, i) => {
      const gl = r[3] === '' || r[3] === null ? null : Number(r[3]);
      const gv = r[5] === '' || r[5] === null ? null : Number(r[5]);
      const estadoRaw = String(r[6] || '').trim();
      const estado = estadoRaw || (gl !== null && gv !== null ? 'Jugado' : 'Pendiente');
      const local = normalizar(r[2]);
      const visitante = normalizar(r[4]);
      const grupoL = (lookup[local.toLowerCase()] || {}).grupo || null;
      const grupoV = (lookup[visitante.toLowerCase()] || {}).grupo || null;
      return {
        id: i + 1,
        jornada: Number(r[0]) || 0,
        fecha: r[1] instanceof Date ? r[1].toISOString() : String(r[1] || ''),
        local: local,
        golesLocal: gl,
        visitante: visitante,
        golesVisitante: gv,
        estado: estado,
        grupo: grupoL && grupoL === grupoV ? grupoL : (grupoL || grupoV || null)
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
    clasifican: 2,
    ordenGrupos: []
  };
  if (!hoja || hoja.getLastRow() < 2) return defaults;
  const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 2).getValues();
  const map = {};
  datos.forEach(r => { if (r[0]) map[String(r[0]).trim()] = r[1]; });
  const ordenRaw = String(map['Orden de grupos'] || '').trim();
  const ordenGrupos = ordenRaw
    ? ordenRaw.split(',').map(s => s.trim()).filter(Boolean)
    : [];
  return {
    nombre: map['Nombre del Campeonato'] || defaults.nombre,
    subtitulo: map['Subtítulo'] || defaults.subtitulo,
    colorPrimario: map['Color Primario'] || defaults.colorPrimario,
    colorSecundario: map['Color Secundario'] || defaults.colorSecundario,
    logo: map['Logo (URL)'] || defaults.logo,
    clasifican: Number(map['Clasifican (primeros N)']) || defaults.clasifican,
    ordenGrupos: ordenGrupos
  };
}

function calcularTablasPorGrupo(equipos, partidos) {
  const porGrupo = {};
  equipos.forEach(e => {
    const g = e.grupo || 'General';
    if (!porGrupo[g]) porGrupo[g] = [];
    porGrupo[g].push(e);
  });

  const result = {};
  Object.keys(porGrupo).forEach(grupo => {
    result[grupo] = calcularTablaGrupo(porGrupo[grupo], partidos);
  });
  return result;
}

function calcularTablaGrupo(equiposGrupo, partidos) {
  const nombres = {};
  equiposGrupo.forEach(e => { nombres[e.nombre] = true; });

  const tabla = {};
  equiposGrupo.forEach(e => {
    tabla[e.nombre] = {
      equipo: e.nombre,
      grado: e.grado,
      grupo: e.grupo,
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
      if (!nombres[p.local] || !nombres[p.visitante]) return;
      const L = tabla[p.local];
      const V = tabla[p.visitante];

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
