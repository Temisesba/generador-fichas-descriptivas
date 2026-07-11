/**
 * BACKEND - Generador de Fichas Descriptivas (Francisco Villa)
 *
 * INSTALACIÓN:
 * 1. Abre tu Google Sheet: https://docs.google.com/spreadsheets/d/1StBJPDMa6e4CuYjoyWH8YALx_XSl68c7t21sdCIxtis/edit
 * 2. Menú Extensiones > Apps Script.
 * 3. Borra todo lo que haya en el editor y pega este archivo completo.
 * 4. Guarda (icono de disco o Ctrl+S).
 * 5. En el desplegable de funciones (junto al botón "Depurar") elige "setup" y presiona "Ejecutar".
 *    La primera vez pedirá autorización: Revisar permisos > (tu cuenta) > Avanzado > Ir a proyecto (no seguro) > Permitir.
 *    Esto crea las pestañas Config, Alumnos, Frases y Fichas, y las llena con los datos iniciales.
 * 6. Menú Implementar > Administrar implementaciones > icono de lápiz (editar) > en "Versión" elige
 *    "Nueva versión" > Implementar. Copia la URL que termina en /exec (debe ser la misma que ya tienes).
 * 7. Pega esa URL en la pestaña "Configuración" de la app HTML si llegara a cambiar.
 *
 * Cada vez que vuelvas a pegar cambios de código aquí, repite el paso 6 (Nueva versión) para que
 * la URL /exec sirva el código actualizado.
 */

const SHEET_CONFIG = 'Config';
const SHEET_ALUMNOS = 'Alumnos';
const SHEET_FRASES = 'Frases';
const SHEET_FICHAS = 'Fichas';

const CAMPOS = [
  { key: 'LENGUAJES', prefix: 'LENGUAJES' },
  { key: 'SABERES Y PENSAMIENTO CIENTÍFICO', prefix: 'SABERES' },
  { key: 'ÉTICA NATURALEZA Y SOCIEDAD', prefix: 'ETICA' },
  { key: 'DE LO HUMANO Y LO COMUNITARIO', prefix: 'HUMANO' }
];

const INDICADORES = [
  { key: 'participacionClase', label: 'Participación en clase' },
  { key: 'entregaTiempo', label: 'Entrega a tiempo de trabajos' },
  { key: 'disposicion', label: 'Muestra disposición para el trabajo' },
  { key: 'concentracion', label: 'Concentración en clase' },
  { key: 'colabora', label: 'Colabora en las actividades' },
  { key: 'externaDudas', label: 'Externa dudas e inquietudes' },
  { key: 'esResponsable', label: 'Es responsable' },
  { key: 'apoyoCasa', label: 'Lo apoyan en casa' },
  { key: 'asistencia', label: 'Asistencia y puntualidad' },
  { key: 'convivencia', label: 'Sana convivencia' }
];

const FICHAS_HEADERS = ['Alumno'].concat(
  CAMPOS.reduce((acc, c) => acc.concat([c.prefix + '_Fortalezas', c.prefix + '_Oportunidades']), []),
  INDICADORES.map(i => 'Ind_' + i.key),
  ['Actualizado']
);

// ============================================
// ENTRADA HTTP
// ============================================

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'getAll';
    if (action === 'getAll') {
      return jsonOut({ ok: true, config: leerConfig(), alumnos: leerAlumnos(), frases: leerFrases(), fichas: leerFichas() });
    }
    return jsonOut({ ok: false, error: 'Acción GET no reconocida: ' + action });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    const payload = body.payload || {};
    let resultado;

    switch (action) {
      case 'guardarConfig':
        resultado = guardarConfig(payload);
        break;
      case 'agregarAlumno':
        resultado = agregarAlumno(payload);
        break;
      case 'eliminarAlumno':
        resultado = eliminarAlumno(payload);
        break;
      case 'editarAlumno':
        resultado = editarAlumno(payload);
        break;
      case 'agregarFrase':
        resultado = agregarFrase(payload);
        break;
      case 'editarFrase':
        resultado = editarFrase(payload);
        break;
      case 'eliminarFrase':
        resultado = eliminarFrase(payload);
        break;
      case 'guardarFicha':
        resultado = guardarFicha(payload);
        break;
      default:
        return jsonOut({ ok: false, error: 'Acción POST no reconocida: ' + action });
    }

    return jsonOut({ ok: true, resultado: resultado, config: leerConfig(), alumnos: leerAlumnos(), frases: leerFrases(), fichas: leerFichas() });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ============================================
// UTILIDADES DE HOJA
// ============================================

function ss() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet(nombre) {
  const hoja = ss().getSheetByName(nombre);
  if (!hoja) throw new Error('No existe la pestaña "' + nombre + '". Corre la función setup() primero.');
  return hoja;
}

function getOrCreateSheet(nombre, headers) {
  let hoja = ss().getSheetByName(nombre);
  if (!hoja) {
    hoja = ss().insertSheet(nombre);
  }
  if (headers && hoja.getLastRow() === 0) {
    hoja.getRange(1, 1, 1, headers.length).setValues([headers]);
    hoja.setFrozenRows(1);
  }
  return hoja;
}

// ============================================
// CONFIG
// ============================================

function leerConfig() {
  const hoja = getSheet(SHEET_CONFIG);
  const filas = hoja.getDataRange().getValues();
  const config = {};
  for (let i = 1; i < filas.length; i++) {
    const clave = filas[i][0];
    const valor = filas[i][1];
    if (clave) config[clave] = valor;
  }
  return config;
}

function guardarConfig(payload) {
  const hoja = getSheet(SHEET_CONFIG);
  const filas = hoja.getDataRange().getValues();
  const indicePorClave = {};
  for (let i = 1; i < filas.length; i++) {
    indicePorClave[filas[i][0]] = i + 1; // fila real (1-indexed)
  }
  Object.keys(payload).forEach(clave => {
    if (indicePorClave[clave]) {
      hoja.getRange(indicePorClave[clave], 2).setValue(payload[clave]);
    } else {
      hoja.appendRow([clave, payload[clave]]);
    }
  });
  return { guardado: true };
}

// ============================================
// ALUMNOS
// ============================================

function leerAlumnos() {
  const hoja = getSheet(SHEET_ALUMNOS);
  const filas = hoja.getDataRange().getValues();
  const alumnos = [];
  for (let i = 1; i < filas.length; i++) {
    const nombre = filas[i][0];
    const activo = filas[i][1];
    if (nombre && activo !== false) alumnos.push(String(nombre).trim());
  }
  return alumnos;
}

function agregarAlumno(payload) {
  const hoja = getSheet(SHEET_ALUMNOS);
  hoja.appendRow([payload.nombre, true]);
  return { agregado: payload.nombre };
}

function eliminarAlumno(payload) {
  const hoja = getSheet(SHEET_ALUMNOS);
  const filas = hoja.getDataRange().getValues();
  for (let i = 1; i < filas.length; i++) {
    if (String(filas[i][0]).trim() === String(payload.nombre).trim()) {
      hoja.deleteRow(i + 1);
      break;
    }
  }
  return { eliminado: payload.nombre };
}

function editarAlumno(payload) {
  const hojaAlumnos = getSheet(SHEET_ALUMNOS);
  const filasA = hojaAlumnos.getDataRange().getValues();
  for (let i = 1; i < filasA.length; i++) {
    if (String(filasA[i][0]).trim() === String(payload.nombreAntiguo).trim()) {
      hojaAlumnos.getRange(i + 1, 1).setValue(payload.nombreNuevo);
      break;
    }
  }

  // Si ya tenía una ficha guardada, se renombra junto con el alumno para no perder su trabajo.
  const hojaFichas = getSheet(SHEET_FICHAS);
  const filasF = hojaFichas.getDataRange().getValues();
  for (let i = 1; i < filasF.length; i++) {
    if (String(filasF[i][0]).trim() === String(payload.nombreAntiguo).trim()) {
      hojaFichas.getRange(i + 1, 1).setValue(payload.nombreNuevo);
      break;
    }
  }

  return { editado: payload.nombreNuevo };
}

// ============================================
// FRASES (catálogo)
// ============================================

function leerFrases() {
  const hoja = getSheet(SHEET_FRASES);
  const filas = hoja.getDataRange().getValues();
  const frases = {};
  for (let i = 1; i < filas.length; i++) {
    const [campo, tipo, frase] = filas[i];
    if (!campo || !tipo || !frase) continue;
    if (!frases[campo]) frases[campo] = {};
    if (!frases[campo][tipo]) frases[campo][tipo] = [];
    frases[campo][tipo].push(frase);
  }
  return frases;
}

function agregarFrase(payload) {
  const hoja = getSheet(SHEET_FRASES);
  hoja.appendRow([payload.campo, payload.tipo, payload.frase]);
  return { agregada: payload.frase };
}

function editarFrase(payload) {
  const hoja = getSheet(SHEET_FRASES);
  const filas = hoja.getDataRange().getValues();
  for (let i = 1; i < filas.length; i++) {
    if (filas[i][0] === payload.campo && filas[i][1] === payload.tipo && filas[i][2] === payload.fraseAntigua) {
      hoja.getRange(i + 1, 3).setValue(payload.fraseNueva);
      break;
    }
  }
  return { editada: true };
}

function eliminarFrase(payload) {
  const hoja = getSheet(SHEET_FRASES);
  const filas = hoja.getDataRange().getValues();
  for (let i = 1; i < filas.length; i++) {
    if (filas[i][0] === payload.campo && filas[i][1] === payload.tipo && filas[i][2] === payload.frase) {
      hoja.deleteRow(i + 1);
      break;
    }
  }
  return { eliminada: true };
}

// ============================================
// FICHAS (una fila por alumno)
// ============================================

function leerFichas() {
  const hoja = getSheet(SHEET_FICHAS);
  const filas = hoja.getDataRange().getValues();
  const headers = filas[0];
  const fichas = {};

  for (let i = 1; i < filas.length; i++) {
    const fila = filas[i];
    const alumno = fila[0];
    if (!alumno) continue;

    const ficha = { campos: {}, indicadores: {} };
    CAMPOS.forEach(c => {
      const colF = headers.indexOf(c.prefix + '_Fortalezas');
      const colO = headers.indexOf(c.prefix + '_Oportunidades');
      ficha.campos[c.key] = {
        fortalezas: splitFrases(fila[colF]),
        oportunidades: splitFrases(fila[colO])
      };
    });
    INDICADORES.forEach(ind => {
      const col = headers.indexOf('Ind_' + ind.key);
      ficha.indicadores[ind.key] = fila[col] || '';
    });

    fichas[alumno] = ficha;
  }
  return fichas;
}

function splitFrases(celda) {
  if (!celda) return [];
  return String(celda).split('\n').map(s => s.trim()).filter(Boolean);
}

function guardarFicha(payload) {
  const hoja = getSheet(SHEET_FICHAS);
  const filas = hoja.getDataRange().getValues();
  const headers = filas[0];
  let filaIndice = -1;

  for (let i = 1; i < filas.length; i++) {
    if (String(filas[i][0]).trim() === String(payload.alumno).trim()) {
      filaIndice = i + 1; // fila real
      break;
    }
  }

  const fila = new Array(headers.length).fill('');
  fila[0] = payload.alumno;

  CAMPOS.forEach(c => {
    const datosCampo = (payload.campos && payload.campos[c.key]) || { fortalezas: [], oportunidades: [] };
    const colF = headers.indexOf(c.prefix + '_Fortalezas');
    const colO = headers.indexOf(c.prefix + '_Oportunidades');
    fila[colF] = (datosCampo.fortalezas || []).join('\n');
    fila[colO] = (datosCampo.oportunidades || []).join('\n');
  });

  INDICADORES.forEach(ind => {
    const col = headers.indexOf('Ind_' + ind.key);
    fila[col] = (payload.indicadores && payload.indicadores[ind.key]) || '';
  });

  fila[headers.indexOf('Actualizado')] = new Date();

  if (filaIndice === -1) {
    hoja.appendRow(fila);
  } else {
    hoja.getRange(filaIndice, 1, 1, fila.length).setValues([fila]);
  }
  return { guardada: payload.alumno };
}

// ============================================
// SETUP (correr UNA sola vez desde el editor)
// ============================================

function setup() {
  const hojaConfig = getOrCreateSheet(SHEET_CONFIG, ['Campo', 'Valor']);
  if (hojaConfig.getLastRow() <= 1) {
    // Datos de ejemplo genéricos: edita estos valores desde la pestaña Config del Sheet
    // (o desde la app, pestaña Configuración) una vez que corras setup(). No se vuelven a
    // sobrescribir en corridas futuras.
    const datosConfig = [
      ['nombreEscuela', 'NOMBRE DE TU ESCUELA'],
      ['claveEscuela', 'C.C.T.'],
      ['zonaEscolar', 'ZONA ESCOLAR'],
      ['direccion', 'DIRECCIÓN'],
      ['cicloEscolar', '2025 - 2026'],
      ['gradoGrupo', 'GRADO Y GRUPO'],
      ['director', ''],
      ['maestro', '']
    ];
    hojaConfig.getRange(2, 1, datosConfig.length, 2).setValues(datosConfig);
  }

  const hojaAlumnos = getOrCreateSheet(SHEET_ALUMNOS, ['Nombre', 'Activo']);
  if (hojaAlumnos.getLastRow() <= 1) {
    // Lista de ejemplo: bórrala o edítala desde la pestaña Alumnos del Sheet (o desde la app).
    const nombres = [
      'ALUMNO DE EJEMPLO 1', 'ALUMNO DE EJEMPLO 2', 'ALUMNO DE EJEMPLO 3'
    ];
    hojaAlumnos.getRange(2, 1, nombres.length, 2).setValues(nombres.map(n => [n, true]));
  }

  const hojaFrases = getOrCreateSheet(SHEET_FRASES, ['CampoFormativo', 'Tipo', 'Frase']);
  if (hojaFrases.getLastRow() <= 1) {
    const catalogo = {
      'LENGUAJES': {
        'Logro destacado': [
          'Demuestra interés y gusto por la lectura de diversos textos adecuados a su edad.',
          'Mantiene una participación en actividades de comunicación oral y escrita.',
          'Desarrolla habilidades de comprensión lectora que le permitan interpretar información con facilidad.',
          'Consolida la producción de textos breves con claridad y coherencia.',
          'Expresa ideas, emociones y experiencias de manera organizada y pertinente.'
        ],
        'Proceso': [
          'Requiere fortalecer la compresión de textos breves mediante practica constante.',
          'Desarrolla progresivamente habilidades relacionadas con la lectura y la escritura.',
          'Consolida gradualmente la identificación de información explicita en los textos.',
          'Requiere apoyo ocasional para organizar ideas de manera escrita.',
          'Fortalece el reconocimiento y el uso de palabras de uso frecuente.'
        ],
        'Acompañamiento': [
          'Requiere apoyo constante para fortalecer el proceso de adquisición de la lectura y escritura.',
          'Desarrolla habilidades comunicativas básicas con acompañamiento permanente.',
          'Fortalece gradualmente el reconocimiento de las letras, silabas y palabras.',
          'Requiere orientación frecuente para comprender las instrucciones orales y escritas.',
          'Participa en actividades de lectura y escrita con apoyo docente.'
        ]
      },
      'SABERES Y PENSAMIENTO CIENTÍFICO': {
        'Logro destacado': [
          'Demuestra interés por explorar, observar y comprender fenómenos de su entorno.',
          'Utiliza estrategias adecuadas para responder problemas matemáticos acordes a su nivel.',
          'Aplica conocimientos matemáticos en situaciones de la vida cotidiana.',
          'Analiza información y propone soluciones utilizado el razonamiento lógico.',
          'Identifica patrones, relaciones y regularidades en diversas situaciones.'
        ],
        'Proceso': [
          'Fortalece progresivamente la comprensión de conceptos matemáticos básicos.',
          'Requiere continuar practicando estrategias para resolver problemas de manera autónoma.',
          'Consolida gradualmente el reconocimiento y uso de números en diferentes situaciones.',
          'Desarrolla habilidades de observación y análisis mediante actividades guiadas.',
          'Participa en actividades de exploración mostrando avances en la comprensión de su entorno.'
        ],
        'Acompañamiento': [
          'Requiere apoyo constante para comprender y aplicar conceptos matemáticos básicos.',
          'Necesita acompañamiento para resolver problemas utilizando estrategias adecuadas.',
          'Requiere fortalecer habilidades de observación, comparación y análisis de su entorno.',
          'Participa en actividades de exploración y resolución de problemas con apoyo docente.',
          'Avanza de acuerdo con su ritmo de aprendizaje, beneficiándose de estrategias diferenciadas.'
        ]
      },
      'ÉTICA NATURALEZA Y SOCIEDAD': {
        'Logro destacado': [
          'Demuestra respeto hacia las normas de convivencia y las aplica de manera responsable.',
          'Participa activamente en acciones que favorecen el bienestar común dentro y fuera del aula.',
          'Valora la diversidad de personas, costumbres y formas de vida presente en su entorno.',
          'Muestra interés por conocer y cuidar los recursos naturales de su comunidad.',
          'Identifica la importancia de sus derechos y responsabilidades como integrante de diversos grupos.'
        ],
        'Proceso': [
          'Fortalece progresivamente el respeto por las normas y acuerdos de convivencia.',
          'Requiere apoyo ocasional para colaborar de manera constante en actividades grupales.',
          'Consolida gradualmente actitudes de respeto hacia las diferencias individuales.',
          'Desarrolla mayor conciencia sobre la importancia del cuidado del entorno.',
          'Participa en actividades de convivencia mostrando avances en la resolución pacífica de conflictos.'
        ],
        'Acompañamiento': [
          'Requiere apoyo constante para respetar normas y acuerdos de convivencia.',
          'Necesita acompañamiento para desarrollar actitudes de respeto y colaboración con los demás.',
          'Requiere fortalecer la identificación de acciones relacionadas con el cuidado del entorno.',
          'Participa en actividades de convivencia y cuidado comunitario cuando recibe orientación adecuada.',
          'Avanza de acuerdo con su ritmo de aprendizaje, beneficiándose del acompañamiento constante.'
        ]
      },
      'DE LO HUMANO Y LO COMUNITARIO': {
        'Logro destacado': [
          'Demuestra seguridad al expresar sus emociones, ideas y necesidades de manera respetuosa.',
          'Participa activamente en actividades que favorecen el bienestar personal y colectivo.',
          'Mantiene relaciones positivas con sus compañeros basadas en el respeto y la colaboración.',
          'Practica hábitos de autocuidado que favorecen su salud y bienestar integral.',
          'Participa con entusiasmo en actividades artísticas, físicas, recreativas y culturales.'
        ],
        'Proceso': [
          'Fortalece progresivamente la identificación y expresión adecuada a sus emociones.',
          'Participa en actividades grupales desarrollando gradualmente habilidades de colaboración.',
          'Consolida hábitos de higiene, alimentación y cuidado personal acordes a su edad.',
          'Desarrolla mayor autonomía en el cumplimiento de responsabilidades escolares.',
          'Requiere fortalecer estrategias para resolver desacuerdos de manera pacífica.'
        ],
        'Acompañamiento': [
          'Requiere apoyo constante para identificar y expresar adecuadamente sus emociones.',
          'Necesita acompañamiento para fortalecer su participación en actividades grupales.',
          'Requiere refuerzo en hábitos de higiene, alimentación y cuidado personal.',
          'Participa en actividades cuando recibe orientaciones y apoyo oportuno.',
          'Avanza de acuerdo con su ritmo de aprendizaje cuando cuenta con acompañamiento adecuado.'
        ]
      }
    };
    const filasFrases = [];
    Object.keys(catalogo).forEach(campo => {
      Object.keys(catalogo[campo]).forEach(tipo => {
        catalogo[campo][tipo].forEach(frase => filasFrases.push([campo, tipo, frase]));
      });
    });
    hojaFrases.getRange(2, 1, filasFrases.length, 3).setValues(filasFrases);
  }

  getOrCreateSheet(SHEET_FICHAS, FICHAS_HEADERS);

  SpreadsheetApp.getUi().alert('Listo. Pestañas creadas y datos iniciales cargados.');
}
