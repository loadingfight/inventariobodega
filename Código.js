/*************************************************
 * CONFIGURACIÓN
 *************************************************/
const HOJA_ESCANER = "escaner";
const HOJA_INVENTARIO = "Inventario";
const HOJA_DOCUMENTOS = "documentos";


/*************************************************
 * CARGAR HTML
 *************************************************/
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Inventario')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/*************************************************
 * INCLUDE
 *************************************************/
function include(nombre) {
  return HtmlService.createHtmlOutputFromFile(nombre).getContent();
}


/*************************************************
 * GUARDAR DESDE QR
 *************************************************/
function guardarQR(qr, ubicacion) {

  if (!qr || qr.length !== 32) {
    return { error: "QR inválido (32 caracteres requeridos)" };
  }

  const hoja = SpreadsheetApp
    .getActive()
    .getSheetByName("escaner");

  const remesa = qr.substring(0,12);
  const unidad = parseInt(qr.substring(12,16));
  const total = parseInt(qr.substring(16,20));

  const clasificacionCodigo = qr.substring(22,24);

  const CLASIFICACION_MAP = {
    "01": "CARGA GENERAL",
    "02": "FARMA",
    "03": "FRAGANCIA",
    "04": "GRADOS DE ALIMENTOS",
    "05": "QUIMICOS",
    "06": "SABORES"
  };

  const clasificacion =
    CLASIFICACION_MAP[clasificacionCodigo] || "SIN CLASIFICAR";

  const lastRow = hoja.getLastRow();

  // VALIDACIÓN DUPLICADO (QR en columna 9)
  if (lastRow > 1) {
    const qrExistentes =
      hoja.getRange(2,9,lastRow-1,1)
           .getValues()
           .flat();

    if (qrExistentes.includes(qr)) {
      return { error: "⚠ QR ya registrado" };
    }
  }

  hoja.appendRow([
   remesa,            // 1 Remesa
    unidad,            // 2 Unidad
    total,             // 3 Total Unidades
    qr.substring(20,22), // 4 Regional Destino
    clasificacion,     // 5 Clasificacion
    qr.substring(24,28), // 6 Zona
    qr.substring(28,32), // 7 Pais
    new Date(),        // 8 Fecha
    qr,                // 9 QR_Completo
    ubicacion          // 10 Ubicacion
  ]);

  
// ==============================
// CALCULAR AVANCE CORRECTAMENTE
// ==============================

const lastRowActual = hoja.getLastRow();

if (lastRowActual < 2) {
  return {
    success: true,
    porcentaje: 0,
    escaneados: 0,
    total,
    unidadesEscaneadas: [],
    faltantes: Array.from({length: total}, (_,i)=>i+1)
  };
}

const data = hoja
  .getRange(2,1,lastRowActual-1,3)
  .getValues();

const unidadesEscaneadas = [];

for (let fila of data) {

  if (String(fila[0]).trim() === String(remesa).trim()) {
    unidadesEscaneadas.push(Number(fila[1]));
  }
}

unidadesEscaneadas.sort((a,b)=>a-b);

const escaneados = unidadesEscaneadas.length;

const porcentaje =
  total > 0
    ? Math.round((escaneados / total) * 100)
    : 0;

const setUnidades = new Set(unidadesEscaneadas);

const faltantes = [];

for (let i = 1; i <= total; i++) {
  if (!setUnidades.has(i)) {
    faltantes.push(i);
  }
}

return {
  success: true,
  porcentaje,
  escaneados,
  total,
  unidadesEscaneadas,
  faltantes
};

}
/*************************************************
 * CSV QR
 *************************************************/
function procesarCSV(contenido, ubicacion) {

  const ss = SpreadsheetApp.getActive();
  const hoja = ss.getSheetByName(HOJA_ESCANER);

  const lineas = contenido.split(/\r?\n/).filter(l => l.trim() !== "");
  if (lineas.length === 0) return { insertados: 0, duplicados: 0 };

  const lastRow = hoja.getLastRow();

  let existentesSet = new Set();
  if (lastRow > 1) {
    existentesSet = new Set(
      hoja.getRange(2,9,lastRow-1,1)
      .getValues()
      .flat()
      .map(String)
    );
  }

  const nuevos = [];
  let duplicados = 0;

  for (let linea of lineas) {

    const qr = linea.trim();
    if (!/^\d{32}$/.test(qr)) continue;

    if (existentesSet.has(qr)) {
      duplicados++;
      continue;
    }

    const remesa = String(qr.substring(0,12)).trim();
    const unidad = parseInt(qr.substring(12,16));
    const totalUnidades = parseInt(qr.substring(16,20));
    const regionalDestino = qr.substring(20,22);
    const codigoClasificacion = qr.substring(22,24).trim();

        const CLASIFICACION_MAP = {
              "01": "CARGA GENERAL",
              "02": "FARMA",
              "03": "FRAGANCIA",
              "04": "GRADOS DE ALIMENTOS",
              "05": "QUIMICOS",
              "06": "SABORES"
            };

            const clasificacionTexto = CLASIFICACION_MAP[clasificacion] || "SIN CLASIFICAR";
    
    const zona = qr.substring(24,28);
    const pais = qr.substring(28,32);

    nuevos.push([
      remesa,
      unidad,
      totalUnidades,
      regionalDestino,
      clasificacion,
      zona,
      pais,
      new Date(),
      qr,
      ubicacion
    ]);

    existentesSet.add(qr);
  }

  if (nuevos.length > 0) {
    hoja.getRange(lastRow + 1, 1, nuevos.length, 10).setValues(nuevos);
  }

  actualizarInventarioGeneral();

  return {
    insertados: nuevos.length,
    duplicados: duplicados
  };
}

/*************************************************
 * GUARDAR DOCUMENTO + ACTUALIZAR INVENTARIO
 *************************************************/
function guardarDocumento(remesa, zona) {

  if (!remesa || remesa.length !== 12) {
    return { error: "La remesa debe tener 12 caracteres" };
  }

  const ss = SpreadsheetApp.getActive();
  const hojaDoc = ss.getSheetByName(HOJA_DOCUMENTOS);

  const lastRow = hojaDoc.getLastRow();

  if (lastRow > 1) {

    const existentes = hojaDoc
      .getRange(2,1,lastRow-1,1)
      .getValues()
      .flat()
      .map(String);

    if (existentes.includes(remesa)) {
      return { error: "Documento ya registrado" };
    }
  }

  hojaDoc
    .getRange(lastRow + 1, 1, 1, 3)
    .setValues([[remesa, zona, new Date()]]);

  return { success: true };
}

/*************************************************
 * PROCESAR CSV DOCUMENTOS + ACTUALIZAR INVENTARIO
 *************************************************/
function procesarCSVDocumentos(contenido, zona) {

  const ss = SpreadsheetApp.getActive();
  const hojaDoc = ss.getSheetByName(HOJA_DOCUMENTOS);
  const hojaInv = ss.getSheetByName(HOJA_INVENTARIO);

  const lineas = contenido
    .split(/\r?\n/)
    .filter(l => l.trim() !== "");

  const lastRowDoc = hojaDoc.getLastRow();

  let existentesSet = new Set();

  if (lastRowDoc > 1) {
    existentesSet = new Set(
      hojaDoc
        .getRange(2,1,lastRowDoc-1,1)
        .getValues()
        .flat()
        .map(String)
    );
  }

  const nuevos = [];
  let duplicados = 0;

  for (let linea of lineas) {

    const remesa = linea.trim();

    if (remesa.length !== 12) continue;

    if (existentesSet.has(remesa)) {
      duplicados++;
      continue;
    }

    nuevos.push([
      remesa,
      zona,
      new Date()   // 🔥 AUTOMÁTICA
    ]);

    existentesSet.add(remesa);
  }

  if (nuevos.length > 0) {

    hojaDoc
      .getRange(lastRowDoc + 1, 1, nuevos.length, 3)
      .setValues(nuevos);

    // ACTUALIZAR INVENTARIO
    const lastRowInv = hojaInv.getLastRow();

    if (lastRowInv > 1) {

      const dataInv = hojaInv
        .getRange(2,1,lastRowInv-1,7)
        .getValues();

      const mapaInv = {};

      dataInv.forEach((fila, i) => {
        mapaInv[String(fila[0]).trim()] = i;
      });

      nuevos.forEach(fila => {

        const remesa = fila[0];

        if (mapaInv.hasOwnProperty(remesa)) {

          hojaInv
            .getRange(mapaInv[remesa] + 2, 7)
            .setValue("SI");
        }
      });
    }
  }

  return {
    insertados: nuevos.length,
    duplicados: duplicados
  };
}

/*************************************************
 * ACTUALIZACIÓN GENERAL INVENTARIO (OPTIMIZADA)
 *************************************************/
function actualizarInventarioGeneral() {

  const ss = SpreadsheetApp.getActive();
  const hojaEscaner = ss.getSheetByName(HOJA_ESCANER);
  const hojaInventario = ss.getSheetByName(HOJA_INVENTARIO);

  const lastRowEscaner = hojaEscaner.getLastRow();
  if (lastRowEscaner < 2) return;

  // ==============================
  // RESUMEN DESDE ESCANER
  // ==============================
  const datosEscaner = hojaEscaner
    .getRange(2,1,lastRowEscaner-1,3)
    .getValues();

  const resumen = {};

  datosEscaner.forEach(fila => {

    const remesa = String(fila[0]).trim();
    const total = fila[2];

    if (!resumen[remesa]) {
      resumen[remesa] = { total: total, escaneadas: 0 };
    }

    resumen[remesa].escaneadas++;
  });

  // ==============================
  // LEER INVENTARIO
  // ==============================
  const lastRowInv = hojaInventario.getLastRow();
  let inventarioData = [];

  if (lastRowInv > 1) {
    inventarioData = hojaInventario
      .getRange(2,1,lastRowInv-1,7)
      .getValues();
  }

  const mapaInv = {};
  inventarioData.forEach((fila,i)=>{
    mapaInv[String(fila[0]).trim()] = i;
  });

  // ==============================
  // ACTUALIZAR / INSERTAR
  // ==============================
  for (let remesa in resumen) {

    const total = resumen[remesa].total;
    const escaneadas = resumen[remesa].escaneadas;
    const faltantes = total - escaneadas;
    const porcentaje = total > 0 ? escaneadas / total : 0;
    const documento = tieneDocumento(remesa) ? "SI" : "NO";

    if (mapaInv.hasOwnProperty(remesa)) {

      const i = mapaInv[remesa];
      inventarioData[i][1] = total;
      inventarioData[i][2] = escaneadas;
      inventarioData[i][3] = faltantes;
      inventarioData[i][4] = porcentaje;
      inventarioData[i][6] = documento;

    } else {

      inventarioData.push([
        remesa,
        total,
        escaneadas,
        faltantes,
        porcentaje,
        "",
        documento
      ]);

    }
  }

  hojaInventario
    .getRange(2,1,inventarioData.length,7)
    .setValues(inventarioData);
}


/*************************************************
 * VALIDAR DOCUMENTOS
 *************************************************/
function tieneDocumento(remesa) {

  const ss = SpreadsheetApp.getActive();
  const hojaDoc = ss.getSheetByName(HOJA_DOCUMENTOS);

  if (!hojaDoc || hojaDoc.getLastRow() < 2) return false;

  const datos = hojaDoc
    .getRange(2,1,hojaDoc.getLastRow()-1,1)
    .getValues()
    .flat()
    .map(x => String(x).trim());

  return datos.includes(String(remesa).trim());
}


/*************************************************
 * Obtener Resumen
 *************************************************/

function obtenerResumenRemesas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("Inventario");

  if (!hoja) {
    throw new Error("No se encontró la hoja Inventario");
  }

  const datos = hoja.getDataRange().getValues();
  if (datos.length <= 1) {
    return {
      totalRemesas: 0,
      totalEscaneadas: 0,
      conFaltantes: 0,
      completas: 0
    };
  }

  let totalRemesas = 0;
  let totalEscaneadas = 0;
  let conFaltantes = 0;
  let completas = 0;

  for (let i = 1; i < datos.length; i++) {
    const total = Number(datos[i][1]) || 0;       // Columna Total
    const escaneadas = Number(datos[i][2]) || 0;  // Escaneadas
    const faltantes = Number(datos[i][3]) || 0;   // Faltantes

    totalRemesas++;
    totalEscaneadas += escaneadas;

    if (faltantes > 0) {
      conFaltantes++;
    }

    if (faltantes === 0 && escaneadas === total && total > 0) {
      completas++;
    }
  }

  return {
    totalRemesas,
    totalEscaneadas,
    conFaltantes,
    completas
  };
}