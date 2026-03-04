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
function procesarBloqueTexto(contenido) {

  const hoja = SpreadsheetApp
    .getActive()
    .getSheetByName("escaner");

  if (!hoja) return {insertados:0, duplicados:0, invalidos:0};

  const lineas = contenido
    .split(/\r?\n/)
    .map(l => l.trim())
    .filter(l => l !== "");

  if (lineas.length === 0)
    return {insertados:0, duplicados:0, invalidos:0};

  const lastRow = hoja.getLastRow();

  let existentesSet = new Set();

  if (lastRow > 1) {
    const existentes = hoja
      .getRange(2,9,lastRow-1,1)
      .getValues()
      .flat();

    existentesSet = new Set(existentes);
  }

  const nuevos = [];
  let duplicados = 0;
  let invalidos = 0;

  const CLASIFICACION_MAP = {
    "01": "CARGA GENERAL",
    "02": "FARMA",
    "03": "FRAGANCIA",
    "04": "GRADOS DE ALIMENTOS",
    "05": "QUIMICOS",
    "06": "SABORES"
  };

  const ahora = new Date();

  for (let linea of lineas) {

    // Divide por TAB o múltiples espacios
    const partes = linea.split(/\s+/);

    if (partes.length < 2) {
      invalidos++;
      continue;
    }

    const qr = partes[0].trim();
    const zonaManual = partes[1].trim().toUpperCase();

    if (!/^\d{32}$/.test(qr)) {
      invalidos++;
      continue;
    }

    if (existentesSet.has(qr)) {
      duplicados++;
      continue;
    }

    const remesa = qr.substring(0,12);
    const unidad = Number(qr.substring(12,16));
    const totalUnidades = Number(qr.substring(16,20));
    const regionalDestino = qr.substring(20,22);
    const codigoClasificacion = qr.substring(22,24);
    const zonaQR = qr.substring(24,28);
    const pais = qr.substring(28,32);

    const clasificacion =
      CLASIFICACION_MAP[codigoClasificacion] || "SIN CLASIFICAR";

    nuevos.push([
  remesa,            // A
  unidad,            // B
  totalUnidades,     // C
  regionalDestino,   // D
  clasificacion,     // E
  zonaQR,            // F 
  pais,              // G
  ahora,             // H
  qr,                // I
  zonaManual         // J  
]);

    existentesSet.add(qr);
  }

  if (nuevos.length > 0) {
    hoja.getRange(lastRow + 1, 1, nuevos.length, 10)
         .setValues(nuevos);

    actualizarInventarioGeneral();
  }

  return {
    insertados: nuevos.length,
    duplicados: duplicados,
    invalidos: invalidos
  };
}


