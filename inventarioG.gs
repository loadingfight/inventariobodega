/*************************************************
 * ACTUALIZACIÓN GENERAL INVENTARIO 
 *************************************************/
function actualizarInventarioGeneral() {

  const ss = SpreadsheetApp.getActive();
  const hojaEscaner = ss.getSheetByName("escaner");
  const hojaInventario = ss.getSheetByName("inventario");

  if (!hojaEscaner || hojaEscaner.getLastRow() < 2) return;

  const datos = hojaEscaner
    .getRange(2,1,hojaEscaner.getLastRow()-1,10)
    .getValues();

  const resumen = {};

  datos.forEach(fila => {

    const remesa = String(fila[0]).trim(); // A
    const total = Number(fila[2]) || 0;    // C
    const fecha = fila[7];                 // H
    const ubic = fila[9];                  // J

    if (!resumen[remesa]) {
      resumen[remesa] = {
        total: total,
        escaneadas: 0,
        ultimaFecha: fecha,
        ubicacion: ubic
      };
    }

    resumen[remesa].escaneadas++;

    // Tomar la fecha más reciente
    if (fecha > resumen[remesa].ultimaFecha) {
      resumen[remesa].ultimaFecha = fecha;
    }
  });

  // ==============================
  // RECONSTRUIR INVENTARIO COMPLETO
  // ==============================

  const salida = [];
  const documentosSet = obtenerMapaDocumentos();

  for (let remesa in resumen) {

    const total = resumen[remesa].total;
    const escaneadas = resumen[remesa].escaneadas;
    const faltantes = total - escaneadas;
    const porcentaje = total > 0 ? escaneadas / total : 0;
    const documento = documentosSet.has(remesa) ? "SI" : "NO";

    salida.push([
      remesa,
      total,
      escaneadas,
      faltantes,
      porcentaje,
      resumen[remesa].ubicacion,
      documento,
      resumen[remesa].ultimaFecha
    ]);
  }

  // Limpiar hoja
  hojaInventario.clearContents();

  // Encabezados
  hojaInventario.getRange(1,1,1,8).setValues([[
    "Remesa",
    "Total",
    "Escaneadas",
    "Faltantes",
    "%",
    "Ubicación",
    "Documento",
    "FechaLec"
  ]]);

  if (salida.length > 0) {
    hojaInventario
      .getRange(2,1,salida.length,8)
      .setValues(salida);
  }
}


/*************************************************
 * VALIDAR DOCUMENTOS
 *************************************************/

function obtenerMapaDocumentos() {

  const hojaDoc = SpreadsheetApp
    .getActive()
    .getSheetByName(HOJA_DOCUMENTOS);

  if (!hojaDoc || hojaDoc.getLastRow() < 2) {
    return new Set();
  }

  const datos = hojaDoc
    .getRange(2,1,hojaDoc.getLastRow()-1,1)
    .getValues()
    .flat()
    .map(x => String(x).trim());

  return new Set(datos);
}

