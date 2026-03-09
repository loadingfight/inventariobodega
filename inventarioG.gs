/*************************************************
 * OBTENER TODO EL INVENTARIO (15 COLUMNAS)
 *************************************************/
function obtenerInventarioCompleto() {
  const ss = SpreadsheetApp.getActive();
  const hoja = ss.getSheetByName("inventario");
  const valores = hoja.getRange(1, 1, hoja.getLastRow(), 15).getDisplayValues(); // <--- getDisplayValues es MUCHO más rápido para leer que getValues
  return valores;
}

/*************************************************
 * OBTENER SOLO ALERTAS
 *************************************************/
function obtenerAlertasInventario() {
  const datos = obtenerInventarioCompleto();
  if (datos.length <= 1) return [];

  const encabezados = datos[0];
  // Filtramos: si hay faltantes (Columna D, índice 3) mayor a 0
  const filasAlertas = datos.slice(1).filter(fila => {
    return Number(fila[3]) > 0;
  });

  return [encabezados, ...filasAlertas];
}

/*************************************************
 * ACTUALIZACIÓN GENERAL INVENTARIO
 *************************************************/
function actualizarInventarioGeneral() {
  const ss = SpreadsheetApp.getActive();
  const hojaEscaner = ss.getSheetByName(HOJA_ESCANER);
  const hojaInventario = ss.getSheetByName(HOJA_INVENTARIO);

  if (!hojaEscaner || hojaEscaner.getLastRow() < 2) return;

  const datos = hojaEscaner.getRange(2, 1, hojaEscaner.getLastRow() - 1, 10).getValues();
  const resumen = {};

  // Agrupar por Remesa
  datos.forEach(fila => {
    const remesa = String(fila[0]).trim();
    const total = Number(fila[2]) || 0;
    const fecha = fila[7];
    const ubic = fila[9];

    if (!resumen[remesa]) {
      resumen[remesa] = { total: total, escaneadas: 0, ultimaFecha: fecha, ubicacion: ubic };
    }
    resumen[remesa].escaneadas++;
    if (fecha > resumen[remesa].ultimaFecha) resumen[remesa].ultimaFecha = fecha;
  });

  // Mapa de datos API existentes para no perderlos
  const mapaAPI = {};
  const ultimaFilaInv = hojaInventario.getLastRow();
  if (ultimaFilaInv > 1) {
    const ids = hojaInventario.getRange(2, 1, ultimaFilaInv - 1, 1).getValues().flat();
    const datosAPI = hojaInventario.getRange(2, 9, ultimaFilaInv - 1, 7).getValues();
    ids.forEach((id, i) => { mapaAPI[String(id).trim()] = datosAPI[i]; });
  }

  const documentosSet = obtenerMapaDocumentos();
  const salida = [];

  for (let remesa in resumen) {
    const total = resumen[remesa].total;
    const esc = resumen[remesa].escaneadas;
    const api = mapaAPI[remesa] || ["", "", "", "", "", "", ""];
    
    salida.push([
      remesa, total, esc, total - esc, total > 0 ? esc / total : 0,
      resumen[remesa].ubicacion, documentosSet.has(remesa) ? "SI" : "NO",
      resumen[remesa].ultimaFecha, ...api
    ]);
  }

  salida.sort((a, b) => a[0].localeCompare(b[0]));

  const encabezados = [["Remesa","Total","Escaneadas","Faltantes","%","Ubicación","Documento","FechaLec","Cliente","Destino","Estado","Entregada","FechaCreación","Anulada","ÚltimoEvento"]];

  hojaInventario.clearContents();
  hojaInventario.getRange(1, 1, 1, encabezados[0].length).setValues(encabezados);
  if (salida.length > 0) {
    hojaInventario.getRange(2, 1, salida.length, salida[0].length).setValues(salida);
  }
}

function obtenerMapaDocumentos() {
  const ss = SpreadsheetApp.getActive();
  const hojaDoc = ss.getSheetByName(HOJA_DOCUMENTOS);
  if (!hojaDoc || hojaDoc.getLastRow() < 2) return new Set();
  return new Set(hojaDoc.getRange(2, 1, hojaDoc.getLastRow() - 1, 1).getValues().flat().map(x => String(x).trim()));
}