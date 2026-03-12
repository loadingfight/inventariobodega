/*************************************************
 * CONFIGURACIÓN
*************************************************/
const HOJA_ESCANER = "escaner";
const HOJA_INVENTARIO = "Inventario";
const HOJA_DOCUMENTOS = "documentos";
const HOJA_ENTREGADO = "ENTREGADO";
const HOJA_REPORTES = "Reportes";


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
 * DASHBOARD
 *************************************************/

function obtenerDashboardFiltrado(fechaInicio, fechaFin, ubicacion) {

  const hoja = SpreadsheetApp
    .getActive()
    .getSheetByName("inventario");

  if (!hoja) {
    return { totalRemesas:0, completas:0, incompletas:0, totalCajas:0 };
  }

  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) {
    return { totalRemesas:0, completas:0, incompletas:0, totalCajas:0 };
  }

  const datos = hoja.getRange(2,1,ultimaFila-1,8).getValues();

  const inicio = fechaInicio ? new Date(fechaInicio + "T00:00:00") : null;
  const fin = fechaFin ? new Date(fechaFin + "T23:59:59") : null;

  const filtroUbicacion = ubicacion ? ubicacion.trim().toUpperCase() : null;

  let totalRemesas = 0;
  let completas = 0;
  let incompletas = 0;
  let totalCajas = 0;

  datos.forEach(fila => {

    const remesa = fila[0];                // A
    const escaneadas = Number(fila[2]) || 0; // C
    const faltantes = Number(fila[3]) || 0;  // D
    const ubic = (fila[5] || "").toString().trim().toUpperCase(); // F
    const fecha = fila[7];                 // H

    if (!remesa) return;
    if (!(fecha instanceof Date)) return;

    // 🔹 Filtro por ubicación (columna F)
    if (filtroUbicacion && ubic !== filtroUbicacion) return;

    // 🔹 Filtro por fecha
    if (inicio && fecha < inicio) return;
    if (fin && fecha > fin) return;

    totalRemesas++;

    if (faltantes === 0) {
      completas++;
    } else {
      incompletas++;
    }

    totalCajas += escaneadas;
  });

  return {
    totalRemesas,
    completas,
    incompletas,
    totalCajas
  };
}

