/*************************************************
 * GUARDAR DOCUMENTO INDIVIDUAL
 *************************************************/

function guardarDocumentosMasivo(remesas, zona) {

  const ss = SpreadsheetApp.getActive();
  const hoja = ss.getSheetByName("documentos");

  const ultimaFila = hoja.getLastRow();

  // 🔹 Si ya hay datos, leerlos UNA sola vez
  let existentes = new Set();

  if (ultimaFila > 1) {
    const datos = hoja
      .getRange(2, 1, ultimaFila - 1, 1)
      .getValues()
      .flat();

    existentes = new Set(datos);
  }

  const nuevos = [];
  let duplicados = 0;

  // 🔹 Validación súper rápida con Set()
  remesas.forEach(r => {
    if (existentes.has(r)) {
      duplicados++;
    } else {
      nuevos.push([r, zona, new Date()]);
      existentes.add(r); // importante para evitar duplicados dentro del mismo envío
    }
  });

  // 🔹 Insertar TODO en un solo bloque
  if (nuevos.length > 0) {
    hoja
      .getRange(ultimaFila + 1, 1, nuevos.length, nuevos[0].length)
      .setValues(nuevos);
  }

  return {
    insertados: nuevos.length,
    duplicados: duplicados
  };
}


/*************************************************
 * OBTENER ALERTAS (< 30%)
 *************************************************/
function obtenerAlertasInventario() {

  const hoja = SpreadsheetApp
    .getActive()
    .getSheetByName("inventario");

  if (!hoja || hoja.getLastRow() < 2) {
    return [];
  }

  const datos = hoja
    .getRange(1,1,hoja.getLastRow(),7) // A → G
    .getValues();

  const encabezados = datos[0];
  const resultado = [encabezados];

  for (let i = 1; i < datos.length; i++) {
    const porcentaje = Number(datos[i][4]) || 0; // Col E

    if (porcentaje < 0.3) {
      resultado.push(datos[i]);
    }
  }

  return resultado;
}


/*************************************************
 * INVENTARIO
 *************************************************/


function obtenerInventarioCompleto() {

  const hoja = SpreadsheetApp
    .getActive()
    .getSheetByName("inventario");

  if (!hoja || hoja.getLastRow() < 2) {
    return [];
  }

  return hoja
    .getRange(1,1,hoja.getLastRow(),7) // A → G
    .getValues();
}