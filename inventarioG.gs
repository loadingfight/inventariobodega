/*************************************************
 * FUNCIONES GLOBLALES
 *************************************************/
function actualizarInventarioCompleto(){

  actualizarInventarioGeneral();   // actualiza datos desde API

  generarAlertasInventario();      // genera alertas columna P

  moverRemesasEntregadas();        // mueve remesas entregadas

}


/*************************************************
 * ALERTAS INVENTARIO
 *************************************************/

function generarAlertasInventario(){

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(HOJA_INVENTARIO);

  const datos = hoja.getDataRange().getValues();

  if(datos.length <= 1) return;

  const hoy = new Date();
  const alertas = [];

  for(let i=1;i<datos.length;i++){

    const escaneadas = Number(datos[i][2]) || 0; // C
    const total = Number(datos[i][3]) || 0; // D
    const documento = String(datos[i][6]).toUpperCase().trim(); // G
    const estadoEntrega = String(datos[i][14]).toUpperCase().trim(); // O

    let fechaTexto = datos[i][12]; // M
    let fechaCreacion = null;

    /***************************
    CONVERTIR TEXTO A FECHA
    ***************************/
    if(typeof fechaTexto === "string" && fechaTexto !== ""){

      fechaTexto = fechaTexto.split(".")[0]; // quitar milisegundos
      fechaTexto = fechaTexto.replace(" ", "T"); // formato ISO

      fechaCreacion = new Date(fechaTexto);

    }

    let alerta = [];

    /**************
    FALTAN UNIDADES
    **************/
    if(escaneadas < total){
      alerta.push("Faltan unidades");
    }

    /**************
    FALTA DOCUMENTOS
    **************/
    if(documento === "NO"){
      alerta.push("Falta documentos");
    }

    /**************
    MUCHOS DIAS
    **************/
    if(fechaCreacion && estadoEntrega !== "ENTREGADO OK"){

      const diffDias = (hoy - fechaCreacion) / (1000*60*60*24);

      if(diffDias > 8){
        alerta.push("Muchos dias");
      }

    }

    alertas.push([alerta.join(" | ")]);
  }

  hoja.getRange(2,16,alertas.length,1).setValues(alertas);

}


/*************************************************
 * OBTENER TODO EL INVENTARIO (15 COLUMNAS)
 *************************************************/
function obtenerInventarioCompleto() {

  const hoja = SpreadsheetApp
    .getActive()
    .getSheetByName("inventario");

  return hoja.getDataRange().getDisplayValues();

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



/*************************************************
  Remesas a Entregadas
*************************************************/

  function moverRemesasEntregadas(){

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const hojaInventario = ss.getSheetByName(HOJA_INVENTARIO);
  let hojaEntregado = ss.getSheetByName(HOJA_ENTREGADO);

  if (!hojaInventario) {
    throw new Error("No existe la hoja: " + HOJA_INVENTARIO);
  }

  if (!hojaEntregado) {
    hojaEntregado = ss.insertSheet(HOJA_ENTREGADO);
  }

  const datos = hojaInventario.getDataRange().getValues();

  if (datos.length <= 1) return;

  const encabezado = datos[0];
  const filas = datos.slice(1);

  const mover = [];
  const mantener = [];

  filas.forEach(fila => {

    const colG = String(fila[6]).trim().toUpperCase();   // Documento
    const colK = String(fila[10]).trim().toUpperCase();  // Estado
    const colO = String(fila[14]).trim().toUpperCase();  // Evento

    if (
      colG === "SI" &&
      colK === "CUMPLIDA" &&
      colO === "ENTREGADO OK"
    ){
      mover.push(fila);
    } else {
      mantener.push(fila);
    }

  });

  if (mover.length === 0){
    Logger.log("No hay remesas para mover.");
    return;
  }

  /*************************************************
   * ESCRIBIR EN ENTREGADO
   *************************************************/

  if (hojaEntregado.getLastRow() === 0){
    hojaEntregado.appendRow(encabezado);
  }

  hojaEntregado
    .getRange(
      hojaEntregado.getLastRow() + 1,
      1,
      mover.length,
      mover[0].length
    )
    .setValues(mover);

  /*************************************************
   * RECONSTRUIR INVENTARIO
   *************************************************/

  hojaInventario.clearContents();

  hojaInventario
    .getRange(1,1,1,encabezado.length)
    .setValues([encabezado]);

  if (mantener.length > 0){

    hojaInventario
      .getRange(2,1,mantener.length,mantener[0].length)
      .setValues(mantener);

  }

  Logger.log(mover.length + " remesas movidas a ENTREGADO");

}


