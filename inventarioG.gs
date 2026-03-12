/*************************************************
 * FUNCIONES GLOBLALES
 *************************************************/
function actualizarInventarioCompleto(){

  limpiarErroresAPIInterno();

  actualizarInventarioGeneral();   

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
 * OBTENER Y COPIAR ALERTAS (Muchos dias)
 *************************************************/
function obtenerAlertasInventario() {
  const ss = SpreadsheetApp.getActive();
  const hojaInventario = ss.getSheetByName(HOJA_INVENTARIO);
  let hojaAlertas = ss.getSheetByName(HOJA_ALERTAS);

  // 1. Validaciones iniciales
  if (!hojaInventario) return [];
  if (!hojaAlertas) {
    hojaAlertas = ss.insertSheet(HOJA_ALERTAS);
  }

  const datos = hojaInventario.getDataRange().getValues();
  if (datos.length <= 1) return [];

  const encabezados = datos[0];
  
  // 2. Filtrar solo las filas que tengan "Muchos dias" en la columna P (índice 15)
  const filasAlertas = datos.slice(1).filter(fila => {
    const alerta = String(fila[15] || "").toUpperCase().trim();
    return alerta.includes("MUCHOS DIAS"); 
  });

  // 3. COPIAR A LA HOJA "Alertas"
  // Limpiamos la hoja de alertas antes de copiar las nuevas
  hojaAlertas.clearContents();
  
  if (filasAlertas.length > 0) {
    // Ponemos los encabezados
    hojaAlertas.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
    // Pegamos las filas encontradas
    hojaAlertas.getRange(2, 1, filasAlertas.length, encabezados.length).setValues(filasAlertas);
    
    Logger.log("Se copiaron " + filasAlertas.length + " alertas a la hoja Alertas.");
  }

  // 4. Retornar los datos al HTML para que se vean en la tabla web
  if (filasAlertas.length === 0) return [];
  
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

  if (!hojaInventario) throw new Error("No existe la hoja: " + HOJA_INVENTARIO);
  if (!hojaEntregado) hojaEntregado = ss.insertSheet(HOJA_ENTREGADO);

  const datos = hojaInventario.getDataRange().getValues();
  if (datos.length <= 1) return;

  const encabezado = datos[0];
  const filas = datos.slice(1);

  // --- 1. CARGAR REMESAS YA EXISTENTES EN ENTREGADOS (EVITAR DUPLICADOS) ---
  let yaEntregadas = new Set();
  if (hojaEntregado.getLastRow() > 0) {
    // Leemos la columna A de la hoja de entregados
    const idsEntregados = hojaEntregado.getRange(1, 1, hojaEntregado.getLastRow(), 1).getValues().flat();
    yaEntregadas = new Set(idsEntregados.map(id => String(id).trim()));
  }

  const mover = [];
  const mantener = [];

  filas.forEach(fila => {
    const remesa = String(fila[0]).trim();
    const colE = String(fila[4]).trim();                 // % 1
    const colG = String(fila[6]).trim().toUpperCase();   // Documento
    const colK = String(fila[10]).trim().toUpperCase();  // Estado
    const colO = String(fila[14]).trim().toUpperCase();  // Evento

    // Condición: 100% escaneado, Documento SI, Estado CUMPLIDA y Evento ENTREGADO OK
    // Nota: colE puede llegar como "1", "1.0" o "100%", por eso usamos includes o Number
    if (
      (colE === "1" || colE === "100%") && 
      colG === "SI" &&
      colK === "CUMPLIDA" &&
      colO === "ENTREGADO OK"
    ){
      // SOLO mover si NO existe ya en la hoja de entregados
      if (!yaEntregadas.has(remesa)) {
        mover.push(fila);
      }
      // Si ya existía, simplemente no la añadimos a 'mantener', así se borra del inventario
    } else {
      mantener.push(fila);
    }
  });

  // --- 2. ESCRIBIR EN ENTREGADO (Si hay nuevos) ---
  if (mover.length > 0) {
    if (hojaEntregado.getLastRow() === 0) {
      hojaEntregado.appendRow(encabezado);
    }
    hojaEntregado.getRange(
      hojaEntregado.getLastRow() + 1,
      1,
      mover.length,
      mover[0].length
    ).setValues(mover);
    Logger.log(mover.length + " remesas nuevas movidas a ENTREGADO.");
  }

  // --- 3. RECONSTRUIR HOJA DE INVENTARIO (LIMPIEZA) ---
  hojaInventario.clearContents();
  hojaInventario.getRange(1, 1, 1, encabezado.length).setValues([encabezado]);
  
  if (mantener.length > 0) {
    hojaInventario.getRange(2, 1, mantener.length, mantener[0].length).setValues(mantener);
  }
  
  Logger.log("Inventario actualizado. Se mantienen " + mantener.length + " filas.");
}


/*************************************************
   * limpiar Errores API
*************************************************/

function limpiarErroresAPIInterno() {
  const ss = SpreadsheetApp.getActive();
  const hoja = ss.getSheetByName(HOJA_INVENTARIO);
  if (!hoja) return;
  
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return;

  // Rango de la columna I (Cliente)
  const rango = hoja.getRange(2, 9, ultimaFila - 1, 1);
  const valores = rango.getValues();
  
  const nuevosValores = valores.map(fila => {
    const txt = String(fila[0]).toUpperCase();
    // Si contiene N/A o No en API, lo dejamos vacío para que el script lo vuelva a intentar
    if (txt.includes("N/A") || txt.includes("NO ENCONTRADA") || txt.includes("ERROR")) {
      return [""]; 
    }
    return [fila[0]];
  });

  rango.setValues(nuevosValores);
}