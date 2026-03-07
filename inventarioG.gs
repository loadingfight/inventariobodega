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
  hojaInventario.getRange(1,1,1,13).setValues([[
    "Remesa",
    "Total",
    "Escaneadas",
    "Faltantes",
    "%",
    "Ubicación",
    "Documento",
    "FechaLec",
    "Cliente",
    "Destino",
    "Estado",
    "Entregada",
    "FechaCreación" 
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


/*************************************************
 * DATOS DESDE LA API
 *************************************************/

function completarDatosAPI(){

  const hoja = SpreadsheetApp
    .getActive()
    .getSheetByName("inventario");

  const datos = hoja
    .getRange(2,1,hoja.getLastRow()-1,13)
    .getValues();

  const LIMITE = 50;
  let contador = 0;

  datos.forEach((fila,i)=>{

    const remesa = String(fila[0]).trim();
    const cliente = fila[8]; // columna I

    // si ya tiene datos no consulta
    if(cliente) return;

    if(contador >= LIMITE) return;

    try{

      const respuesta = GETALDIA.consultar(remesa);

      if(respuesta && respuesta.rows > 0){

        const info = respuesta.data[0];

        const cliente = info.cliente;
        const destino = info.destino;
        const estado = info.estado_remesa;
        const entregada = info.entregada;
        const fecha = info.fecha_hora;

        hoja.getRange(i+2,9,1,5).setValues([[
          cliente,
          destino,
          estado,
          entregada,
          fecha
        ]]);

      }

      contador++;

      Utilities.sleep(1200);

    }catch(e){

      Logger.log("Error remesa "+remesa);

    }

  });

}