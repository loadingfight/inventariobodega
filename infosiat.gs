function seguimientoLogistica() {
  // Aquí puedes poner cualquier número de remesa que necesites
  const remesaPrueba = "220500672440"; 
  const remesaOtra = "22269070450";
  
  // Llamada: IdentificadorDeBiblioteca.NombreDeFuncion(numero)
  // Asumiendo que el identificador que pusiste al añadir la biblioteca es 'Aldia'
  const resultado = GETALDIA.consultar(remesaPrueba);
  const resultado2 = GETALDIA.consultar(remesaOtra);
  
  Logger.log("Resultado 1: " + JSON.stringify(resultado));
  Logger.log("Resultado 2: " + JSON.stringify(resultado2));
}



function obtenerResumenRemesa() {
  const nro = "220500672440";
  const respuesta = GETALDIA.consultar(nro);
  
  // Verificamos que la respuesta tenga datos
  if (respuesta && respuesta.rows > 0) {
    const info = respuesta.data[0]; // Accedemos al primer registro
    
    const resumen = {
      remesa: info.remesa,
      cliente: info.cliente,
      destino: info.destino,
      estado: info.estado_remesa, // Ej: "CUMPLIDA"
      entregada: info.entregada,    // Ej: "SI"
      fechaEntrega: info.fecha_entrega,
      ultimoEvento: info.historico_eventos.slice(-1)[0].nombre
    };
    
    Logger.log("--- RESUMEN DE LA REMESA ---");
    Logger.log("Estado: " + resumen.estado);
    Logger.log("Último movimiento: " + resumen.ultimoEvento);
    
    return resumen;
  } else {
    Logger.log("No se encontró información para la remesa: " + nro);
    return null;
  }
}



/*************************************************
 * ACTUALIZAR INVENTARIO DESDE SIAT (OPTIMIZADO)
 
function actualizarInventarioDesdeSIAT() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaInv = ss.getSheetByName("Inventario");
  const hojaSIAT = ss.getSheetByName("SIAT");

  if (!hojaInv || !hojaSIAT) {
    Logger.log("Faltan hojas necesarias");
    return;
  }

  const datosInv = hojaInv.getDataRange().getValues();
  const datosSIAT = hojaSIAT.getDataRange().getValues();

  // ==============================
  // CREAR MAPA SIAT
  // ==============================
  let mapaSIAT = {};

  for (let i = 1; i < datosSIAT.length; i++) {

    const remesa = String(datosSIAT[i][0]).trim();

    if (remesa) {
      mapaSIAT[remesa] = {
        fechaDesp: datosSIAT[i][20], // Col U
        evento: datosSIAT[i][23]     // Col X
      };
    }
  }

  const hoy = new Date();

  // ==============================
  // ACTUALIZAR INVENTARIO EN MEMORIA
  // ==============================
  for (let i = 1; i < datosInv.length; i++) {

    const remesaInv = String(datosInv[i][0]).trim();

    if (mapaSIAT[remesaInv]) {

      const info = mapaSIAT[remesaInv];

      datosInv[i][7] = "SI"; // Col H
      datosInv[i][8] = info.evento; // Col I
      datosInv[i][9] = info.fechaDesp; // Col J

      if (info.fechaDesp instanceof Date) {
        const dias = Math.floor((hoy - info.fechaDesp) / (1000 * 60 * 60 * 24));
        datosInv[i][10] = dias; // Col K
      } else {
        datosInv[i][10] = "";
      }

    } else {

      datosInv[i][7] = "NO";
      datosInv[i][8] = "";
      datosInv[i][9] = "";
      datosInv[i][10] = "";
    }
  }

  // ==============================
  // ESCRIBIR TODO DE UNA SOLA VEZ
  // ==============================
  hojaInv.getRange(1,1,datosInv.length,datosInv[0].length)
         .setValues(datosInv);

  Logger.log("Inventario sincronizado con SIAT correctamente");
}
*************************************************/