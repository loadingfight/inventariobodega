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