/*************************************************
 * SEGUIMIENTO DESDE API
*************************************************/

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

/*************************************************
 * ACTUALIZAR INVENTARIO DESDE API
*************************************************/

function obtenerResumenRemesa() {
  const nro = "222603244307";
  const respuesta = GETALDIA.consultar(nro);
  
  // Verificamos que la respuesta tenga datos
  if (respuesta && respuesta.rows > 0) {
    const info = respuesta.data[0]; // Accedemos al primer registro
    
    const resumen = {
      remesa: info.remesa,
      cliente: info.cliente,
      destino: info.destino,
      estado: info.estado_remesa,
      anulada: info.anulada, 
      entregada: info.entregada,    
      fecha_hora: info.fecha_hora,
      ultimoEvento: info.historico_eventos.slice(-1)[0].nombre
    };
    
    Logger.log("--- RESUMEN DE LA REMESA ---");
    Logger.log("Cliente: " + resumen.cliente);
    Logger.log("Destino: " + resumen.destino);
    Logger.log("fecha_hora: " + resumen.fecha_hora);
    Logger.log("Estado: " + resumen.estado);
    Logger.log("Anulada: " + resumen.anulada);
    Logger.log("Último movimiento: " + resumen.ultimoEvento);
    
    return resumen;
  } else {
    Logger.log("No se encontró información para la remesa: " + nro);
    return null;
  }
}




/*************************************************
 * completarDatosAPI
*************************************************/

function completarDatosAPI(){

  const hoja = SpreadsheetApp
    .getActive()
    .getSheetByName("inventario");

  const datos = hoja
    .getRange(2,1,hoja.getLastRow()-1,15)
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
        const anulada = info.anulada;
        const ultimoEvento = info.historico_eventos.slice(-1)[0].nombre

        hoja.getRange(i+2,9,1,7).setValues([[
          cliente,
          destino,
          estado,
          entregada,
          fecha,
          anulada,
          ultimoEvento
        ]]);

      }

      contador++;

      Utilities.sleep(1500);

    }catch(e){

      Logger.log("Error remesa "+remesa);

    }

  });

}
 