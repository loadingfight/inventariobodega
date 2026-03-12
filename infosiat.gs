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
  const nro = "220202318441";
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
 * completar Datos API
*************************************************/

function completarDatosAPI(){

  const hoja = SpreadsheetApp.getActive()
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


/*************************************************
 * Procesar Lote API
*************************************************/


function procesarLoteAPI(tamanoLote) {
  const hoja = SpreadsheetApp.getActive().getSheetByName("inventario");
  const datos = hoja.getDataRange().getValues();
  let contador = 0;
  let pendientesPostLote = 0;

  for (let i = 1; i < datos.length; i++) {
    const remesa = String(datos[i][0]).trim();
    const cliente = datos[i][8]; // Columna I

    // Si ya tiene datos, no contar como pendiente ni procesar
    if (cliente !== "" || !remesa) continue;

    // Si todavía estamos dentro del tamaño del lote, procesamos
    if (contador < tamanoLote) {
      try {
        const respuesta = GETALDIA.consultar(remesa);
        if (respuesta && respuesta.rows > 0) {
          const info = respuesta.data[0];
          let uEv = (info.historico_eventos && info.historico_eventos.length > 0) 
                    ? info.historico_eventos.slice(-1)[0].nombre : "SIN EVENTOS";

          hoja.getRange(i + 1, 9, 1, 7).setValues([[
            info.cliente, info.destino, info.estado_remesa, 
            info.entregada, info.fecha_hora, info.anulada, uEv
          ]]);
        } else {
          hoja.getRange(i + 1, 9).setValue("N/A"); // Evitar re-procesar fallidas
        }
        contador++;
        Utilities.sleep(1000); // Pequeño respiro para la API
      } catch (e) {
        Logger.log("Error remesa " + remesa);
      }
    } else {
      // Si ya llenamos el lote, solo contamos cuántas quedan pendientes en total
      pendientesPostLote++;
    }
  }

  return {
    procesados: contador,
    pendientes: pendientesPostLote
  };
}