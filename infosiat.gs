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
  const nro = "222603197688";
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




 