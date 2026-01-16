const HOJA_REGISTROS = "Registros";
const HOJA_MAESTRA = "Maestra";

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Solicita tu transporte')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(nombreArchivo) {
  return HtmlService.createHtmlOutputFromFile(nombreArchivo).getContent();
}

function validarAcceso(cedula) {
  if (!cedula) throw new Error("Debe ingresar su cédula.");
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hojaM = libro.getSheetByName(HOJA_MAESTRA);
  const datosM = hojaM.getDataRange().getValues();
  let usuario = null;

  for (let i = 1; i < datosM.length; i++) {
    if (datosM[i][0].toString() === cedula.toString()) {
      const partes = datosM[i][1].toString().trim().split(/\s+/);
      const nombreMostrar = partes.length >= 2 ? `${partes[0]} ${partes[1]}` : partes[0];
      usuario = {
        cedula: datosM[i][0].toString(),
        nombre: nombreMostrar,
        nombreCompleto: datosM[i][1],
        jefe: datosM[i][4],
        ceco: datosM[i][7],
        correoJefe: buscarCorreoPersona(datosM[i][4], datosM)
      };
      break;
    }
  }

  if (!usuario) throw new Error("Usuario no encontrado en la Maestra.");

  const hojaS = libro.getSheetByName(HOJA_REGISTROS);
  const datosS = hojaS.getDataRange().getValues();

  const historial = datosS.filter(f => f[3].toString() === cedula.toString()).map((fila, index) => {
    let fechaObj = fila[11] instanceof Date ? fila[11] : new Date(fila[11]);
    let fechaValida = !isNaN(fechaObj.getTime());

    return {
      idFila: index + 1,
      fecha: fechaValida ? Utilities.formatDate(fechaObj, "GMT-5", "dd/MM/yyyy") : "S/F",
      fechaIso: fechaValida ? Utilities.formatDate(fechaObj, "GMT-5", "yyyy-MM-dd") : null,
      trayecto: `${fila[9]} ➔ ${fila[10]}`,
      destino: fila[10],
      estado: fila[15] || "Pendiente",
      mes: fechaValida ? fechaObj.getMonth() : null,
      anio: fechaValida ? fechaObj.getFullYear() : null
    };
  }).reverse();

  const conteoDestinos = historial.reduce((acc, cur) => {
    if (cur.destino) acc[cur.destino] = (acc[cur.destino] || 0) + 1;
    return acc;
  }, {});

  const topDestinos = Object.entries(conteoDestinos)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 3)
    .map(d => d[0]);

  const tendencia = [];
  const hoy = new Date();
  for (let i = 9; i >= 0; i--) {
    const d = new Date(hoy.getFullYear(), hoy.getMonth() - i, 1);
    const mesEtiqueta = d.toLocaleString('es-ES', { month: 'short' }).toUpperCase();
    const cuenta = historial.filter(h => h.mes === d.getMonth() && h.anio === d.getFullYear()).length;
    tendencia.push({ mes: mesEtiqueta, total: cuenta });
  }

  return { usuario, historial, total: historial.length, topDestinos, tendencia };
}

function buscarCorreoPersona(nombre, matriz) {
  for (let i = 1; i < matriz.length; i++) {
    if (matriz[i][1] === nombre) return matriz[i][2];
  }
  return "pendiente@empresa.com";
}

function registrarSolicitud(f) {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = libro.getSheetByName(HOJA_REGISTROS);
  const datos = hoja.getDataRange().getValues();

  const existeDuplicado = datos.some(fila =>
    fila[3].toString() === f.cedula.toString() &&
    fila[11] instanceof Date &&
    Utilities.formatDate(fila[11], "GMT-5", "yyyy-MM-dd") === f.fecha &&
    fila[12].toString() === f.hora &&
    fila[15] !== "Cancelado"
  );

  if (existeDuplicado) throw new Error("Ya tienes una solicitud activa para esa misma fecha y hora.");

  const nuevaFila = [
    new Date(), f.virtual, Session.getActiveUser().getEmail(), f.cedula, f.nombre,
    f.celular, f.jefe, f.correoJefe, f.ceco,
    f.puntoRecogida === "Otro" ? f.otroR : f.puntoRecogida,
    f.puntoDestino === "Otro" ? f.otroD : f.puntoDestino,
    f.fecha, f.hora, f.motivo, "Pendiente"
  ];

  hoja.appendRow(nuevaFila);
  return "¡Solicitud registrada con éxito!";
}

function cambiarEstadoSolicitud(idFila, estado) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_REGISTROS);
  hoja.getRange(idFila, 16).setValue(estado);
  return "Estado actualizado.";
}