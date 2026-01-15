const HOJA_SOLICITUDES = "Registros";
const HOJA_MAESTRA = "Maestra";

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Solicitud de Transporte ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function validarAcceso(cedula) {
  if (!cedula) throw new Error("La cédula es requerida.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  actualizarEstadosEjecutados(ss);

  const hojaM = ss.getSheetByName(HOJA_MAESTRA);
  const datosM = hojaM.getDataRange().getValues();

  let usuario = null;
  for (let i = 1; i < datosM.length; i++) {
    if (datosM[i][0].toString() === cedula.toString()) {
      const nombreLider = datosM[i][4];
      usuario = {
        cedula: datosM[i][0].toString(),
        nombre: datosM[i][1],
        jefe: nombreLider,
        ceco: datosM[i][7],
        correoJefe: buscarCorreoPersona(nombreLider, datosM)
      };
      break;
    }
  }

  if (!usuario) throw new Error("El documento " + cedula + " no existe en la Maestra.");

  const hojaS = ss.getSheetByName(HOJA_SOLICITUDES);
  const datosS = hojaS.getDataRange().getValues();

  const historial = datosS.map((fila, index) => ({
    idFila: index + 1,
    cedulaRegistro: fila[3] ? fila[3].toString() : "",
    fecha: fila[11] ? Utilities.formatDate(new Date(fila[11]), "GMT-5", "dd/MM/yyyy") : "S/F",
    trayecto: `${fila[9]} ➔ ${fila[10]}`,
    motivo: fila[14],
    estado: fila[15] || "PENDIENTE"
  })).filter(item => item.cedulaRegistro === cedula.toString()).reverse();

  return { usuario, historial, total: historial.length };
}

function buscarCorreoPersona(nombre, matriz) {
  for (let i = 1; i < matriz.length; i++) {
    if (matriz[i][1] === nombre) return matriz[i][2];
  }
  return "correo@empresa.com";
}

function actualizarEstadosEjecutados(ss) {
  const hoja = ss.getSheetByName(HOJA_SOLICITUDES);
  const datos = hoja.getDataRange().getValues();
  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][11] instanceof Date) {
      let fechaServicio = new Date(datos[i][11]);
      if (fechaServicio < hoy && datos[i][15] === "PENDIENTE") {
        hoja.getRange(i + 1, 16).setValue("EJECUTADO");
      }
    }
  }
}

function registrarSolicitud(f) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(HOJA_SOLICITUDES);
  const fila = [
    new Date(), f.virtual, Session.getActiveUser().getEmail(), f.cedula, f.nombre,
    f.celular, f.jefe, f.correoJefe, f.ceco,
    f.puntoRecogida === "Otro" ? f.otroR : f.puntoRecogida,
    f.puntoDestino === "Otro" ? f.otroD : f.puntoDestino,
    f.fecha, f.hora, f.hora2, f.motivo, "PENDIENTE"
  ];
  hoja.appendRow(fila);
  return "Registro guardado correctamente.";
}

function cambiarEstadoSolicitud(idFila, nuevoEstado) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(HOJA_SOLICITUDES);
  hoja.getRange(idFila, 16).setValue(nuevoEstado);
  return "Estado actualizado.";
}