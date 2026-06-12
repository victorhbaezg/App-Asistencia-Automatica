function doGet(e) {
  return manejarSolicitud(e);
}

function doPost(e) {
  return manejarSolicitud(e);
}

function manejarSolicitud(e) {
  const params = obtenerParametros(e);
  const action = params.action || '';
  const callback = params.callback || '';

  if (esAccionRegistro(action)) {
    try {
      const idEscaneado = obtenerId(params);
      if (!idEscaneado) throw new Error('ID escaneado vacio.');

      const mensaje = registrarAsistencia({
        id_escaneado: idEscaneado
      });

      return responder({
        status: 'ok',
        message: mensaje,
        id_escaneado: idEscaneado
      }, callback);
    } catch (err) {
      return responder({ status: 'error', message: err.message }, callback);
    }
  }

  // Sin parametros sirve el HTML por si alguien entra directo a la URL /exec.
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Registro de Asistencia')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function obtenerParametros(e) {
  const params = e && e.parameter ? Object.assign({}, e.parameter) : {};

  if (e && e.postData && e.postData.contents) {
    const contentType = e.postData.type || '';
    if (contentType.indexOf('application/json') !== -1) {
      try {
        Object.assign(params, JSON.parse(e.postData.contents));
      } catch (err) {
        throw new Error('JSON invalido.');
      }
    }
  }

  return params;
}

function esAccionRegistro(action) {
  return [
    'registrarAsistencia',
    'registrar',
    'guardar',
    'manual'
  ].indexOf(action) !== -1;
}

function obtenerId(params) {
  return String(
    params.id_escaneado ||
    params.idEscaneado ||
    params.id ||
    params.manualId ||
    params.codigo ||
    ''
  ).trim();
}

function responder(data, callback) {
  if (callback) {
    const callbackSeguro = String(callback).replace(/[^\w.$]/g, '');
    return ContentService
      .createTextOutput(callbackSeguro + '(' + JSON.stringify(data) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function registrarAsistencia(datos) {
  const nombreHoja = 'Registro_Asistencia';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName(nombreHoja);

  if (!hoja) {
    hoja = ss.insertSheet(nombreHoja);
    hoja.appendRow(['ID Registro', 'ID Escaneado', 'Fecha', 'Hora', 'Usuario']);
  }

  const ahora = new Date();
  const idRegistro = 'REG-' + ahora.getTime();
  const fecha = Utilities.formatDate(ahora, 'America/Mexico_City', 'dd/MM/yyyy');
  const hora = Utilities.formatDate(ahora, 'America/Mexico_City', 'HH:mm:ss');
  const usuario = Session.getActiveUser().getEmail() || 'Anonimo';

  hoja.appendRow([idRegistro, datos.id_escaneado, fecha, hora, usuario]);
  return 'Asistencia registrada correctamente.';
}
