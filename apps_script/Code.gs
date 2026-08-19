/**
 * UdeSA Horarios — Bridge privado de fuentes
 * Autor de UdeSA Horarios: Juan Ignacio Gutiérrez Julián
 *
 * IMPORTANTE:
 * - Desplegar como Web app.
 * - Ejecutar como: yo (usuario que despliega).
 * - Acceso: cualquiera / anyone anonymous si la cuenta lo permite.
 * - La seguridad real la aporta UDESA_TOKEN enviado por POST y almacenado en Script Properties.
 * - Los IDs de Drive se guardan en Script Properties; nunca quedan en el repositorio público.
 */

const PROP_TOKEN = 'UDESA_TOKEN';
const PROP_SOURCES = 'UDESA_SOURCES_JSON';

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getConfig_() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty(PROP_TOKEN);
  const raw = props.getProperty(PROP_SOURCES);
  if (!token || !raw) throw new Error('Bridge no configurado. Ejecutá configurarBridge() una vez.');
  return {token: token, sources: JSON.parse(raw)};
}

function readBody_(e) {
  if (!e || !e.postData || !e.postData.contents) return {};
  return JSON.parse(e.postData.contents);
}

function checkToken_(given, expected) {
  if (!given || given !== expected) throw new Error('TOKEN_INVALIDO');
}

function metadataFor_(key, source) {
  const file = DriveApp.getFileById(source.fileId);
  return {
    key: key,
    name: source.displayName || file.getName(),
    modified: file.getLastUpdated().toISOString(),
    size: file.getSize(),
    mimeType: file.getMimeType()
  };
}

function bytesFor_(source) {
  const file = DriveApp.getFileById(source.fileId);
  const mime = file.getMimeType();
  // Fuentes Office almacenadas en Drive: devuelve el archivo original, preservando comentarios del XLSX.
  if (mime !== MimeType.GOOGLE_SHEETS) {
    return {blob: file.getBlob(), mimeType: mime};
  }
  // Fallback futuro para Google Sheets nativo. Nota: una exportación puede no conservar comentarios de Drive.
  const exportUrl = 'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(source.fileId)
    + '/export?mimeType=' + encodeURIComponent('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  const resp = UrlFetchApp.fetch(exportUrl, {
    headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()},
    muteHttpExceptions: false
  });
  return {blob: resp.getBlob(), mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'};
}

function doPost(e) {
  try {
    const cfg = getConfig_();
    const body = readBody_(e);
    checkToken_(body.token, cfg.token);
    const action = body.action || 'ping';

    if (action === 'ping') {
      return jsonResponse_({ok: true, service: 'UdeSA Horarios Source Bridge'});
    }

    if (action === 'metadata_all') {
      const rows = Object.keys(cfg.sources).map(k => metadataFor_(k, cfg.sources[k]));
      return jsonResponse_({ok: true, sources: rows});
    }

    if (action === 'metadata') {
      const source = cfg.sources[body.source];
      if (!source) throw new Error('FUENTE_DESCONOCIDA');
      return jsonResponse_({ok: true, source: metadataFor_(body.source, source)});
    }

    if (action === 'download') {
      const source = cfg.sources[body.source];
      if (!source) throw new Error('FUENTE_DESCONOCIDA');
      const meta = metadataFor_(body.source, source);
      const result = bytesFor_(source);
      const bytes = result.blob.getBytes();
      return jsonResponse_({
        ok: true,
        key: body.source,
        name: meta.name,
        modified: meta.modified,
        size: bytes.length,
        mimeType: result.mimeType,
        content_base64: Utilities.base64Encode(bytes)
      });
    }

    throw new Error('ACCION_DESCONOCIDA');
  } catch (err) {
    return jsonResponse_({ok: false, error: String(err && err.message ? err.message : err)});
  }
}

/**
 * EJECUTAR MANUALMENTE UNA SOLA VEZ DESPUÉS DE COMPLETAR LOS IDs.
 * Podés volver a ejecutarla al cambiar de semestre/año.
 */
function configurarBridge() {
  const TOKEN = 'REEMPLAZAR_POR_TOKEN_LARGO_ALEATORIO';
  const SOURCES = {
    aulas_2026_primavera: {fileId: 'PEGAR_ID_AULAS', displayName: 'AULAS- PRIMAVERA 2026- FINAL.xlsx'},
    cursos_2026_primavera: {fileId: 'PEGAR_ID_CURSOS', displayName: 'CURSOS PRIMAVERA 26.xlsx'},
    eventos_2026: {fileId: 'PEGAR_ID_EVENTOS', displayName: 'Registro de actividades 2026.xlsx'},
    plan_master: {fileId: 'PEGAR_ID_PLAN_MASTER', displayName: 'udesa_plan_academico_master.xlsx'},
    catalogo_espacios: {fileId: 'PEGAR_ID_CATALOGO_ESPACIOS', displayName: 'aulas_udesA.xlsx'},
    catalogo_legacy: {fileId: 'PEGAR_ID_CATALOGO_LEGACY', displayName: 'Area de Charlas, Camada y codigos.xlsx'}
  };
  if (TOKEN.indexOf('REEMPLAZAR_') === 0) throw new Error('Primero reemplazá TOKEN e IDs.');
  PropertiesService.getScriptProperties().setProperties({
    [PROP_TOKEN]: TOKEN,
    [PROP_SOURCES]: JSON.stringify(SOURCES)
  });
  Logger.log('Bridge configurado con %s fuentes.', Object.keys(SOURCES).length);
}

function probarBridgeLocal() {
  const cfg = getConfig_();
  Logger.log(JSON.stringify(Object.keys(cfg.sources).map(k => metadataFor_(k, cfg.sources[k])), null, 2));
}
