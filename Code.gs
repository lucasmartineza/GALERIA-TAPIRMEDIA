// GALERIA TAPIR — Backend Apps Script
// v0.2 — conectado directo a la planilla activa (Extensiones > Apps Script),
//         no hace falta pegar ningun ID de Sheet a mano.
//
// SETUP:
// 1. Este codigo va pegado en el editor que se abre desde
//    la planilla "Indice Galeria Tapir" > Extensiones > Apps Script.
// 2. La planilla debe tener una hoja llamada "indice" con columnas:
//    folder_id | nombre | apellido | numero_auto | evento | activo
// 3. Implementar > Nueva implementacion > Aplicacion web.
//    Ejecutar como: "Yo". Quien tiene acceso: "Cualquier usuario".
// 4. Copiar la URL de implementacion y pegarla en CONFIG.API_URL de index.html.

var SHEET_NAME = 'indice';
var ADMIN_KEY = 'tapir2026'; // cambiá esto por algo tuyo antes de publicar admin.html

function doGet(e) {
  var action = e.parameter.action;
  try {
    if (action === 'login') return handleLogin_(e);
    if (action === 'files') return handleFiles_(e);
    if (action === 'zip') return handleZip_(e);
    if (action === 'avisos') return handleAvisos_(e);
    if (action === 'nuevo_aviso') return handleNuevoAviso_(e);
    if (action === 'ubicaciones') return handleUbicaciones_(e);
    if (action === 'nueva_ubicacion') return handleNuevaUbicacion_(e);
    if (action === 'quitar_ubicacion') return handleQuitarUbicacion_(e);
    if (action === 'guardar_ruta') return handleGuardarRuta_(e);
    if (action === 'ruta') return handleRuta_(e);
    return jsonOutput_({ error: 'accion invalida' });
  } catch (err) {
    return jsonOutput_({ error: err.message });
  }
}

// ---------- LOGIN ----------
function handleLogin_(e) {
  var nombre = normalizar_(e.parameter.nombre || '');
  var apellido = normalizar_(e.parameter.apellido || '');
  if (!nombre || !apellido) return jsonOutput_({ error: 'faltan datos' });

  var rows = getIndiceRows_();
  var matches = rows.filter(function (r) {
    return r.activo &&
      normalizar_(r.nombre) === nombre &&
      normalizar_(r.apellido) === apellido;
  });

  if (matches.length === 0) return jsonOutput_({ status: 'no_match' });
  if (matches.length === 1) {
    return jsonOutput_({ status: 'ok', cliente: matches[0] });
  }
  // varios matches -> pedir desambiguacion por numero de auto
  return jsonOutput_({
    status: 'multiple',
    opciones: matches.map(function (m) {
      return { folder_id: m.folder_id, numero_auto: m.numero_auto, evento: m.evento };
    })
  });
}

function normalizar_(s) {
  return String(s)
    .toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '') // saca tildes
    .trim()
    .replace(/\s+/g, ' ');
}

function getIndiceRows_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var headers = data.shift();
  return data.map(function (row) {
    var obj = {};
    headers.forEach(function (h, i) { obj[h] = row[i]; });
    return obj;
  });
}

// ---------- LISTADO DE ARCHIVOS ----------
function handleFiles_(e) {
  var folderId = e.parameter.folder;
  if (!folderId) return jsonOutput_({ error: 'falta folder' });

  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var out = [];

  while (files.hasNext()) {
    var f = files.next();
    // aseguramos que sea visible por link (una sola vez, Drive no repite trabajo si ya esta)
    try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (err) {}

    // traemos la miniatura ya generada por Drive y la mandamos incrustada
    // (data URI) para que el navegador no tenga que pedirla aparte a Google
    var thumbnailUrl;
    try {
      var thumbBlob = f.getThumbnail();
      if (thumbBlob) {
        thumbnailUrl = 'data:' + thumbBlob.getContentType() + ';base64,' + Utilities.base64Encode(thumbBlob.getBytes());
      } else {
        thumbnailUrl = 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w300';
      }
    } catch (err) {
      thumbnailUrl = 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w300';
    }

    out.push({
      id: f.getId(),
      nombre: f.getName(),
      mimeType: f.getMimeType(),
      fechaCreacion: f.getDateCreated().getTime(),
      esVideo: f.getMimeType().indexOf('video') === 0,
      thumbnailUrl: thumbnailUrl,
      previewUrl: 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w1600',
      viewUrl: 'https://drive.google.com/uc?export=view&id=' + f.getId(),
      downloadUrl: 'https://drive.google.com/uc?export=download&id=' + f.getId()
    });
  }

  // mas nuevo primero
  out.sort(function (a, b) { return b.fechaCreacion - a.fechaCreacion; });

  return jsonOutput_({ status: 'ok', archivos: out });
}

// ---------- DESCARGA TOTAL (ZIP) ----------
function handleZip_(e) {
  var folderId = e.parameter.folder;
  if (!folderId) return jsonOutput_({ error: 'falta folder' });

  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var blobs = [];

  while (files.hasNext()) {
    blobs.push(files.next().getBlob());
  }

  if (blobs.length === 0) return jsonOutput_({ error: 'no hay archivos' });

  var nombreZip = (folder.getName() || 'material') + '.zip';
  var zipBlob = Utilities.zip(blobs, nombreZip);
  return zipBlob; // Apps Script dispara la descarga directo con el content-type correcto
}

// ---------- AVISOS (canal de comunicación masiva) ----------
// Hoja "avisos" con columnas: mensaje | fecha | activo
// Se crea sola la primera vez que se publica un aviso desde admin.html.
function handleAvisos_(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('avisos');
  if (!sheet) return jsonOutput_({ status: 'ok', avisos: [] });

  var data = sheet.getDataRange().getValues();
  var headers = data.shift();
  var avisos = data
    .map(function (row) {
      var obj = {};
      headers.forEach(function (h, i) { obj[h] = row[i]; });
      return obj;
    })
    .filter(function (a) { return a.activo; })
    .map(function (a) {
      return { mensaje: a.mensaje, fecha: a.fecha ? String(a.fecha) : '' };
    });

  return jsonOutput_({ status: 'ok', avisos: avisos });
}

// ---------- PUBLICAR AVISO NUEVO (desde admin.html) ----------
function handleNuevoAviso_(e) {
  if (e.parameter.key !== ADMIN_KEY) return jsonOutput_({ error: 'clave incorrecta' });

  var mensaje = (e.parameter.mensaje || '').trim();
  if (!mensaje) return jsonOutput_({ error: 'mensaje vacio' });

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('avisos');
  if (!sheet) {
    sheet = ss.insertSheet('avisos');
    sheet.appendRow(['mensaje', 'fecha', 'activo']);
  }

  var fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM HH:mm');
  sheet.appendRow([mensaje, fecha, true]);

  return jsonOutput_({ status: 'ok' });
}

// ---------- UBICACIONES (mapa de dónde está el equipo) ----------
// Hoja "ubicaciones" con columnas: id | lat | lng | tipo | evento | activo | fecha
function getOrCrearHojaUbicaciones_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('ubicaciones');
  if (!sheet) {
    sheet = ss.insertSheet('ubicaciones');
    sheet.appendRow(['id', 'lat', 'lng', 'tipo', 'evento', 'activo', 'fecha']);
  }
  return sheet;
}

function handleUbicaciones_(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ubicaciones');
  if (!sheet) return jsonOutput_({ status: 'ok', ubicaciones: [] });

  var evento = (e.parameter.evento || '').trim().toLowerCase();
  var data = sheet.getDataRange().getValues();
  var headers = data.shift();
  var ubicaciones = data
    .map(function (row) {
      var obj = {};
      headers.forEach(function (h, i) { obj[h] = row[i]; });
      return obj;
    })
    .filter(function (u) {
      if (!u.activo) return false;
      if (evento && String(u.evento || '').trim().toLowerCase() !== evento) return false;
      return true;
    })
    .map(function (u) {
      return {
        id: u.id, lat: u.lat, lng: u.lng, tipo: u.tipo,
        evento: u.evento, fecha: u.fecha, timestamp: u.id
      };
    });

  return jsonOutput_({ status: 'ok', ubicaciones: ubicaciones });
}

function handleNuevaUbicacion_(e) {
  if (e.parameter.key !== ADMIN_KEY) return jsonOutput_({ error: 'clave incorrecta' });

  var lat = parseFloat(e.parameter.lat);
  var lng = parseFloat(e.parameter.lng);
  var tipo = e.parameter.tipo === 'video' ? 'video' : 'foto';
  var evento = (e.parameter.evento || '').trim();

  if (isNaN(lat) || isNaN(lng) || !evento) return jsonOutput_({ error: 'faltan datos' });

  var sheet = getOrCrearHojaUbicaciones_();
  var id = Date.now();
  var fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM HH:mm');
  sheet.appendRow([id, lat, lng, tipo, evento, true, fecha]);

  return jsonOutput_({ status: 'ok', id: id });
}

function handleQuitarUbicacion_(e) {
  if (e.parameter.key !== ADMIN_KEY) return jsonOutput_({ error: 'clave incorrecta' });

  var id = e.parameter.id;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ubicaciones');
  if (!sheet) return jsonOutput_({ error: 'no hay ubicaciones' });

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.getRange(i + 1, 6).setValue(false); // columna "activo"
      return jsonOutput_({ status: 'ok' });
    }
  }
  return jsonOutput_({ error: 'no encontrado' });
}

// ---------- RUTA DEL EVENTO (KMZ) ----------
// Hoja "eventos" con columnas: evento | kmz_file_id
function getOrCrearHojaEventos_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('eventos');
  if (!sheet) {
    sheet = ss.insertSheet('eventos');
    sheet.appendRow(['evento', 'kmz_file_id']);
  }
  return sheet;
}

function handleGuardarRuta_(e) {
  if (e.parameter.key !== ADMIN_KEY) return jsonOutput_({ error: 'clave incorrecta' });

  var evento = (e.parameter.evento || '').trim();
  var kmzId = (e.parameter.kmz_file_id || '').trim();
  if (!evento || !kmzId) return jsonOutput_({ error: 'faltan datos' });

  var sheet = getOrCrearHojaEventos_();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === evento.toLowerCase()) {
      sheet.getRange(i + 1, 2).setValue(kmzId);
      return jsonOutput_({ status: 'ok' });
    }
  }
  sheet.appendRow([evento, kmzId]);
  return jsonOutput_({ status: 'ok' });
}

function handleRuta_(e) {
  var evento = (e.parameter.evento || '').trim().toLowerCase();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('eventos');
  if (!sheet || !evento) return jsonOutput_({ status: 'ok', kml: null });

  var data = sheet.getDataRange().getValues();
  var kmzId = null;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === evento) {
      kmzId = data[i][1];
      break;
    }
  }
  if (!kmzId) return jsonOutput_({ status: 'ok', kml: null });

  try {
    var blob = DriveApp.getFileById(kmzId).getBlob();
    var zipBlob = blob.copyBlob().setContentType('application/zip');
    var archivos = Utilities.unzip(zipBlob);
    var kmlBlob = null;
    for (var j = 0; j < archivos.length; j++) {
      if (archivos[j].getName().toLowerCase().indexOf('.kml') > -1) {
        kmlBlob = archivos[j];
        break;
      }
    }
    if (!kmlBlob) return jsonOutput_({ status: 'ok', kml: null });
    return jsonOutput_({ status: 'ok', kml: kmlBlob.getDataAsString() });
  } catch (err) {
    return jsonOutput_({ status: 'ok', kml: null, error: err.message });
  }
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
