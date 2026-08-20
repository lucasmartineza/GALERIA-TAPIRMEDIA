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
    if (action === 'eliminar_aviso') return handleEliminarAviso_(e);
    if (action === 'guardar_email') return handleGuardarEmail_(e);
    if (action === 'notificar_ahora') return handleNotificarAhora_(e);
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

  if (matches.length === 1) {
    return jsonOutput_({ status: 'ok', cliente: matches[0] });
  }
  if (matches.length > 1) {
    // varios matches -> pedir desambiguacion por numero de auto
    return jsonOutput_({
      status: 'multiple',
      opciones: matches.map(function (m) {
        return { folder_id: m.folder_id, numero_auto: m.numero_auto, evento: m.evento };
      })
    });
  }

  // sin match en indice -> puede ser un CM con acceso a varias carpetas
  var cmResultado = buscarAccesosCM_(nombre, apellido, rows);
  if (cmResultado) return jsonOutput_({ status: 'ok_cm', cm: cmResultado });

  return jsonOutput_({ status: 'no_match' });
}

function normalizar_(s) {
  return String(s)
    .toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '') // saca tildes
    .trim()
    .replace(/\s+/g, ' ');
}

function buscarAccesosCM_(nombreNorm, apellidoNorm, filasIndice) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('cm_accesos');
  if (!sheet) return null;

  var data = sheet.getDataRange().getValues();
  var headers = data.shift();
  var filas = data.map(function (row) {
    var obj = {};
    headers.forEach(function (h, i) { obj[h] = row[i]; });
    return obj;
  });

  var matches = filas.filter(function (r) {
    return normalizar_(r.nombre) === nombreNorm && normalizar_(r.apellido) === apellidoNorm;
  });

  if (matches.length === 0) return null;

  var carpetas = matches.map(function (m) {
    // buscamos el evento de esa carpeta cruzando con la hoja indice
    var enIndice = filasIndice.filter(function (r) { return r.folder_id === m.folder_id; })[0];
    return {
      folder_id: m.folder_id,
      evento: enIndice ? enIndice.evento : '',
      etiqueta: m.etiqueta || (enIndice ? (enIndice.nombre + ' ' + enIndice.apellido) : m.folder_id)
    };
  });

  return { nombre: matches[0].nombre, apellido: matches[0].apellido, carpetas: carpetas };
}

function getIndiceSheet_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  // migracion: si la hoja es de antes de esta funcion, le agregamos las columnas que falten
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  ['email', 'ultimo_aviso'].forEach(function (col) {
    if (headers.indexOf(col) === -1) {
      sheet.getRange(1, headers.length + 1).setValue(col);
      headers.push(col);
    }
  });
  return sheet;
}

function getIndiceRows_() {
  var sheet = getIndiceSheet_();
  var data = sheet.getDataRange().getValues();
  var headers = data.shift();
  return data.map(function (row) {
    var obj = {};
    headers.forEach(function (h, i) { obj[h] = row[i]; });
    return obj;
  });
}

// ---------- ACTIVAR AVISOS POR MAIL (lo hace el propio cliente) ----------
function handleGuardarEmail_(e) {
  var folderId = e.parameter.folder_id;
  var email = (e.parameter.email || '').trim();
  if (!folderId || !email) return jsonOutput_({ error: 'faltan datos' });
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return jsonOutput_({ error: 'email invalido' });

  var sheet = getIndiceSheet_();
  var data = sheet.getDataRange().getValues();
  var headers = data.shift();
  var colFolder = headers.indexOf('folder_id');
  var colEmail = headers.indexOf('email');

  for (var i = 0; i < data.length; i++) {
    if (data[i][colFolder] === folderId) {
      sheet.getRange(i + 2, colEmail + 1).setValue(email);
      return jsonOutput_({ status: 'ok' });
    }
  }
  return jsonOutput_({ error: 'carpeta no encontrada' });
}

// ---------- CHEQUEO AUTOMATICO DE MATERIAL NUEVO (correr con un trigger de tiempo) ----------
// Configurar en Apps Script: Triggers (reloj) > Agregar activador >
// funcion: revisarYNotificar > basado en tiempo > cada 15-30 minutos.
function revisarYNotificar() {
  var rows = getIndiceRows_();
  var ahora = new Date();
  var enviados = 0;

  rows.forEach(function (r) {
    if (!r.activo || !r.email || !r.folder_id) return;

    var ultimoAviso = r.ultimo_aviso ? new Date(r.ultimo_aviso).getTime() : 0;
    var nuevos = 0;

    try {
      var resp = Drive.Files.list({
        q: "'" + r.folder_id + "' in parents and trashed = false",
        fields: 'files(id,createdTime)',
        pageSize: 1000
      });
      (resp.files || []).forEach(function (f) {
        if (new Date(f.createdTime).getTime() > ultimoAviso) nuevos++;
      });
    } catch (err) { return; }

    if (nuevos > 0) {
      var asunto = 'Tapir Media — ' + nuevos + (nuevos === 1 ? ' foto nueva' : ' fotos nuevas') +
                   (r.evento ? ' de ' + r.evento : '');
      var cuerpo = 'Hola ' + r.nombre + ',\n\n' +
                   'Se subieron ' + nuevos + (nuevos === 1 ? ' archivo nuevo' : ' archivos nuevos') +
                   (r.evento ? ' de ' + r.evento : '') + '.\n\n' +
                   'Entrá a tu galería: https://lucasmartineza.github.io/GALERIA-TAPIRMEDIA/\n\n' +
                   '— Tapir Media';
      try {
        MailApp.sendEmail(r.email, asunto, cuerpo);
        actualizarUltimoAviso_(r.folder_id, ahora);
        enviados++;
      } catch (err) {}
    }
  });

  return enviados;
}

// Disparo manual desde admin.html (boton "Notificar material nuevo")
function handleNotificarAhora_(e) {
  if (e.parameter.key !== ADMIN_KEY) return jsonOutput_({ error: 'clave incorrecta' });
  var enviados = revisarYNotificar();
  return jsonOutput_({ status: 'ok', enviados: enviados });
}

function actualizarUltimoAviso_(folderId, fecha) {
  var sheet = getIndiceSheet_();
  var data = sheet.getDataRange().getValues();
  var headers = data.shift();
  var colFolder = headers.indexOf('folder_id');
  var colAviso = headers.indexOf('ultimo_aviso');
  for (var i = 0; i < data.length; i++) {
    if (data[i][colFolder] === folderId) {
      sheet.getRange(i + 2, colAviso + 1).setValue(fecha);
      return;
    }
  }
}

// ---------- LISTADO DE ARCHIVOS ----------
function handleFiles_(e) {
  var folderId = e.parameter.folder;
  if (!folderId) return jsonOutput_({ error: 'falta folder' });

  // compartimos la carpeta entera de una sola vez (mucho mas rapido que
  // archivo por archivo, y los archivos de adentro heredan el permiso)
  try {
    DriveApp.getFolderById(folderId).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err) {}

  var out = [];
  var pageToken = null;

  do {
    var resp = Drive.Files.list({
      q: "'" + folderId + "' in parents and trashed = false",
      fields: 'nextPageToken, files(id,name,mimeType,createdTime,thumbnailLink,size)',
      pageSize: 1000,
      pageToken: pageToken
    });

    (resp.files || []).forEach(function (f) {
      var thumbnailUrl, previewUrl;
      if (f.thumbnailLink) {
        var base = f.thumbnailLink.replace(/=s\d+$/, '');
        thumbnailUrl = base + '=s150';
        previewUrl = base + '=s1600';
      } else {
        thumbnailUrl = 'https://drive.google.com/thumbnail?id=' + f.id + '&sz=w300';
        previewUrl = 'https://drive.google.com/thumbnail?id=' + f.id + '&sz=w1600';
      }
      out.push({
        id: f.id,
        nombre: f.name,
        mimeType: f.mimeType,
        fechaCreacion: new Date(f.createdTime).getTime(),
        peso: f.size ? parseInt(f.size, 10) : 0,
        esVideo: f.mimeType.indexOf('video') === 0,
        thumbnailUrl: thumbnailUrl,
        previewUrl: previewUrl,
        viewUrl: 'https://drive.google.com/uc?export=view&id=' + f.id,
        downloadUrl: 'https://drive.google.com/uc?export=download&id=' + f.id
      });
    });

    pageToken = resp.nextPageToken;
  } while (pageToken);

  // mas nuevo primero
  out.sort(function (a, b) { return b.fechaCreacion - a.fechaCreacion; });

  return jsonOutput_({ status: 'ok', archivos: out });
}

// ---------- DESCARGA TOTAL (ZIP) ----------
function handleZip_(e) {
  var folderId = e.parameter.folder;
  if (!folderId) return jsonOutput_({ error: 'falta folder' });

  var idsParam = e.parameter.ids; // opcional: lista de ids separados por coma, para "descargar seleccionados"
  var blobs = [];
  var nombreZip;

  if (idsParam) {
    var ids = idsParam.split(',').filter(function (s) { return s.trim(); });
    if (ids.length === 0) return jsonOutput_({ error: 'no hay archivos' });
    ids.forEach(function (id) {
      try { blobs.push(DriveApp.getFileById(id.trim()).getBlob()); } catch (err) {}
    });
    nombreZip = 'seleccionados.zip';
  } else {
    var folder = DriveApp.getFolderById(folderId);
    var files = folder.getFiles();
    while (files.hasNext()) blobs.push(files.next().getBlob());
    nombreZip = (folder.getName() || 'material') + '.zip';
  }

  if (blobs.length === 0) return jsonOutput_({ error: 'no hay archivos' });

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
  var avisos = [];
  data.forEach(function (row, i) {
    var obj = {};
    headers.forEach(function (h, j) { obj[h] = row[j]; });
    if (obj.activo) {
      var fechaTexto = obj.fecha instanceof Date
        ? Utilities.formatDate(obj.fecha, Session.getScriptTimeZone(), 'dd/MM HH:mm')
        : String(obj.fecha || '');
      avisos.push({ fila: i + 2, mensaje: obj.mensaje, fecha: fechaTexto });
    }
  });

  return jsonOutput_({ status: 'ok', avisos: avisos });
}

function handleEliminarAviso_(e) {
  if (e.parameter.key !== ADMIN_KEY) return jsonOutput_({ error: 'clave incorrecta' });

  var fila = parseInt(e.parameter.fila, 10);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('avisos');
  if (!sheet || !fila) return jsonOutput_({ error: 'no encontrado' });

  sheet.getRange(fila, 3).setValue(false); // columna "activo"
  return jsonOutput_({ status: 'ok' });
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
    sheet.appendRow(['id', 'lat', 'lng', 'tipo', 'evento', 'activo', 'fecha', 'descripcion']);
    return sheet;
  }
  // migracion: si la hoja ya existia sin la columna "descripcion", se la agregamos
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf('descripcion') === -1) {
    sheet.getRange(1, headers.length + 1).setValue('descripcion');
  }
  return sheet;
}

function handleUbicaciones_(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ubicaciones');
  if (!sheet) return jsonOutput_({ status: 'ok', ubicaciones: [] });
  getOrCrearHojaUbicaciones_(); // asegura que la columna "descripcion" exista

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
      var fechaTexto = u.fecha instanceof Date
        ? Utilities.formatDate(u.fecha, Session.getScriptTimeZone(), 'dd/MM HH:mm')
        : String(u.fecha || '');
      return {
        id: u.id, lat: u.lat, lng: u.lng, tipo: u.tipo,
        evento: u.evento, fecha: fechaTexto, timestamp: u.id,
        descripcion: u.descripcion || ''
      };
    });

  return jsonOutput_({ status: 'ok', ubicaciones: ubicaciones });
}

function handleNuevaUbicacion_(e) {
  if (e.parameter.key !== ADMIN_KEY) return jsonOutput_({ error: 'clave incorrecta' });

  var lat = parseFloat(e.parameter.lat);
  var lng = parseFloat(e.parameter.lng);
  var tipo = e.parameter.tipo;
  if (['foto', 'video', 'ambos'].indexOf(tipo) === -1) tipo = 'foto';
  var evento = (e.parameter.evento || '').trim();
  var descripcion = (e.parameter.descripcion || '').trim();

  if (isNaN(lat) || isNaN(lng) || !evento) return jsonOutput_({ error: 'faltan datos' });

  var sheet = getOrCrearHojaUbicaciones_();
  var id = Date.now();
  var fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM HH:mm');
  sheet.appendRow([id, lat, lng, tipo, evento, true, fecha, descripcion]);

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
