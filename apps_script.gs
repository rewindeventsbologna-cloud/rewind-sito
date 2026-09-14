/* ═══════════════════════════════════════════════════════════════
   REWIND — Google Apps Script v10 (MULTI-EVENTO)
   ---------------------------------------------------------------
   • Salva le prenotazioni nel foglio "Foglio1"
   • Ogni riga contiene l'evento a cui si riferisce (colonne evento/eventid)
   • Invia una email di conferma DINAMICA in base all'evento
   ---------------------------------------------------------------
   NON serve modificare questo file quando aggiungi un nuovo evento:
   i dati dell'evento arrivano automaticamente dal sito.
═══════════════════════════════════════════════════════════════ */

const SHEET_NAME        = 'Foglio1';   // foglio prenotazioni
const CONFIG_SHEET_NAME = 'config';    // foglio configurazione (facoltativo)

/* Colonne previste. Se mancano vengono create in automatico. */
const HEADERS = ['cognome','nome','email','telefono','evento','eventid','data','pagato'];


/* ═══════════════ CONFIG ═══════════════ */

function getConfig() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  const cfg = {};
  if (sheet) {
    sheet.getDataRange().getValues().forEach(r => {
      if (r[0] && String(r[0]).trim()) cfg[String(r[0]).trim()] = r[1];
    });
  }
  const def = getDefaultConfig();
  Object.keys(def).forEach(k => { if (cfg[k] === undefined || cfg[k] === '') cfg[k] = def[k]; });
  return cfg;
}

function getDefaultConfig() {
  return {
    nome_brand:        'Rewind Private Club',
    whatsapp:          '393920559059',
    email_contatto:    'rewindeventsbologna@gmail.com',
    nome_dj_contatto:  'Dilan',
    telefono_contatto: '3249947917',
    instagram:         'https://www.instagram.com/rewind.eventss'
  };
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  ensureHeaders(sheet);
  return sheet;
}

/* Crea le intestazioni mancanti senza toccare i dati esistenti */
function ensureHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    return;
  }
  const lastCol  = Math.max(sheet.getLastColumn(), 1);
  const current  = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
                     .map(h => String(h).toLowerCase().trim());
  HEADERS.forEach(h => {
    if (current.indexOf(h) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(h);
    }
  });
}


/* ═══════════════ GET ═══════════════ */

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'getAll';
    if (action === 'getConfig') return buildResp(getConfig());
    if (action === 'getAll')    return buildResp(getAllRows(e && e.parameter ? e.parameter.eventId : null));
    return buildResp({error: 'Azione non valida: ' + action});
  } catch (err) {
    return buildResp({error: err.message});
  }
}

/* Se passi ?eventId=halloween-2026 restituisce solo quelle prenotazioni */
function getAllRows(eventId) {
  const sheet = getSheet();
  const last  = sheet.getLastRow();
  if (last < 2) return {data: []};

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                    .map(h => String(h).toLowerCase().trim());
  const vals = sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
  const idx  = name => headers.indexOf(name);

  const data = [];
  vals.forEach((r, i) => {
    if (!r.join('').trim()) return;
    const get = name => idx(name) >= 0 ? String(r[idx(name)] || '') : '';
    const row = {
      rowIndex: i + 2,
      cognome:  get('cognome'),
      nome:     get('nome'),
      email:    get('email'),
      telefono: get('telefono'),
      evento:   get('evento'),
      eventId:  get('eventid'),
      data:     get('data'),
      pagato:   get('pagato') || 'Non pagato'
    };
    if (!eventId || row.eventId === eventId) data.push(row);
  });
  return {data: data};
}


/* ═══════════════ POST ═══════════════ */

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    Logger.log('doPost: ' + action + ' | ' + (body.email || '') + ' | ' + (body.eventId || ''));
    if (action === 'addRow')    return buildResp(addRow(body));
    if (action === 'setPaid')   return buildResp(setPaid(body));
    if (action === 'deleteRow') return buildResp(deleteRow(body));
    return buildResp({error: 'Azione non valida: ' + action});
  } catch (err) {
    Logger.log('doPost ERROR: ' + err.message);
    return buildResp({error: err.message});
  }
}

function addRow(body) {
  const cfg   = getConfig();
  const sheet = getSheet();

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                    .map(h => String(h).toLowerCase().trim());

  const newRow = headers.map(h => {
    switch (h) {
      case 'cognome':   return body.cognome  || '';
      case 'nome':      return body.nome     || '';
      case 'email':     return body.email    || '';
      case 'telefono':
      case 'tel':       return body.telefono || '';
      case 'evento':    return body.eventName || '';
      case 'eventid':   return body.eventId   || '';
      case 'data':      return body.data || new Date().toLocaleString('it-IT');
      case 'pagato':    return 'Non pagato';
      default:          return '';
    }
  });

  sheet.appendRow(newRow);
  Logger.log('Riga aggiunta: ' + sheet.getLastRow() + ' | evento: ' + (body.eventId || '-'));

  if (body.email) {
    const res = sendEmail(body, cfg);
    Logger.log('Email: ' + JSON.stringify(res));
  }

  return {ok: true, rowIndex: sheet.getLastRow()};
}

function setPaid(body) {
  const sheet = getSheet();
  const rowIndex = Number(body.rowIndex);
  if (rowIndex < 2 || rowIndex > sheet.getLastRow()) return {error: 'Riga non valida'};
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                    .map(h => String(h).toLowerCase().trim());
  let col = headers.indexOf('pagato') + 1;
  if (col === 0) col = headers.length;
  sheet.getRange(rowIndex, col).setValue(String(body.pagato));
  return {ok: true, updated: rowIndex};
}

function deleteRow(body) {
  const sheet = getSheet();
  const rowIndex = Number(body.rowIndex);
  if (rowIndex < 2 || rowIndex > sheet.getLastRow()) return {error: 'Riga non valida'};
  sheet.deleteRow(rowIndex);
  return {ok: true, deleted: rowIndex};
}


/* ═══════════════ EMAIL DINAMICA ═══════════════ */

function sendEmail(body, cfg) {
  const nome      = body.nome || 'Ospite';
  const email     = body.email;
  const eventName = body.eventName || 'Rewind';
  if (!email) return {error: 'Email mancante'};

  try {
    MailApp.sendEmail({
      to:       email,
      subject:  '✓ Prenotazione confermata — ' + eventName,
      htmlBody: buildEmailHtml(body, cfg),
      name:     cfg.nome_brand || 'Rewind Private Club',
      replyTo:  cfg.email_contatto
    });
    return {ok: true, sent: email};
  } catch (err) {
    Logger.log('Email error: ' + err.message);
    return {error: err.message};
  }
}

/* Template unico, si adatta da solo all'evento ricevuto dal sito */
function buildEmailHtml(body, cfg) {
  const nome      = body.nome || 'Ospite';
  const eventName = body.eventName     || 'Rewind';
  const eventDate = body.eventDate     || '';
  const eventLoc  = body.eventLocation || '';

  const righe =
    (eventDate ? '<tr><td style="padding:6px 0;font-size:13px;color:#C9A96E;letter-spacing:1px;text-transform:uppercase;width:90px;">Data</td><td style="padding:6px 0;font-size:15px;color:#F5EFE6;"><strong>' + eventDate + '</strong></td></tr>' : '') +
    (eventLoc  ? '<tr><td style="padding:6px 0;font-size:13px;color:#C9A96E;letter-spacing:1px;text-transform:uppercase;">Location</td><td style="padding:6px 0;font-size:15px;color:#F5EFE6;"><strong>' + eventLoc + '</strong></td></tr>' : '');

  return '<div style="font-family:Georgia,serif;max-width:600px;margin:0 auto;background:#1E0009;color:#F5EFE6;">' +
    '<div style="background:linear-gradient(135deg,#8B0022 0%,#C8102E 100%);padding:40px;text-align:center;border-bottom:1px solid rgba(201,169,110,0.3);">' +
      '<h1 style="font-family:Arial,sans-serif;letter-spacing:.3em;font-size:2rem;margin:0;color:#F5EFE6;">REWIND</h1>' +
      '<p style="font-style:italic;color:rgba(245,239,230,.85);margin:10px 0 0;">' + eventName + (eventDate ? ' · ' + eventDate : '') + '</p>' +
    '</div>' +
    '<div style="padding:44px 36px;background:#1E0009;">' +
      '<p style="font-size:15px;line-height:1.9;margin:0 0 8px;">Ciao <strong>' + nome + '</strong>,</p>' +
      '<p style="font-size:15px;color:rgba(245,239,230,.82);margin:0 0 6px;">la tua registrazione a <strong>' + eventName + '</strong> è stata <strong style="color:#C9A96E;">ricevuta correttamente</strong>.</p>' +
      '<p style="font-size:15px;color:rgba(245,239,230,.82);margin:0 0 28px;">Conserveremo la tua prenotazione associata a questo indirizzo email.</p>' +
      (righe ? '<hr style="border:none;border-top:1px solid rgba(245,239,230,0.1);margin-bottom:20px;">' +
               '<table style="width:100%;border-collapse:collapse;margin-bottom:28px;">' + righe + '</table>' : '') +
      '<div style="background:rgba(201,169,110,.08);border-left:3px solid #C9A96E;padding:16px 20px;margin-bottom:26px;">' +
        '<p style="font-size:13px;line-height:1.8;color:rgba(245,239,230,.88);margin:0;">' +
        'Seguiranno eventuali ulteriori informazioni relative alla serata.</p>' +
      '</div>' +
      '<hr style="border:none;border-top:1px solid rgba(245,239,230,0.1);margin-bottom:22px;">' +
      '<p style="font-size:14px;color:rgba(245,239,230,.75);margin:0 0 8px;">Per informazioni: <strong style="color:#F5EFE6;">' + cfg.telefono_contatto + '</strong> ' + cfg.nome_dj_contatto + '</p>' +
      '<p style="font-size:13px;letter-spacing:3px;color:rgba(245,239,230,.4);margin:0;">Staff Rewind.</p>' +
    '</div>' +
    '<div style="background:#150007;border-top:1px solid rgba(201,169,110,.15);padding:20px 36px;text-align:center;">' +
      '<p style="font-size:11px;color:rgba(245,239,230,.3);letter-spacing:.2em;text-transform:uppercase;margin:0;">Rewind Private Club · Bologna</p>' +
    '</div></div>';
}

function buildResp(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
