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

const SHEET_NAME        = 'Foglio1';   // foglio predefinito
const CONFIG_SHEET_NAME = 'config';    // pannello di configurazione

/* ═══════════════════════════════════════════════════════════════
   OGNI EVENTO HA IL SUO FOGLIO
   Le prenotazioni non si mescolano mai: ciascun evento scrive
   soltanto nella propria scheda. Per aggiungere un evento futuro
   basta una riga qui sotto.
═══════════════════════════════════════════════════════════════ */
const EVENT_SHEETS = {
  'halloween-2026': 'Foglio1',   // REWIND HALLOWEEN
  'temakinho-2026': 'Foglio2'    // REWIND TEMAKINHO
};

/* Colonne previste. Se mancano vengono create in automatico. */
const HEADERS = ['cognome','nome','email','telefono','evento','eventid','tipologia','data','pagato'];


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

/* ═══════════════════════════════════════════════════════════════
   PANNELLO DI CONFIGURAZIONE
   ---------------------------------------------------------------
   Nel Google Sheet crea (o usa) la scheda "config" con due colonne:
      A = chiave              B = valore
   Le chiavi qui sotto sono quelle riconosciute dal sito.
   Scrivendo un valore nel foglio, il sito lo mostra SUBITO
   senza dover toccare il codice né ricaricare nulla su GitHub.

   ── TEMAKINHO ──────────────────────────────────────────────────
   temakinho.prezzo_cena     40
   temakinho.prezzo_dj       15
   temakinho.orario          20:00 — 02:00
   temakinho.location        Temakinho — Via Farini 13/A, 40124 Bologna
   temakinho.dress           Curato ed elegante
   temakinho.data_breve      16.10.2026
   temakinho.nota_dj         Drink incluso.
   temakinho.mail_orario     22:00 – 02:00
   temakinho.mail_prezzo     15
   temakinho.mail_luogo      TEMAKINHO
   temakinho.mail_data       venerdì 16 ottobre 2026

   ── HALLOWEEN ──────────────────────────────────────────────────
   halloween.badge_prezzo    Early Bird        (etichetta sopra il prezzo;
                                                 lasciala vuota per toglierla)
   halloween.prezzo          25 €
   halloween.orario          22:00 — 04:00
   halloween.location        Chiesa sconsacrata — Zola Predosa
   halloween.dress           (vuoto se non previsto)

   ── CONTATTI (validi per tutte le email) ───────────────────────
   telefono_contatto         3249947917
   nome_dj_contatto          Dilan
═══════════════════════════════════════════════════════════════ */
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

/* Restituisce (creandolo se serve) il foglio dell'evento indicato. */
function getSheetForEvent(eventId) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const nome = EVENT_SHEETS[eventId] || SHEET_NAME;
  let sheet  = ss.getSheetByName(nome);
  if (!sheet) sheet = ss.insertSheet(nome);   // es. "Foglio2" alla prima prenotazione
  ensureHeaders(sheet);
  return sheet;
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  /* Se il foglio "Foglio1" non esiste (es. è stato rinominato), NON ne creiamo
     uno nuovo: userebbe un tab invisibile e le prenotazioni sembrerebbero sparite.
     Si usa invece il primo foglio disponibile che non sia "config". */
  if (!sheet) {
    const all = ss.getSheets();
    for (let i = 0; i < all.length; i++) {
      if (all[i].getName().toLowerCase() !== CONFIG_SHEET_NAME) { sheet = all[i]; break; }
    }
  }
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
    if (action === 'diag')      return buildResp(diagnostica());
    if (action === 'getConfig') return buildResp(getConfig());
    if (action === 'getAll')    return buildResp(getAllRows(e && e.parameter ? e.parameter.eventId : null));
    return buildResp({error: 'Azione non valida: ' + action});
  } catch (err) {
    return buildResp({error: err.message});
  }
}

/* Se passi ?eventId=halloween-2026 restituisce solo quelle prenotazioni */
function getAllRows(eventId) {
  const sheet = eventId ? getSheetForEvent(eventId) : getSheet();
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
  const sheet = getSheetForEvent(body.eventId);   // ← Halloween e Temakinho separati

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
      case 'tipologia': return body.tipologia || '';
      case 'data':      return body.data || new Date().toLocaleString('it-IT');
      case 'pagato':    return 'Non pagato';
      default:          return '';
    }
  });

  /* appendRow scrive SEMPRE dopo l'ultima riga occupata: se in fondo al
     foglio restano celle con spazi o formattazioni, le prenotazioni finivano
     a riga 500. Qui si cerca invece la prima riga davvero vuota. */
  const riga = primaRigaLibera(sheet);
  sheet.getRange(riga, 1, 1, newRow.length).setValues([newRow]);
  Logger.log('Riga aggiunta: ' + riga + ' | foglio: ' + sheet.getName() + ' | evento: ' + (body.eventId || '-'));

  if (body.email) {
    const res = sendEmail(body, cfg);
    Logger.log('Email: ' + JSON.stringify(res));
  }

  return {ok: true, rowIndex: riga};
}

function setPaid(body) {
  const sheet = getSheetForEvent(body.eventId);
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
  const sheet = getSheetForEvent(body.eventId);
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
      subject:  (body.eventId === 'temakinho-2026')
                  ? '✓ Prenotazione confermata — REWIND TEMAKINHO'
                  : ((getTheme(body.eventId || '').subject) || ('✓ Prenotazione confermata — ' + eventName)),
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

/* ═══════════════════════════════════════════════════════════════
   TEMI EMAIL
   ---------------------------------------------------------------
   Ogni evento può avere il proprio stile. Se un evento non è
   elencato qui, viene usato automaticamente il tema "default"
   (classico Rewind). Per dare un tema a una festa futura basta
   aggiungere una voce con la stessa chiave usata come eventId.
═══════════════════════════════════════════════════════════════ */
function getTheme(eventId){
  var themes = {

    // ── Halloween: notturno, gotico, rosso profondo ──
    'halloween-2026': {
      bg:       '#0A0509',
      panel:    '#0A0509',
      headFrom: '#1A0308',
      headTo:   '#4A0A16',
      text:     '#EDE4D8',
      soft:     'rgba(237,228,216,0.78)',
      gold:     '#C9A96E',
      goldl:    '#E3CB99',
      accent:   '#C2334A',
      subject:  'Il tuo nome è sulla lista — REWIND HALLOWEEN',
      claim:    'Sin will find you here',
      opening:  'la tua registrazione a {EVENTO} è stata ricevuta correttamente.',
      opening2: 'Conserveremo la tua prenotazione associata a questo indirizzo email.',
      note:     'Seguiranno eventuali ulteriori informazioni relative alla serata.',
      rows:     ['data','luogo'],          // solo queste due righe
      contact:  'Per informazioni:',
      sign:     'Memories worth rewinding'
    }

  };
  var def = {
    bg:'#1E0009', panel:'#1E0009', headFrom:'#8B0022', headTo:'#C8102E',
    text:'#F5EFE6', soft:'rgba(245,239,230,0.82)',
    gold:'#C9A96E', goldl:'#E3CB99', accent:'#C8102E',
    subject:'', claim:'',
    opening:'la tua registrazione a {EVENTO} è stata ricevuta correttamente.',
    opening2:'Conserveremo la tua prenotazione associata a questo indirizzo email.',
    note:'Seguiranno eventuali ulteriori informazioni relative alla serata.',
    rows:['data','luogo','orario','formula','ingresso'],
    contact:'Per informazioni:',
    sign:'Memories worth rewinding'
  };
  var t = themes[eventId] || {};
  for (var k in def) { if (t[k] === undefined) t[k] = def[k]; }
  return t;
}

/* Template unico: colori e testi cambiano in base al tema dell'evento */
function buildEmailHtml(body, cfg) {
  /* Ogni evento può avere un template tutto suo.
     Temakinho NON usa il template di Halloween e viceversa. */
  if ((body.eventId || '') === 'temakinho-2026') return buildEmailTemakinho(body, cfg);

  var nome      = body.nome || 'Ospite';
  var eventName = body.eventName     || 'Rewind';
  var eventDate = body.eventDate     || '';
  var eventLoc  = body.eventLocation || '';
  var eventTime = body.eventTime     || '';
  var eventPrice= body.eventPrice    || '';
  var eventForm = body.eventFormat   || '';
  var t = getTheme(body.eventId || '');

  function row(label, value){
    if(!value) return '';
    return '<tr>' +
      '<td style="padding:9px 0;border-bottom:1px solid rgba(201,169,110,0.14);' +
      'font-family:Arial,sans-serif;font-size:10px;letter-spacing:2px;text-transform:uppercase;' +
      'color:' + t.gold + ';width:110px;vertical-align:top;">' + label + '</td>' +
      '<td style="padding:9px 0;border-bottom:1px solid rgba(201,169,110,0.14);' +
      'font-size:15px;color:' + t.text + ';"><strong>' + value + '</strong></td></tr>';
  }

  return '' +
  '<div style="margin:0;padding:24px 12px;background:' + t.bg + ';">' +
  '<div style="font-family:Georgia,serif;max-width:600px;margin:0 auto;background:' + t.panel + ';' +
       'border:1px solid rgba(201,169,110,0.22);">' +

    // intestazione
    '<div style="background:linear-gradient(135deg,' + t.headFrom + ' 0%,' + t.headTo + ' 100%);' +
         'padding:42px 36px;text-align:center;border-bottom:1px solid rgba(201,169,110,0.30);">' +
      '<div style="font-family:Arial,sans-serif;letter-spacing:9px;font-size:13px;' +
           'color:' + t.goldl + ';margin-bottom:12px;">REWIND</div>' +
      '<div style="font-family:Georgia,serif;font-size:30px;line-height:1.1;color:' + t.text + ';">' +
        (eventName.replace(/^REWIND\s*[\u2014\u2013-]?\s*/i,'') || eventName) + '</div>' +
      (t.claim ? '<div style="font-style:italic;font-size:15px;color:' + t.accent +
                 ';margin-top:10px;">' + t.claim + '</div>' : '') +
      (eventDate ? '<div style="font-family:Arial,sans-serif;font-size:11px;letter-spacing:3px;' +
                   'text-transform:uppercase;color:rgba(237,228,216,0.62);margin-top:14px;">' +
                   eventDate + '</div>' : '') +
    '</div>' +

    // corpo
    '<div style="padding:40px 36px;">' +
      '<p style="font-size:16px;line-height:1.8;margin:0 0 14px;color:' + t.text + ';">Ciao <strong>' + nome + '</strong>,</p>' +
      '<p style="font-size:15px;line-height:1.9;margin:0 0 14px;color:' + t.soft + ';">' +
        String(t.opening).replace('{EVENTO}', '<strong style="color:' + t.text + ';">' + eventName + '</strong>') + '</p>' +
      (t.opening2 ? '<p style="font-size:15px;line-height:1.9;margin:0 0 28px;color:' + t.soft + ';">' + t.opening2 + '</p>' : '') +

      '<table style="width:100%;border-collapse:collapse;margin-bottom:26px;">' +
        (t.rows.indexOf('data')     >= 0 ? row('Data',     eventDate)  : '') +
        (t.rows.indexOf('luogo')    >= 0 ? row('Location', eventLoc)   : '') +
        (t.rows.indexOf('orario')   >= 0 ? row('Orario',   eventTime)  : '') +
        (t.rows.indexOf('formula')  >= 0 ? row('Formula',  eventForm)  : '') +
        (t.rows.indexOf('ingresso') >= 0 ? row('Ingresso', eventPrice) : '') +
      '</table>' +

      '<div style="border-left:2px solid ' + t.accent + ';padding:14px 18px;margin-bottom:28px;' +
           'background:rgba(201,169,110,0.06);">' +
        '<p style="margin:0;font-size:14px;line-height:1.8;color:' + t.soft + ';">' + t.note + '</p>' +
      '</div>' +

      '<p style="font-size:14px;line-height:1.8;color:' + t.soft + ';margin:0 0 6px;">' +
        t.contact + ' <strong style="color:' + t.text + ';">' +
        cfg.telefono_contatto + '</strong> ' + cfg.nome_dj_contatto + '</p>' +
      '<p style="font-family:Arial,sans-serif;font-size:11px;letter-spacing:3px;' +
         'text-transform:uppercase;color:rgba(237,228,216,0.42);margin:22px 0 0;">Staff Rewind</p>' +
    '</div>' +

    // chiusura
    '<div style="border-top:1px solid rgba(201,169,110,0.18);padding:20px 36px;text-align:center;">' +
      '<p style="margin:0;font-style:italic;font-size:13px;color:' + t.gold + ';">' + t.sign + '</p>' +
      '<p style="margin:8px 0 0;font-family:Arial,sans-serif;font-size:10px;letter-spacing:2px;' +
         'text-transform:uppercase;color:rgba(237,228,216,0.35);">Rewind Private Club · Bologna</p>' +
    '</div>' +

  '</div></div>';
}



/* Prima riga senza alcun contenuto reale (ignora spazi vuoti). */
function primaRigaLibera(sheet) {
  var last = sheet.getLastRow();
  var cols = Math.max(sheet.getLastColumn(), 1);
  if (last < 1) return 2;                     // foglio vuoto: subito sotto le intestazioni
  var vals = sheet.getRange(1, 1, last, cols).getValues();
  for (var i = 1; i < vals.length; i++) {     // si parte da 1: la riga 1 è l'intestazione
    if (vals[i].join('').trim() === '') return i + 1;
  }
  return last + 1;
}

/* ═══════════════════════════════════════════════════════════════
   ► PULISCI FOGLIO ◄
   Seleziona questa funzione nel menu in alto e premi Esegui.
   Elimina le righe vuote rimaste in fondo (quelle che facevano
   finire le prenotazioni a riga 500) e compatta i dati in alto.
   I dati veri NON vengono toccati.
═══════════════════════════════════════════════════════════════ */
function pulisciFoglio() {
  var sheet = getSheet();
  var last  = sheet.getLastRow();
  var cols  = Math.max(sheet.getLastColumn(), 1);
  if (last < 2) { Logger.log('Niente da pulire.'); return; }

  var vals  = sheet.getRange(2, 1, last - 1, cols).getValues();
  var piene = vals.filter(function(r){ return r.join('').trim() !== ''; });

  Logger.log('Righe con dati: ' + piene.length + '  ·  righe occupate prima: ' + (last - 1));

  // riscrive i dati compattati subito sotto l'intestazione
  sheet.getRange(2, 1, last - 1, cols).clearContent();
  if (piene.length) sheet.getRange(2, 1, piene.length, cols).setValues(piene);

  // elimina le righe in eccesso rimaste sotto
  var ultima = piene.length + 1;
  if (sheet.getMaxRows() > ultima + 50) {
    sheet.deleteRows(ultima + 51, sheet.getMaxRows() - ultima - 50);
  }
  Logger.log('Fatto: ora i dati partono da riga 2 e arrivano a riga ' + ultima + '.');
}

/* ═══════════════════════════════════════════════════════════════
   ► VERIFICA FOGLIO  ◄
   Seleziona questa funzione nel menu in alto e premi Esegui.
   Poi apri "Log esecuzione": ti dice ESATTAMENTE su quale file
   Google Sheet e su quale scheda vengono scritte le prenotazioni.
   Non serve distribuire nulla.
═══════════════════════════════════════════════════════════════ */
function verificaFoglio() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    Logger.log('ATTENZIONE: questo script NON è collegato a nessun foglio Google.');
    Logger.log('Va aperto da Foglio -> Estensioni -> Apps Script, oppure collegato al file giusto.');
    return;
  }

  Logger.log('════════════════════════════════════════');
  Logger.log('FILE USATO DALLO SCRIPT: ' + ss.getName());
  Logger.log('INDIRIZZO DEL FILE     : ' + ss.getUrl());
  Logger.log('(apri questo indirizzo: è QUI che finiscono le prenotazioni)');
  Logger.log('════════════════════════════════════════');

  var fogli = ss.getSheets();
  Logger.log('SCHEDE PRESENTI NEL FILE (' + fogli.length + '):');
  for (var i = 0; i < fogli.length; i++) {
    var f = fogli[i];
    Logger.log('   • "' + f.getName() + '"  →  righe con dati: ' + Math.max(f.getLastRow() - 1, 0));
  }

  var sheet = getSheet();
  Logger.log('════════════════════════════════════════');
  Logger.log('SCHEDA DOVE SCRIVE: "' + sheet.getName() + '"');

  var cols = sheet.getLastColumn();
  var last = sheet.getLastRow();
  if (cols > 0) Logger.log('COLONNE: ' + sheet.getRange(1,1,1,cols).getValues()[0].join(' | '));
  Logger.log('PRENOTAZIONI SALVATE: ' + Math.max(last - 1, 0));

  if (last > 1) {
    var n = Math.min(5, last - 1);
    var righe = sheet.getRange(last - n + 1, 1, n, cols).getValues();
    Logger.log('ULTIME ' + n + ' RIGHE:');
    for (var j = 0; j < righe.length; j++) Logger.log('   ' + righe[j].join(' | '));
  } else {
    Logger.log('Nessuna riga presente in questa scheda.');
  }
  Logger.log('════════════════════════════════════════');
}

/* ═══════════════════════════════════════════════════════════════
   DIAGNOSTICA
   Apri l'URL del web app aggiungendo  ?action=diag
   Dice su QUALE foglio vengono scritte le prenotazioni e mostra
   le ultime righe salvate. Serve per capire dove finiscono i dati.
═══════════════════════════════════════════════════════════════ */
function diagnostica() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheet();
  const last  = sheet.getLastRow();
  const cols  = sheet.getLastColumn();

  const headers = cols > 0
    ? sheet.getRange(1, 1, 1, cols).getValues()[0]
    : [];

  let ultime = [];
  if (last > 1) {
    const n = Math.min(3, last - 1);
    ultime = sheet.getRange(last - n + 1, 1, n, cols).getValues();
  }

  return {
    ok:                true,
    fileGoogleSheet:   ss.getName(),
    urlFoglio:         ss.getUrl(),
    fogliPresenti:     ss.getSheets().map(function(s){ return s.getName(); }),
    foglioUsato:       sheet.getName(),
    intestazioni:      headers,
    righeTotali:       Math.max(last - 1, 0),
    ultimeRighe:       ultime,
    emailProprietario: Session.getEffectiveUser().getEmail()
  };
}

/* Test manuale: premi ▶ Esegui su questa funzione per inserire
   una riga di prova e verificare che la scrittura funzioni. */
function testInserimento() {
  const r = addRow({
    nome:'TEST', cognome:'PROVA', email:'test@example.com', telefono:'3330000000',
    eventId:'test', eventName:'RIGA DI PROVA', eventDate:'—', eventLocation:'—'
  });
  Logger.log(JSON.stringify(r));
}

/* ═══════════════════════════════════════════════════════════════
   EMAIL — REWIND TEMAKINHO (DJ Set)
   Template dedicato: colori del locale (smeraldo, crema, ottone).
   I testi si possono cambiare dal foglio "config" con le chiavi
   temakinho.mail_*  senza toccare il codice.
═══════════════════════════════════════════════════════════════ */
function buildEmailTemakinho(body, cfg) {
  var nome = body.nome || 'Ospite';

  var C = {
    bg:'#071410', head1:'#0E2A20', head2:'#14503C',
    text:'#F4EFE3', soft:'rgba(244,239,227,0.80)',
    brass:'#C9A55C', brassl:'#E4CE96', line:'rgba(201,165,92,0.20)'
  };

  var v = function(chiave, predefinito){
    return (cfg[chiave] !== undefined && cfg[chiave] !== '') ? cfg[chiave] : predefinito;
  };

  var orario   = v('temakinho.mail_orario',   '22:00 – 02:00');
  var prezzo   = v('temakinho.mail_prezzo',   '15');
  var luogo    = v('temakinho.mail_luogo',    'TEMAKINHO');
  var dataEv   = v('temakinho.mail_data',     'venerdì 16 ottobre 2026');
  var contatto = v('telefono_contatto',       '3249947917');
  var referente= v('nome_dj_contatto',        'Dilan');

  return '' +
  '<div style="margin:0;padding:24px 12px;background:' + C.bg + ';">' +
  '<div style="font-family:Georgia,serif;max-width:600px;margin:0 auto;background:' + C.bg + ';' +
       'border:1px solid ' + C.line + ';">' +

    '<div style="background:linear-gradient(135deg,' + C.head1 + ' 0%,' + C.head2 + ' 100%);' +
         'padding:42px 36px;text-align:center;border-bottom:1px solid rgba(201,165,92,0.30);">' +
      '<div style="font-family:Arial,sans-serif;letter-spacing:9px;font-size:13px;color:' + C.brassl + ';margin-bottom:12px;">REWIND</div>' +
      '<div style="font-family:Georgia,serif;font-size:30px;line-height:1.1;color:' + C.text + ';">TEMAKINHO</div>' +
      '<div style="font-style:italic;font-size:15px;color:' + C.brassl + ';margin-top:10px;">Cena spettacolo &amp; DJ Set</div>' +
      '<div style="font-family:Arial,sans-serif;font-size:11px;letter-spacing:3px;text-transform:uppercase;' +
           'color:rgba(244,239,227,0.62);margin-top:14px;">16 Ottobre 2026</div>' +
    '</div>' +

    '<div style="padding:40px 36px;">' +
      '<p style="font-size:16px;line-height:1.8;margin:0 0 14px;color:' + C.text + ';">Ciao <strong>' + nome + '</strong>,</p>' +
      '<p style="font-size:15px;line-height:1.9;margin:0 0 10px;color:' + C.soft + ';">' +
        'grazie per esserti registrato/a a <strong style="color:' + C.text + ';">Rewind Private Club TEMAKINHO</strong>.</p>' +
      '<p style="font-size:15px;line-height:1.9;margin:0 0 28px;color:' + C.soft + ';">' +
        'La tua prenotazione è stata registrata con successo.</p>' +

      '<p style="font-family:Arial,sans-serif;font-size:11px;letter-spacing:2px;text-transform:uppercase;' +
         'color:' + C.brass + ';margin:0 0 12px;">Programma della serata</p>' +
      '<div style="border-left:2px solid ' + C.brass + ';padding:14px 18px;margin-bottom:26px;' +
           'background:rgba(201,165,92,0.06);">' +
        '<p style="margin:0 0 6px;font-size:15px;color:' + C.text + ';"><strong>' + orario + ' · DJ Set</strong></p>' +
        '<p style="margin:0;font-size:14px;line-height:1.7;color:' + C.soft + ';">DJ Set con un drink incluso.</p>' +
      '</div>' +

      '<p style="font-size:15px;line-height:1.9;margin:0 0 26px;color:' + C.soft + ';">' +
        'Ti aspettiamo <strong style="color:' + C.text + ';">' + dataEv + '</strong> a partire dalle ' +
        '<strong style="color:' + C.text + ';">22:00</strong> presso <strong style="color:' + C.text + ';">' + luogo + '</strong>.</p>' +

      '<p style="font-family:Arial,sans-serif;font-size:11px;letter-spacing:2px;text-transform:uppercase;' +
         'color:' + C.brass + ';margin:0 0 10px;">Pagamento</p>' +
      '<div style="border:1px solid rgba(201,165,92,0.30);background:rgba(201,165,92,0.07);' +
           'padding:18px 20px;margin-bottom:26px;text-align:center;">' +
        '<p style="margin:0;font-size:15px;line-height:1.8;color:' + C.text + ';">' +
          'Il pagamento di <strong>€' + prezzo + '</strong> dovrà essere effettuato ' +
          '<strong>direttamente in cassa</strong> la sera dell\'evento.</p>' +
      '</div>' +

      '<p style="font-family:Arial,sans-serif;font-size:11px;letter-spacing:2px;text-transform:uppercase;' +
         'color:' + C.brass + ';margin:0 0 8px;">Informazione importante</p>' +
      '<p style="font-size:14px;line-height:1.8;color:' + C.soft + ';margin:0 0 26px;">' +
        'Per qualsiasi informazione contattare <strong style="color:' + C.text + ';">' + contatto + '</strong> ' + referente + '</p>' +

      '<p style="font-family:Arial,sans-serif;font-size:11px;letter-spacing:3px;text-transform:uppercase;' +
         'color:rgba(244,239,227,0.45);margin:0;">Staff Rewind</p>' +
    '</div>' +

    '<div style="border-top:1px solid ' + C.line + ';padding:20px 36px;text-align:center;">' +
      '<p style="margin:0;font-style:italic;font-size:13px;color:' + C.brass + ';">Memories worth rewinding</p>' +
      '<p style="margin:8px 0 0;font-family:Arial,sans-serif;font-size:10px;letter-spacing:2px;' +
         'text-transform:uppercase;color:rgba(244,239,227,0.35);">Rewind Private Club · Bologna</p>' +
    '</div>' +

  '</div></div>';
}

function buildResp(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
