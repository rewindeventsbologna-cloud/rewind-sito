/* ═══════════════════════════════════════════════════════════════
   REWIND — DATI EVENTI
   ---------------------------------------------------------------
   QUESTO È L'UNICO FILE DA MODIFICARE PER AGGIUNGERE UNA FESTA.
   Il sito genera da solo: card in homepage, pagina evento,
   ordinamento, passaggio automatico da "prossimi" a "passati".

   Campi obbligatori: id, title, date, displayDate
   Tutti gli altri campi sono facoltativi: se li lasci vuoti ('')
   semplicemente non vengono mostrati sul sito.
═══════════════════════════════════════════════════════════════ */

const REWIND_EVENTS = [

  // ==============================================================
  // AGGIUNGI QUI I NUOVI EVENTI (copia un blocco e modificalo)
  // ==============================================================

  {
    id:            'temakinho-2026',
    title:         'REWIND TEMAKINHO',
    date:          '2026-09-16',
    displayDate:   '16 Settembre 2026',
    location:      'Temakinho',
    locationDetail:'',                    // ← indirizzo/città da completare
    tags:          ['Sushi', 'Cena'],
    image:         '',                    // ← metti qui la locandina quando è pronta
    customPage:    'temakinho.html',
    description:   '',
    time:          '',                    // ← orario da definire
    dressCode:     '',
    lineup:        '',
    price:         '40 €',
    ingresso:      'Cena + Sushi',
    program:       [],
    registrationEnabled: true,
    galleryUrl:    ''
  },

  {
    // --- identificativo univoco: usato nell'URL e nel Google Sheet ---
    id:            'halloween-2026',

    // --- dati principali ---
    title:         'REWIND HALLOWEEN',
    date:          '2026-10-31',          // formato AAAA-MM-GG (serve al sito per capire se è passato)
    displayDate:   '31 Ottobre 2026',     // come viene scritto sul sito
    location:      'Chiesa sconsacrata',
    locationDetail:'Zola Predosa',

    // --- etichette mostrate sulla card ---
    tags:          ['Halloween Party', 'Free Bar'],

    // --- locandina ---
    // Metti il file nella cartella "locandine/" e scrivi qui il percorso.
    // Se il file non esiste, la card mostra comunque un segnaposto elegante.
    image:         'halloween-2026.jpg',

    // --- pagina dedicata (facoltativo) ---
    // Se valorizzato, la card porta a questa pagina invece che a evento.html.
    // Serve per eventi con grafica speciale. Lascia '' per usare la pagina standard.
    customPage:    'halloween.html',

    // --- testo descrittivo (facoltativo) ---
    description:   '',

    // --- dettagli: lascia '' finché non li hai definiti ---
    time:          '22:00 – 04:00',
    dressCode:     '',   // es. 'Costume obbligatorio'
    lineup:        '',   // es. 'DJ X · DJ Y'
    price:         '25 €',
    ingresso:      'Free Bar',

    // --- programma della serata (facoltativo) ---
    // esempio: [{time:'22:00', label:'Apertura porte'}, {time:'23:00', label:'DJ Set'}]
    program:       [],

    // --- prenotazioni ---
    registrationEnabled: true,   // true = form attivo · false = "prossimamente"

    // --- archivio foto (si usa solo quando l'evento è passato) ---
    galleryUrl:    ''
  },

  // ==============================================================
  // EVENTI PASSATI (archivio)
  // ==============================================================

  {
    id:            'colli-edition-settembre-2026',
    title:         'REWIND — COLLI EDITION',
    date:          '2026-09-12',
    displayDate:   '12 Settembre 2026',
    location:      'Sette Camini Food & Spirits',
    locationDetail:'Bologna — Colli',
    tags:          ['DJ Set', 'Aperitivo'],
    image:         '',
    description:   'Aperitivo, tigelle e DJ Set tra i Colli Bolognesi.',
    time:          '18:00 – 02:00',
    dressCode:     'Elegante',
    lineup:        '',
    price:         '',
    ingresso:      '',
    program:       [],
    registrationEnabled: false,
    galleryUrl:    ''
  },

  {
    id:            'sette-camini-giugno-2026',
    title:         'REWIND — SETTE CAMINI',
    date:          '2026-06-26',
    displayDate:   '26 Giugno 2026',
    location:      'Sette Camini Food & Spirits',
    locationDetail:'Bologna — Colli',
    tags:          ['DJ Set', 'Apericena'],
    image:         '',
    description:   'La nostra apericena più DJ Set tra i colli bolognesi.',
    time:          '',
    dressCode:     '',
    lineup:        '',
    price:         '',
    ingresso:      '',
    program:       [],
    registrationEnabled: false,
    galleryUrl:    'https://drive.google.com/drive/folders/1C9YuYOOGmT0jTzKj4cOZne3SZHaF-AgT'
  },

  {
    id:            'sette-camini-maggio-2026',
    title:         'REWIND — SETTE CAMINI',
    date:          '2026-05-01',
    displayDate:   '01 Maggio 2026',
    location:      'Sette Camini Food & Spirits',
    locationDetail:'Bologna — Colli',
    tags:          ['DJ Set', 'Cena spettacolo'],
    image:         '',
    description:   'La nostra cena spettacolo + dj set tra i colli bolognesi.',
    time:          '',
    dressCode:     '',
    lineup:        '',
    price:         '',
    ingresso:      '',
    program:       [],
    registrationEnabled: false,
    galleryUrl:    'https://drive.google.com/drive/folders/1_V5Bd44NEEdztlp5TiHbCiVljMG6H9bp'
  },

  {
    id:            'accademia-notturni-2025',
    title:         'REWIND — ACCADEMIA DEI NOTTURNI',
    date:          '2025-11-22',
    displayDate:   '22 Novembre 2025',
    location:      'Accademia dei Notturni',
    locationDetail:'Bologna',
    tags:          ['DJ Set'],
    image:         '',
    description:   "Una notte nell'Accademia dei Notturni, atmosfera elegante e raffinata.",
    time:          '',
    dressCode:     '',
    lineup:        '',
    price:         '',
    ingresso:      '',
    program:       [],
    registrationEnabled: false,
    galleryUrl:    'https://drive.google.com/drive/folders/1oXuTJFDiauJlTcERh6mrJmD8TFmJu_FH'
  }

];


/* ═══════════════════════════════════════════════════════════════
   LOGICA DATE — non serve modificarla
   Un evento resta "prossimo" fino alle 06:00 del giorno dopo.
═══════════════════════════════════════════════════════════════ */

function rwEventEnd(ev){
  const d = new Date(ev.date + 'T00:00:00');
  d.setDate(d.getDate() + 1);
  d.setHours(6, 0, 0, 0);
  return d;
}

/** Eventi futuri, dal più vicino al più lontano */
function rwUpcoming(){
  const now = new Date();
  return REWIND_EVENTS
    .filter(e => rwEventEnd(e) > now)
    .sort((a, b) => new Date(a.date) - new Date(b.date));
}

/** Eventi passati, dal più recente al più vecchio */
function rwPast(){
  const now = new Date();
  return REWIND_EVENTS
    .filter(e => rwEventEnd(e) <= now)
    .sort((a, b) => new Date(b.date) - new Date(a.date));
}

/** Trova un evento dal suo id */
function rwEventById(id){
  return REWIND_EVENTS.find(e => e.id === id) || null;
}
