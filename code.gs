/**
 * =========================================================================
 * 🏠 BACKEND SCRIPT – Rekap Sewa & Pajak
 * Divisi Riayah - Masjid Munzalan Mubarakan
 * =========================================================================
 */

// ── Sheet names ──────────────────────────────────────────────
var SHEET_PAJAK  = 'DataPajak';
var SHEET_MOTOR  = 'DataPajakMotor';
var SHEET_SEWA   = 'DataSewa';
var SHEET_ADMIN  = 'AdminSettings';

// ── Spreadsheet ID (Kosongkan jika Apps Script dibuka dari Google Sheet) ────────
var SPREADSHEET_ID = '1y14QaNlv0NQDIz4F3qagDP9ijQdXlbRqk7MoWhgYiYg';

// ── Column headers ────────────────────────────────────────────
var PAJAK_HEADERS = [
  'ID','Nama Unit','No Polisi','Atas Nama','Warna',
  'Tgl Bayar Pajak','Nominal Pajak','Jatuh Tempo Pajak',
  'Tgl Bayar STNK','Nominal STNK','Jatuh Tempo STNK',
  'Keterangan','Created At','Updated At'
];
var SEWA_HEADERS = [
  'ID','Nama Rumah','Nama Penyewa','Nama Pemilik',
  'Keterangan Sewa','Biaya Sewa','Tgl Mulai','Tgl Berakhir',
  'Link MOU','Link Kondisi Awal','Catatan','Created At','Updated At'
];

/* ─────────────────────────────────────────────────────────────
   getSpreadsheet – Helper untuk membuka Spreadsheet otomatis
───────────────────────────────────────────────────────────── */
function getSpreadsheet() {
  if (typeof SPREADSHEET_ID !== 'undefined' && SPREADSHEET_ID && SPREADSHEET_ID !== 'GANTI_DENGAN_ID_SPREADSHEET_ANDA') {
    try { return SpreadsheetApp.openById(SPREADSHEET_ID); } catch(e) {}
  }
  try {
    var activeSs = SpreadsheetApp.getActiveSpreadsheet();
    if (activeSs) return activeSs;
  } catch(e) {}
  throw new Error('ID Spreadsheet belum diisi! Silakan ganti SPREADSHEET_ID di baris 14 code.gs dengan ID Spreadsheet Anda.');
}

/* ─────────────────────────────────────────────────────────────
   doGet – Handles Web App UI & API GET Requests
───────────────────────────────────────────────────────────── */
function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    var action = e.parameter.action;
    var result = { success: false, message: 'Action tidak ditemukan' };
    try {
      if (action === 'getAllData') {
        result = getAllData();
      } else if (action === 'saveRecord' && e.parameter.type && e.parameter.data) {
        result = saveRecord(e.parameter.type, e.parameter.data);
      } else if (action === 'deleteRecord' && e.parameter.type && e.parameter.id) {
        result = deleteRecord(e.parameter.type, e.parameter.id);
      } else if (action === 'saveAdminSettings' && e.parameter.data) {
        result = saveAdminSettings(JSON.parse(e.parameter.data));
      }
    } catch(err) {
      result = { success: false, message: err.toString() };
    }
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }
  try {
    var ss = getSpreadsheet();
    ensureSheetsExist(ss);
  } catch(err) {}

  try {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Riayah Center – Rekapitulasi Pembayaran Sewa Rumah dan Pajak Kendaraan')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'active',
      message: 'Riayah Center API is running successfully!'
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/* ─────────────────────────────────────────────────────────────
   doPost – Handles API Requests from External Frontends (Vercel)
───────────────────────────────────────────────────────────── */
function doPost(e) {
  try {
    var contents = e.postData ? e.postData.contents : '';
    var request = JSON.parse(contents);
    var action = request.action;
    var result = { success: false, message: 'Action tidak ditemukan' };

    if (action === 'getAllData') {
      result = getAllData();
    } else if (action === 'saveAdminSettings') {
      result = saveAdminSettings(request.data);
    } else if (action === 'sendTestReminderEmail') {
      result = sendTestReminderEmail(request.data);
    } else if (action === 'saveRecord') {
      result = saveRecord(request.type, JSON.stringify(request.data));
    } else if (action === 'deleteRecord') {
      result = deleteRecord(request.type, request.id);
    } else if (action === 'uploadFile') {
      result = uploadFileToDrive(request.data);
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/* ─────────────────────────────────────────────────────────────
   ensureSheetsExist
───────────────────────────────────────────────────────────── */
function ensureSheetsExist(ss) {
  if (!ss) ss = getSpreadsheet(); // Fallback: ambil otomatis jika dipanggil tanpa argumen
  var sheetDefs = [
    { name: SHEET_PAJAK,  headers: PAJAK_HEADERS  },
    { name: SHEET_MOTOR,  headers: PAJAK_HEADERS  },
    { name: SHEET_SEWA,   headers: SEWA_HEADERS   },
    { name: SHEET_ADMIN,  headers: ['Key','Value'] }
  ];
  sheetDefs.forEach(function(def) {
    var sh = ss.getSheetByName(def.name);
    if (!sh) {
      sh = ss.insertSheet(def.name);
      sh.appendRow(def.headers);
      sh.getRange(1, 1, 1, def.headers.length)
        .setBackground('#1a2e1e')
        .setFontColor('#7ecf65')
        .setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  });
}

/* ─────────────────────────────────────────────────────────────
   getAllData – Called from frontend on load
───────────────────────────────────────────────────────────── */
function getAllData() {
  try {
    var ss = getSpreadsheet();
    return {
      success: true,
      pajak: getPajakList(ss),
      motor: getMotorList(ss),
      sewa:  getSewaList(ss),
      admin: getAdminConfig(ss)
    };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

/* ─────────────────────────────────────────────────────────────
   getPajakList
───────────────────────────────────────────────────────────── */
function getPajakList(ss) {
  if (!ss) ss = getSpreadsheet(); // Fallback jika dipanggil langsung dari editor
  var sh   = ss.getSheetByName(SHEET_PAJAK);
  if (!sh) return [];
  var data = sh.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(function(r) {
    return {
      id:             r[0],
      namaUnit:       r[1],
      nopol:          r[2],
      atasNama:       r[3],
      warna:          r[4],
      tglBayar:       r[5] ? Utilities.formatDate(new Date(r[5]), 'Asia/Jakarta', 'yyyy-MM-dd') : '',
      nominal:        r[6],
      jatuhTempo:     r[7] ? Utilities.formatDate(new Date(r[7]), 'Asia/Jakarta', 'yyyy-MM-dd') : '',
      tglStnk:        r[8] ? Utilities.formatDate(new Date(r[8]), 'Asia/Jakarta', 'yyyy-MM-dd') : '',
      nominalStnk:    r[9],
      jatuhTempoStnk: r[10] ? Utilities.formatDate(new Date(r[10]), 'Asia/Jakarta', 'yyyy-MM-dd') : '',
      keterangan:     r[11],
      createdAt:      r[12],
      updatedAt:      r[13]
    };
  }).filter(function(r){ return r.id; });
}

/* ─────────────────────────────────────────────────────────────
   getMotorList
───────────────────────────────────────────────────────────── */
function getMotorList(ss) {
  if (!ss) ss = getSpreadsheet();
  var sh   = ss.getSheetByName(SHEET_MOTOR);
  if (!sh) return [];
  var data = sh.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(function(r) {
    return {
      id:             r[0],
      namaUnit:       r[1],
      nopol:          r[2],
      atasNama:       r[3],
      warna:          r[4],
      tglBayar:       r[5] ? Utilities.formatDate(new Date(r[5]), 'Asia/Jakarta', 'yyyy-MM-dd') : '',
      nominal:        r[6],
      jatuhTempo:     r[7] ? Utilities.formatDate(new Date(r[7]), 'Asia/Jakarta', 'yyyy-MM-dd') : '',
      tglStnk:        r[8] ? Utilities.formatDate(new Date(r[8]), 'Asia/Jakarta', 'yyyy-MM-dd') : '',
      nominalStnk:    r[9],
      jatuhTempoStnk: r[10] ? Utilities.formatDate(new Date(r[10]), 'Asia/Jakarta', 'yyyy-MM-dd') : '',
      keterangan:     r[11],
      createdAt:      r[12],
      updatedAt:      r[13]
    };
  }).filter(function(r){ return r.id; });
}

/* ─────────────────────────────────────────────────────────────
   getSewaList
───────────────────────────────────────────────────────────── */
function getSewaList(ss) {
  if (!ss) ss = getSpreadsheet(); // Fallback jika dipanggil langsung dari editor
  var sh   = ss.getSheetByName(SHEET_SEWA);
  if (!sh) return [];
  var data = sh.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(function(r) {
    return {
      id:          r[0],
      namaRumah:   r[1],
      namaPenyewa: r[2],
      namaPemilik: r[3],
      keterangan:  r[4],
      biaya:       r[5],
      tglMulai:    r[6] ? Utilities.formatDate(new Date(r[6]), 'Asia/Jakarta', 'yyyy-MM-dd') : '',
      tglBerakhir: r[7] ? Utilities.formatDate(new Date(r[7]), 'Asia/Jakarta', 'yyyy-MM-dd') : '',
      linkMou:     r[8],
      linkKondisi: r[9],
      catatan:     r[10],
      createdAt:   r[11],
      updatedAt:   r[12]
    };
  }).filter(function(r){ return r.id; });
}

/* ─────────────────────────────────────────────────────────────
   getAdminConfig
───────────────────────────────────────────────────────────── */
function getAdminConfig(ss) {
  if (!ss) ss = getSpreadsheet(); // Fallback jika dipanggil langsung dari editor
  var sh   = ss.getSheetByName(SHEET_ADMIN);
  var cfg  = {
    emails: [], emailsPajak: [], emailsSewa: [],
    rangePajak: 30, rangeSewa: 30, rangeStnk: 14,
    frequency: 'daily', togglePajak: true, toggleStnk: true,
    toggleSewa: true, toggleMonthly: false, toggleSeparate: true, pin: '1234'
  };
  if (!sh) return cfg;
  var data = sh.getDataRange().getValues();
  data.slice(1).forEach(function(r) {
    var key = r[0]; var val = r[1];
    if (key === 'emails' || key === 'emailsPajak' || key === 'emailsSewa') {
      try { cfg[key] = JSON.parse(val); } catch(e) { cfg[key] = []; }
    } else if (key in cfg) {
      if (val === 'true') cfg[key] = true;
      else if (val === 'false') cfg[key] = false;
      else if (!isNaN(Number(val)) && val !== '') cfg[key] = Number(val);
      else cfg[key] = val;
    }
  });
  return cfg;
}

/* ─────────────────────────────────────────────────────────────
   saveRecord – Upsert a single record
───────────────────────────────────────────────────────────── */
function saveRecord(type, dataJson) {
  try {
    var d  = typeof dataJson === 'string' ? JSON.parse(dataJson) : dataJson;
    var ss = getSpreadsheet();
    if (type === 'pajak') {
      upsertPajak(ss, d);
    } else if (type === 'motor') {
      upsertMotor(ss, d);
    } else if (type === 'sewa') {
      upsertSewa(ss, d);
    }
    return { success: true };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function upsertPajak(ss, d) {
  if (!d || !d.id) return; // Tolak record tanpa ID – cegah baris kosong
  ensureSheetsExist(ss);
  var sh   = ss.getSheetByName(SHEET_PAJAK);
  cleanEmptyRows(sh); // Bersihkan baris tanpa ID sebelum upsert
  var data = sh.getDataRange().getValues();
  var row  = [
    d.id, d.namaUnit, d.nopol, d.atasNama, d.warna,
    d.tglBayar||'', d.nominal||0, d.jatuhTempo||'',
    d.tglStnk||'', d.nominalStnk||0, d.jatuhTempoStnk||'',
    d.keterangan||'', d.createdAt||new Date().toISOString(), d.updatedAt||new Date().toISOString()
  ];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === d.id) { sh.getRange(i + 1, 1, 1, row.length).setValues([row]); return; }
  }
  sh.appendRow(row);
}

function upsertMotor(ss, d) {
  if (!d || !d.id) return;
  ensureSheetsExist(ss);
  var sh   = ss.getSheetByName(SHEET_MOTOR);
  cleanEmptyRows(sh);
  var data = sh.getDataRange().getValues();
  var row  = [
    d.id, d.namaUnit, d.nopol, d.atasNama, d.warna,
    d.tglBayar||'', d.nominal||0, d.jatuhTempo||'',
    d.tglStnk||'', d.nominalStnk||0, d.jatuhTempoStnk||'',
    d.keterangan||'', d.createdAt||new Date().toISOString(), d.updatedAt||new Date().toISOString()
  ];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === d.id) { sh.getRange(i + 1, 1, 1, row.length).setValues([row]); return; }
  }
  sh.appendRow(row);
}

function upsertSewa(ss, d) {
  if (!d || !d.id) return; // Tolak record tanpa ID – cegah baris kosong
  ensureSheetsExist(ss);
  var sh   = ss.getSheetByName(SHEET_SEWA);
  cleanEmptyRows(sh); // Bersihkan baris tanpa ID sebelum upsert
  var data = sh.getDataRange().getValues();
  var row  = [
    d.id, d.namaRumah, d.namaPenyewa, d.namaPemilik,
    d.keterangan||'', d.biaya||0, d.tglMulai||'', d.tglBerakhir||'',
    d.linkMou||'', d.linkKondisi||'', d.catatan||'',
    d.createdAt||new Date().toISOString(), d.updatedAt||new Date().toISOString()
  ];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === d.id) { sh.getRange(i + 1, 1, 1, row.length).setValues([row]); return; }
  }
  sh.appendRow(row);
  
}

/* ───────────────────────────────────────────────────────────
   cleanEmptyRows – Hapus baris tanpa ID dari sebuah sheet
   Jika dipanggil tanpa argumen (dari editor), bersihkan semua sheet data.
─────────────────────────────────────────────────────────── */
function cleanEmptyRows(sh) {
  // Fallback: jika dipanggil langsung dari editor tanpa argumen
  if (!sh) {
    var ss = getSpreadsheet();
    [SHEET_PAJAK, SHEET_MOTOR, SHEET_SEWA].forEach(function(name) {
      var s = ss.getSheetByName(name);
      if (s) cleanEmptyRows(s);
    });
    Logger.log('✅ Semua baris kosong berhasil dihapus dari DataPajak, DataPajakMotor & DataSewa.');
    return;
  }
  var data = sh.getDataRange().getValues();
  // Hapus dari bawah ke atas agar index tidak bergeser
  for (var i = data.length - 1; i >= 1; i--) {
    if (!data[i][0] || data[i][0] === '') {
      sh.deleteRow(i + 1);
    }
  }
}

/* ─────────────────────────────────────────────────────────────
   deleteRecord
───────────────────────────────────────────────────────────── */
function deleteRecord(type, id) {
  try {
    var ss      = getSpreadsheet();
    var shName  = type === 'pajak' ? SHEET_PAJAK : (type === 'motor' ? SHEET_MOTOR : SHEET_SEWA);
    var sh      = ss.getSheetByName(shName);
    if (!sh) return { success: true };
    var data    = sh.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === id) { sh.deleteRow(i + 1); break; }
    }
    return { success: true };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

/* ─────────────────────────────────────────────────────────────
   saveAdminSettings
───────────────────────────────────────────────────────────── */
function saveAdminSettings(cfg) {
  try {
    var ss   = getSpreadsheet();
    ensureSheetsExist(ss);
    var sh   = ss.getSheetByName(SHEET_ADMIN);
    // Clear old data (keep header)
    if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
    // Write settings
    var toggleSep = cfg.toggleSeparate !== undefined ? cfg.toggleSeparate : (cfg.tgSeparate !== undefined ? cfg.tgSeparate : true);
    var rows = [
      ['emails',        JSON.stringify(cfg.emails || [])],
      ['emailsPajak',   JSON.stringify(cfg.emailsPajak || [])],
      ['emailsSewa',    JSON.stringify(cfg.emailsSewa || [])],
      ['rangePajak',    cfg.rangePajak],
      ['rangeSewa',     cfg.rangeSewa],
      ['rangeStnk',     cfg.rangeStnk],
      ['frequency',     cfg.frequency],
      ['togglePajak',   cfg.togglePajak !== undefined ? cfg.togglePajak : cfg.tgPajak],
      ['toggleStnk',    cfg.toggleStnk !== undefined ? cfg.toggleStnk : cfg.tgStnk],
      ['toggleSewa',    cfg.toggleSewa !== undefined ? cfg.toggleSewa : cfg.tgSewa],
      ['toggleMonthly', cfg.toggleMonthly !== undefined ? cfg.toggleMonthly : cfg.tgMonthly],
      ['toggleSeparate', toggleSep],
      ['pin',           cfg.pin || '1234']
    ];
    sh.getRange(2, 1, rows.length, 2).setValues(rows);
    // Setup / update trigger
    setupReminderTrigger(cfg.frequency);
    return { success: true };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

/* ─────────────────────────────────────────────────────────────
   setupReminderTrigger – Create time-based trigger
───────────────────────────────────────────────────────────── */
function setupReminderTrigger(frequency) {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'checkAndSendReminders') {
      ScriptApp.deleteTrigger(t);
    }
  });
  var builder = ScriptApp.newTrigger('checkAndSendReminders').timeBased();
  if (frequency === 'weekly') {
    builder.onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
  } else if (frequency === 'monthly') {
    builder.onMonthDay(1).atHour(8).create();
  } else {
    builder.everyDays(1).atHour(8).create();
  }
}

/* ─────────────────────────────────────────────────────────────
   checkAndSendReminders – Triggered automatically
───────────────────────────────────────────────────────────── */
function checkAndSendReminders() {
  var ss   = getSpreadsheet();
  var cfg  = getAdminConfig(ss);

  var recipientsPajak = (cfg.emailsPajak && cfg.emailsPajak.length) ? cfg.emailsPajak : (cfg.emails || []);
  var recipientsSewa  = (cfg.emailsSewa && cfg.emailsSewa.length) ? cfg.emailsSewa : (cfg.emails || []);
  var recipientsAll   = (cfg.emails && cfg.emails.length) ? cfg.emails : Array.from(new Set(recipientsPajak.concat(recipientsSewa)));

  if (!recipientsAll || !recipientsAll.length) return;

  var today = new Date(); today.setHours(0,0,0,0);
  var kendaraanItems = [];
  var sewaItems = [];

  // ── Check Pajak ───────────────────────────────────────────
  if (cfg.togglePajak) {
    var allPajakList = getPajakList(ss).concat(getMotorList(ss));
    allPajakList.forEach(function(d) {
      if (!d.jatuhTempo) return;
      var days = Math.ceil((new Date(d.jatuhTempo) - today) / 86400000);
      if (days >= 0 && days <= cfg.rangePajak) {
        kendaraanItems.push({
          type: 'PAJAK',
          title: d.namaUnit + ' (' + d.nopol + ')',
          detail: 'Atas nama: ' + d.atasNama,
          days: days,
          date: d.jatuhTempo,
          nominal: d.nominal
        });
      }
    });
  }

  // ── Check STNK ───────────────────────────────────────────
  if (cfg.toggleStnk) {
    var allPajakList2 = getPajakList(ss).concat(getMotorList(ss));
    allPajakList2.forEach(function(d) {
      if (!d.jatuhTempoStnk) return;
      var days = Math.ceil((new Date(d.jatuhTempoStnk) - today) / 86400000);
      if (days >= 0 && days <= cfg.rangeStnk) {
        kendaraanItems.push({
          type: 'STNK',
          title: d.namaUnit + ' (' + d.nopol + ')',
          detail: 'Atas nama: ' + d.atasNama,
          days: days,
          date: d.jatuhTempoStnk,
          nominal: d.nominalStnk
        });
      }
    });
  }

  // ── Check Sewa ───────────────────────────────────────────
  if (cfg.toggleSewa) {
    var sewaList = getSewaList(ss);
    sewaList.forEach(function(d) {
      if (!d.tglBerakhir) return;
      var days = Math.ceil((new Date(d.tglBerakhir) - today) / 86400000);
      if (days >= 0 && days <= cfg.rangeSewa) {
        sewaItems.push({
          type: 'SEWA',
          title: d.namaRumah,
          detail: 'Penyewa: ' + d.namaPenyewa + ' | Pemilik: ' + d.namaPemilik,
          days: days,
          date: d.tglBerakhir,
          nominal: d.biaya
        });
      }
    });
  }

  var isSeparate = (cfg.toggleSeparate !== undefined ? cfg.toggleSeparate : cfg.tgSeparate);

  if (isSeparate) {
    if (kendaraanItems.length > 0 && recipientsPajak.length > 0) {
      sendReminderEmail(recipientsPajak, kendaraanItems, false, '🚗 Reminder Pajak & STNK Kendaraan');
    }
    if (sewaItems.length > 0 && recipientsSewa.length > 0) {
      sendReminderEmail(recipientsSewa, sewaItems, false, '🏠 Reminder Sewa Rumah');
    }
  } else {
    var allItems = kendaraanItems.concat(sewaItems);
    if (allItems.length > 0 && recipientsAll.length > 0) {
      sendReminderEmail(recipientsAll, allItems, false, '⚠️ Reminder Riayah Center');
    }
  }
}

/* ─────────────────────────────────────────────────────────────
   sendTestReminderEmail – Called from frontend
───────────────────────────────────────────────────────────── */
function sendTestReminderEmail(cfg) {
  try {
    var ss   = getSpreadsheet();
    var pajak = getPajakList(ss);
    var sewa  = getSewaList(ss);
    var kendaraanItems = [];
    var sewaItems = [];

    pajak.slice(0, 2).forEach(function(d) {
      kendaraanItems.push({ type:'PAJAK', title: d.namaUnit+' ('+d.nopol+')', detail:'Atas nama: '+d.atasNama, days:15, date:d.jatuhTempo, nominal:d.nominal });
    });
    sewa.slice(0, 2).forEach(function(d) {
      sewaItems.push({ type:'SEWA', title: d.namaRumah, detail:'Penyewa: '+d.namaPenyewa, days:20, date:d.tglBerakhir, nominal:d.biaya });
    });

    if (!kendaraanItems.length && !sewaItems.length) {
      kendaraanItems = [{ type:'TEST', title:'Contoh Kendaraan', detail:'Nomor Polisi: BM 1234 AB', days:10, date:'2025-12-31', nominal:500000 }];
      sewaItems = [{ type:'TEST', title:'Contoh Rumah Sewa', detail:'Penyewa: Fulan', days:12, date:'2025-12-31', nominal:15000000 }];
    }

    var recipientsPajak = (cfg.emailsPajak && cfg.emailsPajak.length) ? cfg.emailsPajak : (cfg.emails || []);
    var recipientsSewa  = (cfg.emailsSewa && cfg.emailsSewa.length) ? cfg.emailsSewa : (cfg.emails || []);
    var recipientsAll   = (cfg.emails && cfg.emails.length) ? cfg.emails : Array.from(new Set(recipientsPajak.concat(recipientsSewa)));

    var isSeparate = (cfg.toggleSeparate !== undefined ? cfg.toggleSeparate : cfg.tgSeparate);

    if (isSeparate) {
      if (kendaraanItems.length && recipientsPajak.length) sendReminderEmail(recipientsPajak, kendaraanItems, true, '🚗 Reminder Pajak & STNK Kendaraan');
      if (sewaItems.length && recipientsSewa.length) sendReminderEmail(recipientsSewa, sewaItems, true, '🏠 Reminder Sewa Rumah');
    } else {
      var allItems = kendaraanItems.concat(sewaItems);
      if (recipientsAll.length) sendReminderEmail(recipientsAll, allItems, true, '⚠️ Reminder Riayah Center');
    }
    return { success: true };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

/* ─────────────────────────────────────────────────────────────
   sendReminderEmail – Core email sender
───────────────────────────────────────────────────────────── */
function sendReminderEmail(recipients, items, isTest, customSubject) {
  var today     = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd MMMM yyyy');
  var subjectPrefix = customSubject || '⚠️ Reminder Riayah Center';
  var subject   = (isTest ? '[TES] ' : '') + subjectPrefix + ' – ' + today;

  var typeLabel = { PAJAK:'Pajak Kendaraan', STNK:'STNK Kendaraan', SEWA:'Sewa Rumah', TEST:'Contoh Data' };
  var typeColor = { PAJAK:'#d4a944', STNK:'#5da64a', SEWA:'#4c9fcf', TEST:'#888' };

  // Format tanggal 'yyyy-MM-dd' → '20 Oktober 2021'
  var BULAN_ID = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
  function formatDateID(dateStr) {
    if (!dateStr) return '-';
    try {
      var d = new Date(dateStr + 'T00:00:00+07:00');
      if (isNaN(d.getTime())) return dateStr;
      return d.getDate() + ' ' + BULAN_ID[d.getMonth()] + ' ' + d.getFullYear();
    } catch(e) { return dateStr; }
  }

  var rows = items.map(function(item) {
    var urgency = item.days <= 7  ? '🔴 Sangat Mendesak' :
                  item.days <= 14 ? '🟠 Mendesak' :
                  item.days <= 30 ? '🟡 Perlu Perhatian' : '🟢 Informatif';
    return '<tr style="border-bottom:1px solid #2a3d2a">' +
      '<td style="padding:12px 16px">' +
        '<span style="background:' + (typeColor[item.type]||'#888') + '20;color:' + (typeColor[item.type]||'#888') + ';padding:3px 10px;border-radius:12px;font-size:11px;font-weight:600">' + (typeLabel[item.type]||item.type) + '</span>' +
      '</td>' +
      '<td style="padding:12px 16px"><strong style="color:#e8f5e2">' + item.title + '</strong><br><span style="font-size:12px;color:#9cbf90">' + item.detail + '</span></td>' +
      '<td style="padding:12px 16px;text-align:center"><strong style="color:' + (item.days<=7?'#e05555':item.days<=14?'#d4a944':'#7ecf65') + ';font-size:18px">' + item.days + '</strong><br><span style="font-size:11px;color:#5a8050">hari lagi</span></td>' +
      '<td style="padding:12px 16px;color:#9cbf90">' + formatDateID(item.date) + '</td>' +
      '<td style="padding:12px 16px;color:#7ecf65;font-weight:600">' + formatCurrencyGAS(item.nominal) + '</td>' +
      '<td style="padding:12px 16px;font-size:12px">' + urgency + '</td>' +
    '</tr>';
  }).join('');

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body style="margin:0;padding:0;background:#0d1f10;font-family:Arial,sans-serif">' +
    '<div style="max-width:700px;margin:20px auto;background:#122015;border-radius:16px;overflow:hidden;border:1px solid rgba(100,180,80,0.2)">' +
    '<div style="background:linear-gradient(135deg,#1a3a1e,#0d2a10);padding:32px 28px;border-bottom:1px solid rgba(100,180,80,0.15)">' +
      '<div style="display:flex;align-items:center;gap:12px">' +
        '<div style="width:44px;height:44px;background:linear-gradient(135deg,#5da64a,#7ecf65);border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:22px">🕌</div>' +
        '<div><h1 style="margin:0;font-size:20px;font-weight:700;color:#7ecf65">Riayah Center</h1><p style="margin:0;font-size:12px;color:#5a8050">' + (customSubject || 'Reminder Otomatis') + ' – ' + today + (isTest?' &nbsp;<span style="background:#d4a944;color:#000;padding:2px 8px;border-radius:8px;font-size:10px;font-weight:700">TES</span>':'') + '</p></div>' +
      '</div>' +
    '</div>' +
    '<div style="padding:28px">' +
      '<p style="color:#9cbf90;font-size:14px;margin-bottom:20px">Berikut adalah data yang <strong>mendekati jatuh tempo</strong> dan memerlukan tindakan segera:</p>' +
      '<table style="width:100%;border-collapse:collapse;background:#0d1f10;border-radius:12px;overflow:hidden">' +
        '<thead><tr style="background:#1a3a1e">' +
          '<th style="padding:12px 16px;text-align:left;color:#5da64a;font-size:11px;text-transform:uppercase;letter-spacing:0.7px">Tipe</th>' +
          '<th style="padding:12px 16px;text-align:left;color:#5da64a;font-size:11px;text-transform:uppercase">Data</th>' +
          '<th style="padding:12px 16px;text-align:center;color:#5da64a;font-size:11px;text-transform:uppercase">Sisa Hari</th>' +
          '<th style="padding:12px 16px;text-align:left;color:#5da64a;font-size:11px;text-transform:uppercase">Tanggal</th>' +
          '<th style="padding:12px 16px;text-align:left;color:#5da64a;font-size:11px;text-transform:uppercase">Nominal</th>' +
          '<th style="padding:12px 16px;text-align:left;color:#5da64a;font-size:11px;text-transform:uppercase">Status</th>' +
        '</tr></thead>' +
        '<tbody>' + rows + '</tbody>' +
      '</table>' +
    '</div>' +
    '<div style="padding:20px 28px;border-top:1px solid rgba(100,180,80,0.15);text-align:center">' +
      '<p style="margin:0;font-size:12px;color:#5a8050">Email ini dikirim otomatis oleh sistem <strong style="color:#7ecf65">Riayah Center</strong> | Masjid Munzalan Mubarakan</p>' +
    '</div>' +
    '</div></body></html>';

  var plain = '[Riayah Center] ' + (customSubject || 'Reminder') + ' – ' + today + '\n\n' +
    items.map(function(item){ return '• [' + (typeLabel[item.type]||item.type) + '] ' + item.title + ' – ' + item.days + ' hari lagi (' + formatDateID(item.date) + ')'; }).join('\n');

  // Kirim satu email ke semua penerima sekaligus:
  // - Penerima pertama masuk ke "to"
  // - Penerima lainnya masuk ke "bcc" agar semua menerima email dalam 1 API call
  // (Hemat kuota & menghindari gagal sebagian akibat rate-limit)
  if (!recipients || recipients.length === 0) return;
  var validRecipients = recipients.filter(function(e){ return e && e.trim(); });
  if (validRecipients.length === 0) return;

  try {
    var toEmail  = validRecipients[0];
    var bccEmail = validRecipients.slice(1).join(',');
    var mailOptions = { to: toEmail, subject: subject, body: plain, htmlBody: html };
    if (bccEmail) mailOptions.bcc = bccEmail;
    MailApp.sendEmail(mailOptions);
    Logger.log('Email berhasil dikirim ke: ' + validRecipients.join(', '));
  } catch(e) {
    Logger.log('Gagal mengirim email: ' + e.toString());
    // Fallback: coba kirim satu per satu jika batch gagal
    validRecipients.forEach(function(email) {
      try {
        MailApp.sendEmail({ to: email, subject: subject, body: plain, htmlBody: html });
      } catch(e2) {
        Logger.log('Fallback gagal kirim ke ' + email + ': ' + e2.toString());
      }
    });
  }
}

/* ─────────────────────────────────────────────────────────────
   formatCurrencyGAS
───────────────────────────────────────────────────────────── */
function formatCurrencyGAS(n) {
  if (!n || isNaN(Number(n))) return 'Rp 0';
  return 'Rp ' + Number(n).toLocaleString('id-ID');
}

/* ─────────────────────────────────────────────────────────────
   sendMonthlyReport – For monthly laporan bulanan
───────────────────────────────────────────────────────────── */
function sendMonthlyReport() {
  var ss  = getSpreadsheet();
  var cfg = getAdminConfig(ss);
  var recipients = (cfg.emails && cfg.emails.length) ? cfg.emails : Array.from(new Set((cfg.emailsPajak||[]).concat(cfg.emailsSewa||[])));
  if (!cfg.toggleMonthly || !recipients.length) return;
  var pajak = getPajakList(ss);
  var sewa  = getSewaList(ss);
  var today = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd MMMM yyyy');
  var totalPajakNominal = pajak.reduce(function(s,d){ return s+(Number(d.nominal)||0); },0);
  var totalSewaBy       = sewa.reduce(function(s,d){ return s+(Number(d.biaya)||0); },0);
  var subject = '📊 Laporan Bulanan Riayah Center – ' + today;
  var html =
    '<div style="font-family:Arial;background:#0d1f10;color:#e8f5e2;padding:20px;max-width:600px;margin:0 auto;border-radius:16px">' +
    '<h2 style="color:#7ecf65">📊 Laporan Bulanan Riayah Center</h2>' +
    '<p style="color:#9cbf90">' + today + '</p>' +
    '<table style="width:100%;border-collapse:collapse;margin-top:16px">' +
    '<tr><td style="padding:10px;border:1px solid #2a3d2a;color:#9cbf90">Total Kendaraan</td><td style="padding:10px;border:1px solid #2a3d2a;color:#7ecf65;font-weight:bold">' + pajak.length + ' unit</td></tr>' +
    '<tr><td style="padding:10px;border:1px solid #2a3d2a;color:#9cbf90">Total Nominal Pajak</td><td style="padding:10px;border:1px solid #2a3d2a;color:#7ecf65;font-weight:bold">' + formatCurrencyGAS(totalPajakNominal) + '</td></tr>' +
    '<tr><td style="padding:10px;border:1px solid #2a3d2a;color:#9cbf90">Total Rumah Sewa</td><td style="padding:10px;border:1px solid #2a3d2a;color:#7ecf65;font-weight:bold">' + sewa.length + ' unit</td></tr>' +
    '<tr><td style="padding:10px;border:1px solid #2a3d2a;color:#9cbf90">Total Biaya Sewa</td><td style="padding:10px;border:1px solid #2a3d2a;color:#7ecf65;font-weight:bold">' + formatCurrencyGAS(totalSewaBy) + '</td></tr>' +
    '</table>' +
    '<p style="color:#5a8050;font-size:12px;margin-top:20px">Dikirim otomatis oleh Riayah Center – Masjid Munzalan Mubarakan</p></div>';
    recipients.forEach(function(email) {
    try { MailApp.sendEmail({ to: email, subject: subject, body: 'Laporan Bulanan Riayah Center', htmlBody: html }); } catch(e) {}
  });
}

/* ─────────────────────────────────────────────────────────────
   uploadFileToDrive – Upload file (base64) to Google Drive folder
───────────────────────────────────────────────────────────── */
function uploadFileToDrive(fileData) {
  try {
    var data = typeof fileData === 'string' ? JSON.parse(fileData) : fileData;
    if (!data || !data.base64) throw new Error('Data file tidak valid atau kosong');

    var folderName = 'Dokumen Riayah Center';
    var folders = DriveApp.getFoldersByName(folderName);
    var folder;
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(folderName);
    }

    var base64Parts = data.base64.split(',');
    var base64Content = base64Parts.length > 1 ? base64Parts[1] : base64Parts[0];
    var mimeType = data.mimeType || 'application/octet-stream';
    var fileName = data.fileName || ('Dokumen_' + new Date().getTime());

    var bytes = Utilities.base64Decode(base64Content);
    var blob = Utilities.newBlob(bytes, mimeType, fileName);
    var file = folder.createFile(blob);

    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (permErr) {
      // Ignored if domain policy limits public sharing
    }

    return {
      success: true,
      url: file.getUrl(),
      fileId: file.getId(),
      fileName: file.getName()
    };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}
