/**
 * PF-Protokoły — backend V1
 * Wklej ten kod w: Arkusz "Baza_glowna_PF-Protokoly" -> Rozszerzenia -> Apps Script
 * Wdroz jako aplikacje internetowa: Wykonaj jako "Ja", Kto ma dostep "Kazdy".
 *
 * Endpointy:
 *  GET  ?action=list_klienci
 *  GET  ?action=list_obiekty&klientId=XXX
 *  POST {action:'add_klient', nazwa}
 *  POST {action:'add_obiekt', klientId, klientNazwa, nazwa}
 *  POST {action:'delete_obiekt', id}
 *  POST {action:'bulk_import', rows:[{klient, obiekt}, ...]}
 */

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function getKlienciSheet() {
  const sheet = getSheet('Klienci');
  if (sheet.getLastRow() === 0) sheet.appendRow(['ID', 'Nazwa', 'DataUtworzenia']);
  return sheet;
}

function getObiektySheet() {
  const sheet = getSheet('Obiekty');
  if (sheet.getLastRow() === 0) sheet.appendRow(['ID', 'KlientID', 'KlientNazwa', 'Nazwa', 'DataUtworzenia']);
  return sheet;
}

function newId(prefix) {
  return prefix + '_' + new Date().getTime() + '_' + Math.floor(Math.random() * 10000);
}

function nowStr() {
  return Utilities.formatDate(new Date(), 'Europe/Warsaw', 'dd.MM.yyyy HH:mm');
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Musi być identyczne z ADMIN_PASSWORD w index.html — to bariera przed
// wywołaniem tego /exec przez kogoś, kto zna sam adres URL, ale nie appkę.
const SHARED_SECRET = 'Hyousng$19892508';

function checkAuth(token) {
  if (token !== SHARED_SECRET) {
    throw new Error('Brak autoryzacji');
  }
}

function listKlienci() {
  const sheet = getKlienciSheet();
  const rows = sheet.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    out.push({ id: rows[i][0], nazwa: rows[i][1], dataUtworzenia: rows[i][2] });
  }
  out.sort((a, b) => a.nazwa.localeCompare(b.nazwa, 'pl'));
  return out;
}

function listObiekty(klientId) {
  const sheet = getObiektySheet();
  const rows = sheet.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    if (String(rows[i][1]) !== String(klientId)) continue;
    out.push({ id: rows[i][0], klientId: rows[i][1], klientNazwa: rows[i][2], nazwa: rows[i][3], dataUtworzenia: rows[i][4] });
  }
  out.sort((a, b) => a.nazwa.localeCompare(b.nazwa, 'pl'));
  return out;
}

function findKlientByNazwa(nazwa) {
  const sheet = getKlienciSheet();
  const rows = sheet.getDataRange().getValues();
  const q = String(nazwa).trim().toLowerCase();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).trim().toLowerCase() === q) {
      return { id: rows[i][0], nazwa: rows[i][1] };
    }
  }
  return null;
}

function findObiektByNazwa(klientId, nazwa) {
  const sheet = getObiektySheet();
  const rows = sheet.getDataRange().getValues();
  const q = String(nazwa).trim().toLowerCase();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]) === String(klientId) && String(rows[i][3]).trim().toLowerCase() === q) {
      return { id: rows[i][0], nazwa: rows[i][3] };
    }
  }
  return null;
}

function addKlient(nazwa) {
  nazwa = String(nazwa).trim();
  if (!nazwa) throw new Error('Brak nazwy klienta');
  const existing = findKlientByNazwa(nazwa);
  if (existing) return existing;
  const sheet = getKlienciSheet();
  const id = newId('k');
  sheet.appendRow([id, nazwa, nowStr()]);
  return { id: id, nazwa: nazwa };
}

function addObiekt(klientId, klientNazwa, nazwa) {
  nazwa = String(nazwa).trim();
  if (!nazwa) throw new Error('Brak nazwy obiektu');
  if (!klientId) throw new Error('Brak klienta');
  const existing = findObiektByNazwa(klientId, nazwa);
  if (existing) return existing;
  const sheet = getObiektySheet();
  const id = newId('o');
  sheet.appendRow([id, klientId, klientNazwa, nazwa, nowStr()]);
  return { id: id, klientId: klientId, klientNazwa: klientNazwa, nazwa: nazwa };
}

function deleteObiekt(id) {
  const sheet = getObiektySheet();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

function bulkImport(rows) {
  const results = { klienciDodani: 0, obiektyDodane: 0, pominiete: 0 };
  const klientCache = {};
  rows.forEach(function (r) {
    const klientNazwa = String(r.klient || '').trim();
    const obiektNazwa = String(r.obiekt || '').trim();
    if (!klientNazwa || !obiektNazwa) {
      results.pominiete++;
      return;
    }
    let klient = klientCache[klientNazwa.toLowerCase()];
    if (!klient) {
      const beforeExisted = !!findKlientByNazwa(klientNazwa);
      klient = addKlient(klientNazwa);
      if (!beforeExisted) results.klienciDodani++;
      klientCache[klientNazwa.toLowerCase()] = klient;
    }
    const obiektExisted = !!findObiektByNazwa(klient.id, obiektNazwa);
    addObiekt(klient.id, klient.nazwa, obiektNazwa);
    if (!obiektExisted) results.obiektyDodane++;
  });
  return results;
}

function doGet(e) {
  try {
    checkAuth(e.parameter.token);
    const action = e.parameter.action;
    if (action === 'list_klienci') {
      return jsonOut({ ok: true, data: listKlienci() });
    }
    if (action === 'list_obiekty') {
      return jsonOut({ ok: true, data: listObiekty(e.parameter.klientId) });
    }
    return jsonOut({ ok: false, error: 'Nieznana akcja' });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    checkAuth(body.token);
    const action = body.action;
    if (action === 'add_klient') {
      return jsonOut({ ok: true, data: addKlient(body.nazwa) });
    }
    if (action === 'add_obiekt') {
      return jsonOut({ ok: true, data: addObiekt(body.klientId, body.klientNazwa, body.nazwa) });
    }
    if (action === 'delete_obiekt') {
      return jsonOut({ ok: true, deleted: deleteObiekt(body.id) });
    }
    if (action === 'bulk_import') {
      return jsonOut({ ok: true, data: bulkImport(body.rows || []) });
    }
    return jsonOut({ ok: false, error: 'Nieznana akcja' });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}
