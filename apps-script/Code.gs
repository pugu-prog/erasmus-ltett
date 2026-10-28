/**
 * Erasmus+ LTEtt – Verbindung zwischen der Schülerseite (GitHub Pages) und dieser Google-Tabelle.
 *
 * Einrichten (einmalig, in der Tabelle „Erasmus+ Stand 2026-27 (Lehrer)“):
 *   1. Erweiterungen › Apps Script öffnen, den ganzen Inhalt durch diese Datei ersetzen, speichern.
 *   2. Links „Projekteinstellungen“ (Zahnrad) › Skripteigenschaften › Eigenschaft hinzufügen:
 *        LEHRER_CODE = <ein Code, den du den Lehrern gibst>
 *   3. Bereitstellen › Neue Bereitstellung › Typ „Web-App“
 *        Ausführen als: Ich      Zugriff: Jeder
 *      › Bereitstellen › Zugriff erlauben › die Web-App-URL (endet auf /exec) kopieren.
 *   4. Die URL Claude geben (oder in index.html bei API_URL eintragen).
 *
 * Dokumente abgeben (Upload durch Schüler):
 *   - Jeder Schüler hat einen eigenen Code. Die Codes stehen im Blatt „Codes“ (Name | Code) dieser Tabelle;
 *     die Webseite liest dieses Blatt nie. Fehlende Codes werden automatisch erzeugt, sobald sich ein Lehrer
 *     auf der Seite anmeldet oder speichert (oder hier im Editor die Funktion codesErzeugen ausführen).
 *   - Einen Code ändern: im Blatt „Codes“ einfach überschreiben.
 *   - Die Dateien landen in Google Drive im Ordner „Erasmus+ Abgaben 2026-27“, ein Unterordner pro Schüler.
 *   - Nach einer Änderung am Skript: Bereitstellen › Bereitstellungen verwalten › Bearbeiten (Stift)
 *     › Version: Neue Version › Bereitstellen. Die URL bleibt gleich. Beim ersten Mal Drive-Zugriff erlauben.
 *
 * Datenschutz (Stufe A, 28.10.2026):
 *   - Schüler sehen nur ihren eigenen Stand: Aktion „mein_stand“ mit dem persönlichen Code.
 *   - Die ganze Liste gibt es nur mit dem Lehrer-Code (Aktion „liste“, „save“).
 *   - Skripteigenschaft STAND_PRIVAT = 1  →  der offene Abruf (GET) gibt keine Schülerdaten mehr heraus.
 *     Erst setzen, wenn die neue Webseite online ist (die alte Seite braucht den GET noch).
 *   - Zu viele falsche Codes → kurze Sperre (siehe SPERRE unten).
 *
 * Verschlüsselung (Stufe B):
 *   - Der Browser des Schülers verschlüsselt Fiche und Dateien (AES-GCM, Schlüssel mit RSA-OAEP eingepackt).
 *     Google speichert nur Geheimtext: Dateien enden auf .enc, Fiches stehen im Blatt „Verschlüsselte Fiches“.
 *   - Den Schlüssel richtet Guy einmal in der Lehreransicht ein. Gespeichert werden hier nur der öffentliche
 *     Schlüssel und der private Schlüssel, eingepackt mit Passphrase bzw. Wiederherstellungscode
 *     (Skripteigenschaften SCHLUESSEL_PUB, SCHLUESSEL_FP, SCHLUESSEL_PASS, SCHLUESSEL_REC).
 *     Neuen Schlüssel einrichten: diese vier Eigenschaften löschen (alte Abgaben sind dann nicht mehr lesbar!).
 *   - Skripteigenschaft NUR_VERSCHLUESSELT = 1  →  unverschlüsselte Fiches/Dateien werden abgelehnt.
 *     Erst setzen, wenn die neue Webseite online ist.
 *   - Mit dem Lehrer-Code bekommt man nur Geheimtext; lesen kann ihn nur, wer die Passphrase kennt.
 *
 * Ändern (POST) geht nur mit dem Lehrer-Code bzw. für die eigenen Abgaben mit dem Schüler-Code.
 * Erstes Blatt der Tabelle = Liste. Zeile 1 = Spaltenköpfe. Zeilen, deren Name mit # beginnt, werden ignoriert.
 */

function sheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
}

function readAll_() {
  const values = sheet_().getDataRange().getDisplayValues();
  return {
    headers: values[0],
    rows: values.slice(1).filter(function (r) { return r[0] && String(r[0]).charAt(0) !== '#'; }),
    updated: new Date().toISOString()
  };
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function standPrivat_() {
  return String(PropertiesService.getScriptProperties().getProperty('STAND_PRIVAT') || '').trim() === '1';
}

function doGet() {
  if (standPrivat_()) return json_({ privat: true, headers: [], rows: [], updated: new Date().toISOString() });
  try { statusAbgleichen_(); } catch (err) {}
  return json_(readAll_());
}

// ---------- Sperre bei falschen Codes ----------
// pro Schüler: 5 Fehlversuche, dann 15 Min. gesperrt; Code ohne Namen (alle zusammen): 30 Fehlversuche / 15 Min.;
// Lehrer-Code (alle zusammen): 10 Fehlversuche / 15 Min.
var SPERRE = { schueler: [5, 900], suche: [30, 900], lehrer: [10, 900] };

function sperrKey_(art, wer) { return 'sperre:' + art + ':' + String(wer || '').trim().toLowerCase(); }

function gesperrt_(art, wer) {
  var n = Number(CacheService.getScriptCache().get(sperrKey_(art, wer)) || 0);
  return n >= SPERRE[art][0];
}

function fehlversuch_(art, wer) {
  var c = CacheService.getScriptCache(), k = sperrKey_(art, wer);
  c.put(k, String(Number(c.get(k) || 0) + 1), SPERRE[art][1]);
}

function lehrerOk_(req) {
  if (gesperrt_('lehrer', '')) return 'gesperrt';
  var expected = PropertiesService.getScriptProperties().getProperty('LEHRER_CODE');
  if (expected && String(req.code || '') === expected) return 'ok';
  fehlversuch_('lehrer', '');
  return 'code';
}

/** Prüft Name + persönlichen Code. Gibt 'ok', 'gesperrt' oder 'scode' zurück. */
function schuelerOk_(name, scode) {
  if (gesperrt_('schueler', name)) return 'gesperrt';
  var own = codeOf_(name);
  if (own && String(scode || '').trim().toUpperCase() === own.toUpperCase()) return 'ok';
  fehlversuch_('schueler', name);
  return 'scode';
}

/** Schüler sieht nur die eigene Zeile. Der Code allein reicht (der Name kommt vom Server). */
function meinStand_(req) {
  if (gesperrt_('suche', '')) return json_({ ok: false, error: 'gesperrt' });
  var scode = String(req.scode || '').trim().toUpperCase();
  var name = '';
  if (scode.length >= 6) {
    var v = codesSheet_().getDataRange().getValues().slice(1)
      .filter(function (r) { return r[0] && String(r[1]).trim().toUpperCase() === scode; });
    if (v.length === 1) name = String(v[0][0]).trim();
  }
  if (!name || gesperrt_('schueler', name)) {
    fehlversuch_('suche', '');
    return json_({ ok: false, error: name ? 'gesperrt' : 'scode' });
  }
  try { statusAbgleichen_(); } catch (err) {}
  var all = readAll_();
  var row = all.rows.filter(function (r) { return String(r[0]).trim().toLowerCase() === name.toLowerCase(); })[0];
  if (!row) return json_({ ok: false, error: 'name' });
  return json_({ ok: true, headers: all.headers, rows: [row], updated: all.updated });
}

/** Hakt „Fiche Donnée“ ab für alle, deren Fiche im Blatt „Fiche données“ steht,
 *  und „Wunsch abgegeben“ für alle, die selbst einen eigenen Betrieb eingetragen haben. */
function statusAbgleichen_() {
  var sh = sheet_();
  var values = sh.getDataRange().getValues();
  var headers = values[0].map(String);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var done = {};
  var fs = findFicheSheet_(ss);
  if (fs && fs.getLastRow() > 1) fs.getRange(2, 2, fs.getLastRow() - 1, 1).getValues().forEach(function (r) { if (r[0]) done[String(r[0]).trim().toLowerCase()] = true; });
  var es = ss.getSheetByName(ENC_FICHES);
  if (es && es.getLastRow() > 1) es.getRange(2, 2, es.getLastRow() - 1, 1).getValues().forEach(function (r) { if (r[0]) done[String(r[0]).trim().toLowerCase()] = true; });
  var wunsch = {};
  var vs = ss.getSheetByName('Betriebsvorschläge');
  if (vs && vs.getLastRow() > 1) vs.getRange(2, 2, vs.getLastRow() - 1, 5).getValues().forEach(function (r) {
    if (r[0] && String(r[4]).indexOf('aus Motivationsschreiben') < 0) wunsch[String(r[0]).trim().toLowerCase()] = true;
  });
  var kF = headers.indexOf('Fiche Donnée');
  var kW = headers.indexOf('Wunsch abgegeben');
  for (var i = 1; i < values.length; i++) {
    var n = String(values[i][0]).trim().toLowerCase();
    if (kF >= 0 && done[n] && String(values[i][kF]).trim().toLowerCase() !== 'x') sh.getRange(i + 1, kF + 1).setValue('x');
    if (wunsch[n]) {
      if (kW < 0) kW = colOf_(sh, headers, 'Wunsch abgegeben') - 1;
      if (String(values[i][kW] || '').trim().toLowerCase() !== 'x') sh.getRange(i + 1, kW + 1).setValue('x');
    }
  }
}

function doPost(e) {
  try {
    return handlePost_(e);
  } catch (err) {
    return json_({ ok: false, error: 'server', message: String(err && err.message || err) });
  }
}

/** Einmal im Editor ausführen: fragt die Erlaubnis für Google Drive ab und legt den Abgabe-Ordner an. */
function ordnerAnlegen() {
  var f = rootFolder_();
  Logger.log('Abgabe-Ordner: ' + f.getUrl());
  return f.getUrl();
}

function handlePost_(e) {
  var req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_request' }); }

  if (req.action === 'mein_stand') return meinStand_(req);
  if (req.action === 'pubkey') return json_({ ok: true, pub: prop_('SCHLUESSEL_PUB'), fp: prop_('SCHLUESSEL_FP') });
  if (req.action === 'upload') return upload_(req);
  if (req.action === 'fiche') return saveFiche_(req);
  if (req.action === 'wunsch') return saveWunsch_(req);

  var lo = lehrerOk_(req);
  if (lo !== 'ok') return json_({ ok: false, error: lo });
  if (req.action === 'schluessel') return schluessel_();
  if (req.action === 'schluessel_speichern') return schluesselSpeichern_(req);
  if (req.action === 'abgaben_liste') return abgabenListe_();
  if (req.action === 'abgabe_holen') return abgabeHolen_(req);
  if (req.action === 'fiches_holen') return fichesHolen_();
  if (req.action === 'liste') {
    try { statusAbgleichen_(); } catch (err) {}
    return json_({ ok: true, data: readAll_() });
  }
  if (req.action === 'check') {
    codesErzeugen();
    var url = null;
    try { url = rootFolder_().getUrl(); } catch (err) { url = null; }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ficheUrl = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx&gid=' + ficheSheet_().getSheetId();
    return json_({ ok: true, folderUrl: url, ficheUrl: ficheUrl });
  }
  if (req.action !== 'save') return json_({ ok: false, error: 'unknown_action' });

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet_();
    var values = sh.getDataRange().getValues();
    var headers = values[0].map(String);

    // Neue Spalten anlegen, falls die Seite welche schickt, die es noch nicht gibt
    (req.upsert || []).forEach(function (row) {
      Object.keys(row).forEach(function (h) {
        if (headers.indexOf(h) < 0) { headers.push(h); sh.getRange(1, headers.length).setValue(h); }
      });
    });

    function rowOf(name) {
      var col = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues();
      for (var i = 1; i < col.length; i++) {
        if (String(col[i][0]).trim().toLowerCase() === String(name).trim().toLowerCase()) return i + 1;
      }
      return -1;
    }

    (req.upsert || []).forEach(function (row) {
      if (!row.Name) return;
      var r = rowOf(row.Name);
      if (r < 0) { sh.appendRow([row.Name]); r = sh.getLastRow(); }
      var line = sh.getRange(r, 1, 1, headers.length).getValues()[0];
      headers.forEach(function (h, k) { if (Object.prototype.hasOwnProperty.call(row, h)) line[k] = row[h]; });
      sh.getRange(r, 1, 1, headers.length).setValues([line]);
    });

    (req.remove || []).forEach(function (name) {
      var r = rowOf(name);
      if (r > 1) sh.deleteRow(r);
    });

    SpreadsheetApp.flush();
    codesErzeugen();
    return json_({ ok: true, data: readAll_() });
  } finally {
    lock.releaseLock();
  }
}

// ---------- Dokumente abgeben ----------
var MAX_BYTES = 15 * 1024 * 1024;
var ERLAUBT = /\.(pdf|docx?|jpe?g|png|heic)$/i;

function rootFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('ABGABE_ORDNER');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var f = DriveApp.createFolder('Erasmus+ Abgaben 2026-27');
  props.setProperty('ABGABE_ORDNER', f.getId());
  return f;
}

function studentFolder_(name) {
  var root = rootFolder_();
  var it = root.getFoldersByName(name);
  return it.hasNext() ? it.next() : root.createFolder(name);
}

function upload_(req) {
  var name = String(req.name || '').trim();
  var typ = String(req.typ || '').trim();
  var col = String(req.col || '').trim();
  var filename = String(req.filename || '').trim();
  if (!name || !typ || !filename || !req.data) return json_({ ok: false, error: 'bad_request' });
  if (!ERLAUBT.test(filename)) return json_({ ok: false, error: 'type' });

  var sh = sheet_();
  var values = sh.getDataRange().getValues();
  var headers = values[0].map(String);
  var row = -1;
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim().toLowerCase() === name.toLowerCase()) { row = i + 1; break; }
  }
  if (row < 0) return json_({ ok: false, error: 'name' });
  var so = schuelerOk_(name, req.scode);
  if (so !== 'ok') return json_({ ok: false, error: so });

  var enc = req.enc === true;
  if (!enc && nurVerschluesselt_()) return json_({ ok: false, error: 'unverschluesselt' });
  var bytes = Utilities.base64Decode(req.data);
  if (bytes.length > MAX_BYTES + (enc ? 4096 : 0)) return json_({ ok: false, error: 'size' });

  var ext = filename.replace(/^.*\./, '').toLowerCase();
  var stamp = Utilities.formatDate(new Date(), 'Europe/Luxembourg', 'yyyy-MM-dd HH-mm');
  var blob = enc
    ? Utilities.newBlob(bytes, 'application/octet-stream', typ + ' - ' + name + ' - ' + stamp + '.' + ext + '.enc')
    : Utilities.newBlob(bytes, req.mime || 'application/octet-stream', typ + ' - ' + name + ' - ' + stamp + '.' + ext);
  var file;
  try { file = studentFolder_(name).createFile(blob); }
  catch (err) { return json_({ ok: false, error: 'drive', message: String(err && err.message || err) }); }

  // Im Stand als „eingegangen“ (i) markieren, wenn der Punkt noch offen ist
  var k = headers.indexOf(col);
  if (k >= 0) {
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      var cell = sh.getRange(row, k + 1);
      if (!String(cell.getValue()).trim()) cell.setValue('i');
    } finally { lock.releaseLock(); }
  }
  return json_({ ok: true, file: file.getName() });
}

// ---------- Persönliche Schüler-Codes ----------
var CODE_ZEICHEN = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';

function codesSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Codes');
  if (!sh) {
    sh = ss.insertSheet('Codes', ss.getSheets().length);
    sh.getRange(1, 1, 1, 2).setValues([['Name', 'Code']]).setFontWeight('bold');
    sh.setColumnWidth(1, 220);
  }
  return sh;
}

function codeOf_(name) {
  var v = codesSheet_().getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][0]).trim().toLowerCase() === String(name).trim().toLowerCase()) return String(v[i][1]).trim();
  }
  return '';
}

/** Legt für jeden Schüler ohne Code einen neuen 6-stelligen Code an. Bestehende Codes bleiben. */
function codesErzeugen() {
  var names = sheet_().getRange(2, 1, Math.max(sheet_().getLastRow() - 1, 1), 1).getValues()
    .map(function (r) { return String(r[0]).trim(); })
    .filter(function (n) { return n && n.charAt(0) !== '#'; });
  var cs = codesSheet_();
  var have = {};
  cs.getDataRange().getValues().slice(1).forEach(function (r) { if (r[0]) have[String(r[0]).trim().toLowerCase()] = true; });
  var used = {};
  cs.getDataRange().getValues().slice(1).forEach(function (r) { if (r[1]) used[String(r[1]).trim().toUpperCase()] = true; });
  var neu = [];
  names.forEach(function (n) {
    if (have[n.toLowerCase()]) return;
    var c = '';
    do { // Codes müssen eindeutig sein, weil „mein_stand“ den Schüler nur am Code erkennt
      c = '';
      for (var i = 0; i < 6; i++) c += CODE_ZEICHEN.charAt(Math.floor(Math.random() * CODE_ZEICHEN.length));
    } while (used[c]);
    used[c] = true;
    neu.push([n, c]);
  });
  if (neu.length) cs.getRange(cs.getLastRow() + 1, 1, neu.length, 2).setValues(neu);
  return neu.length;
}

// ---------- Fiche données élève ----------
var FICHE_FELDER = [["classe", "Classe actuelle"], ["nom", "Nom"], ["prenom", "Prénom"], ["tel_mobile", "Téléphone mobile"], ["tel_fixe", "Téléphone fixe"], ["nationalite", "Nationalité"], ["matricule", "Matricule (13 chiffres)"], ["adresse", "Adresse postale"], ["email", "Adresse électronique"], ["iban", "Numéro de compte (IBAN)"], ["banque", "Nom de la banque ou CCP"], ["bic", "Code BIC"], ["titulaire", "Titulaire du compte"], ["rep_nom", "Représentant légal (nom et prénom)"], ["rep_adresse", "Adresse du représentant légal"], ["rep_tel", "Téléphone du représentant légal"], ["rep_email", "Adresse électronique du représentant légal"]];

/** Findet das Blatt „Fiche données“ auch bei abweichender Schreibweise (Leerzeichen, é-Kodierung). */
function findFicheSheet_(ss) {
  var all = ss.getSheets();
  for (var i = 0; i < all.length; i++) {
    var n = all[i].getName().normalize('NFC').trim().toLowerCase();
    if (n.indexOf('fiche') === 0 && n !== 'fiche donnée') return all[i];
  }
  return null;
}

function ficheSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = findFicheSheet_(ss);
  if (!sh) {
    sh = ss.insertSheet('Fiche données', ss.getSheets().length);
    var head = ['Eingang', 'Name (Liste)'].concat(FICHE_FELDER.map(function (f) { return f[1]; }));
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function saveFiche_(req) {
  var name = String(req.name || '').trim();
  if (!name) return json_({ ok: false, error: 'bad_request' });
  var sh = sheet_();
  var values = sh.getDataRange().getValues();
  var headers = values[0].map(String);
  var row = -1;
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim().toLowerCase() === name.toLowerCase()) { row = i + 1; break; }
  }
  if (row < 0) return json_({ ok: false, error: 'name' });
  var so = schuelerOk_(name, req.scode);
  if (so !== 'ok') return json_({ ok: false, error: so });

  if (req.enc) return saveFicheEnc_(req, name, sh, row, headers);
  if (nurVerschluesselt_()) return json_({ ok: false, error: 'unverschluesselt' });

  var d = req.data || {};
  var OPTIONAL = { tel_fixe: true };
  var missing = FICHE_FELDER.filter(function (f) { return !OPTIONAL[f[0]] && !String(d[f[0]] || '').trim(); }).map(function (f) { return f[1]; });
  if (missing.length) return json_({ ok: false, error: 'missing', message: missing.join(', ') });

  var line = [new Date(), name].concat(FICHE_FELDER.map(function (f) { return "'" + String(d[f[0]] || '').trim(); }));
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var fs = ficheSheet_();
    var names = fs.getRange(1, 2, Math.max(fs.getLastRow(), 1), 1).getValues();
    var r = -1;
    for (var j = 1; j < names.length; j++) if (String(names[j][0]).trim().toLowerCase() === name.toLowerCase()) { r = j + 1; break; }
    if (r < 0) fs.appendRow(line); else fs.getRange(r, 1, 1, line.length).setValues([line]);
    var k = headers.indexOf('Fiche Donnée');
    if (k >= 0) {
      var cell = sh.getRange(row, k + 1);
      cell.setValue('x'); // alle Pflichtfelder geprüft → automatisch erledigt
    }
  } finally { lock.releaseLock(); }
  return json_({ ok: true });
}

// ---------- Wunschland / Wunschbetriebe ----------
var LAENDER = ['Deutschland', 'Portugal', 'Österreich', 'Kosovo'];

function colOf_(sh, headers, name) {
  var k = headers.indexOf(name);
  if (k < 0) { headers.push(name); k = headers.length - 1; sh.getRange(1, k + 1).setValue(name); }
  return k + 1;
}

function saveWunsch_(req) {
  var name = String(req.name || '').trim();
  var land = String(req.land || '').trim();
  var betriebe = (req.betriebe || []).map(function (b) { return String(b).trim(); }).filter(String);
  if (!name || !land || land.length > 60) return json_({ ok: false, error: 'bad_request' });
  var sh = sheet_();
  var values = sh.getDataRange().getValues();
  var headers = values[0].map(String);
  var row = -1;
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim().toLowerCase() === name.toLowerCase()) { row = i + 1; break; }
  }
  if (row < 0) return json_({ ok: false, error: 'name' });
  var so = schuelerOk_(name, req.scode);
  if (so !== 'ok') return json_({ ok: false, error: so });
  if (land !== 'Portugal') betriebe = [];
  var eigene = (req.eigene || []).slice(0, 10).map(function (e) {
    return { name: String(e.name || '').trim().slice(0, 100), ort: String(e.ort || '').trim().slice(0, 80), kontakt: String(e.kontakt || '').trim().slice(0, 200) };
  }).filter(function (e) { return e.name; });
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var cL = colOf_(sh, headers, 'Wunschland');
    var cB = colOf_(sh, headers, 'Wunschbetriebe');
    sh.getRange(row, cL).setValue(land);
    sh.getRange(row, cB).setValue(betriebe.join(', ').slice(0, 400));
    var cE = colOf_(sh, headers, 'Eigene Betriebe');
    // komplett = Land und (bei Portugal) mindestens ein Betrieb oder eigener Vorschlag → automatisch abhaken
    if (land !== 'Portugal' || betriebe.length || eigene.length) sh.getRange(row, colOf_(sh, headers, 'Wunsch abgegeben')).setValue('x');
    sh.getRange(row, cE).setValue(eigene.map(function (e) { return e.name + (e.ort ? ' (' + e.ort + ')' : ''); }).join('; ').slice(0, 500));
    if (eigene.length) {
      var vs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Betriebsvorschläge');
      if (!vs) {
        vs = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Betriebsvorschläge', SpreadsheetApp.getActiveSpreadsheet().getSheets().length);
        vs.getRange(1, 1, 1, 6).setValues([['Eingang', 'Schüler', 'Land', 'Betrieb', 'Ort', 'Webseite / Kontakt']]).setFontWeight('bold');
        vs.setFrozenRows(1);
      }
      eigene.forEach(function (e) { vs.appendRow([new Date(), name, land, e.name, e.ort, "'" + e.kontakt]); });
    }
  } finally { lock.releaseLock(); }
  return json_({ ok: true });
}

/** Im Editor ausführen: meldet doppelte Schüler-Codes (die müssen eindeutig sein, sonst geht „mein_stand“ nicht). */
function codesPruefen() {
  var seen = {}, doppelt = [];
  codesSheet_().getDataRange().getValues().slice(1).forEach(function (r) {
    var c = String(r[1]).trim().toUpperCase();
    if (!c) return;
    if (seen[c]) doppelt.push(seen[c] + ' / ' + r[0]); else seen[c] = String(r[0]);
  });
  Logger.log(doppelt.length ? 'Doppelte Codes: ' + doppelt.join('; ') : 'Alle Codes sind eindeutig.');
  return doppelt;
}

// ---------- Verschlüsselung (Stufe B) ----------
var ENC_FICHES = 'Verschlüsselte Fiches';
var B64 = /^[A-Za-z0-9+\/]+=*$/;

function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k) || ''; }
function nurVerschluesselt_() { return String(prop_('NUR_VERSCHLUESSELT')).trim() === '1'; }

function saveFicheEnc_(req, name, sh, row, headers) {
  var blob = String(req.enc || '');
  if (!prop_('SCHLUESSEL_PUB')) return json_({ ok: false, error: 'kein_schluessel' });
  if (blob.length < 500 || blob.length > 40000 || !B64.test(blob)) return json_({ ok: false, error: 'bad_request' });
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var es = ss.getSheetByName(ENC_FICHES);
    if (!es) {
      es = ss.insertSheet(ENC_FICHES, ss.getSheets().length);
      es.getRange(1, 1, 1, 3).setValues([['Eingang', 'Name (Liste)', 'Fiche (verschlüsselt, nur mit Passphrase lesbar)']]).setFontWeight('bold');
      es.setFrozenRows(1);
    }
    var line = [new Date(), name, "'" + blob];
    var names = es.getRange(1, 2, Math.max(es.getLastRow(), 1), 1).getValues();
    var r = -1;
    for (var j = 1; j < names.length; j++) if (String(names[j][0]).trim().toLowerCase() === name.toLowerCase()) { r = j + 1; break; }
    if (r < 0) es.appendRow(line); else es.getRange(r, 1, 1, 3).setValues([line]);
    var k = headers.indexOf('Fiche Donnée');
    if (k >= 0) sh.getRange(row, k + 1).setValue('x'); // Pflichtfelder hat der Browser vor dem Verschlüsseln geprüft
  } finally { lock.releaseLock(); }
  return json_({ ok: true });
}

function schluessel_() {
  return json_({ ok: true, pub: prop_('SCHLUESSEL_PUB'), fp: prop_('SCHLUESSEL_FP'), pass: prop_('SCHLUESSEL_PASS'), rec: prop_('SCHLUESSEL_REC') });
}

/** Nur einmal möglich: ein vorhandener Schlüssel wird nie überschrieben (sonst wären alte Abgaben verloren). */
function schluesselSpeichern_(req) {
  if (prop_('SCHLUESSEL_PUB')) return json_({ ok: false, error: 'schon_da' });
  var f = { SCHLUESSEL_PUB: req.pub, SCHLUESSEL_FP: req.fp, SCHLUESSEL_PASS: req.pass, SCHLUESSEL_REC: req.rec };
  for (var k in f) {
    var v = String(f[k] || '');
    if (!v || v.length > 8500) return json_({ ok: false, error: 'bad_request', message: k });
    f[k] = v;
  }
  PropertiesService.getScriptProperties().setProperties(f);
  return json_({ ok: true, fp: f.SCHLUESSEL_FP });
}

/** Liste der verschlüsselten Abgaben (ein Unterordner pro Schüler). Unverschlüsselte Dateien werden nur gezählt. */
function abgabenListe_() {
  var out = [], klartext = 0;
  var folders = rootFolder_().getFolders();
  while (folders.hasNext()) {
    var fo = folders.next(), files = fo.getFiles();
    while (files.hasNext()) {
      var f = files.next();
      if (/\.enc$/.test(f.getName())) out.push({ id: f.getId(), ordner: fo.getName(), name: f.getName(), size: f.getSize(), datum: f.getDateCreated().toISOString() });
      else klartext++;
    }
  }
  return json_({ ok: true, dateien: out, klartext: klartext });
}

/** Gibt eine verschlüsselte Abgabe als base64 zurück – nur .enc-Dateien aus dem Abgabe-Ordner, nichts anderes aus Drive. */
function abgabeHolen_(req) {
  var f;
  try { f = DriveApp.getFileById(String(req.id || '')); } catch (e) { return json_({ ok: false, error: 'nicht_gefunden' }); }
  if (!/\.enc$/.test(f.getName())) return json_({ ok: false, error: 'nicht_erlaubt' });
  var rootId = rootFolder_().getId(), ok = false;
  var parents = f.getParents();
  while (parents.hasNext()) {
    var p = parents.next(), pp = p.getParents();
    while (pp.hasNext()) if (pp.next().getId() === rootId) ok = true;
  }
  if (!ok) return json_({ ok: false, error: 'nicht_erlaubt' });
  return json_({ ok: true, name: f.getName(), data: Utilities.base64Encode(f.getBlob().getBytes()) });
}

/** Verschlüsselte Fiches für die Lehreransicht (unverschlüsselte alte Fiches werden nur gezählt). */
function fichesHolen_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var es = ss.getSheetByName(ENC_FICHES), out = [];
  if (es && es.getLastRow() > 1) es.getRange(2, 1, es.getLastRow() - 1, 3).getValues().forEach(function (r) {
    if (r[1] && r[2]) out.push({ eingang: r[0] instanceof Date ? r[0].toISOString() : String(r[0]), name: String(r[1]), enc: String(r[2]).replace(/^'/, '') });
  });
  var fs = findFicheSheet_(ss);
  return json_({ ok: true, fiches: out, klartext: fs ? Math.max(fs.getLastRow() - 1, 0) : 0 });
}

// ---------- Umzug in eine Tabellenkopie ohne Versionsverlauf (Stufe C) ----------
// 1. ALTE Tabelle: uebertragVorbereiten ausführen  →  Blatt „_uebertrag“ mit den Skripteigenschaften (ohne LEHRER_CODE).
// 2. Datei › Kopie erstellen.  3. In der KOPIE: uebertragUebernehmen ausführen, LEHRER_CODE von Hand eintragen,
//    dann Bereitstellen › Neue Bereitstellung › Web-App (Ausführen als: Ich, Zugriff: Jeder).
var UEBERTRAG = '_uebertrag';
var UEBERTRAG_KEYS = ['STAND_PRIVAT', 'NUR_VERSCHLUESSELT', 'ABGABE_ORDNER', 'SCHLUESSEL_PUB', 'SCHLUESSEL_FP', 'SCHLUESSEL_PASS', 'SCHLUESSEL_REC'];

function uebertragVorbereiten() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(UEBERTRAG) || ss.insertSheet(UEBERTRAG, ss.getSheets().length);
  sh.clear();
  var rows = UEBERTRAG_KEYS.map(function (k) { return [k, prop_(k)]; });
  sh.getRange(1, 1, rows.length, 2).setNumberFormat('@').setValues(rows);
  var leer = rows.filter(function (r) { return !r[1]; }).map(function (r) { return r[0]; });
  Logger.log('Blatt „_uebertrag“ angelegt.' + (leer.length ? ' Leer: ' + leer.join(', ') : ' Alle Eigenschaften drin.') + ' Jetzt Datei › Kopie erstellen.');
}

function uebertragUebernehmen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(UEBERTRAG);
  if (!sh) throw new Error('Blatt „_uebertrag“ fehlt. Erst in der alten Tabelle uebertragVorbereiten ausführen, dann kopieren.');
  var p = {};
  sh.getDataRange().getValues().forEach(function (r) {
    if (UEBERTRAG_KEYS.indexOf(String(r[0])) >= 0 && String(r[1])) p[String(r[0])] = String(r[1]);
  });
  PropertiesService.getScriptProperties().setProperties(p);
  ss.deleteSheet(sh);
  Logger.log('Übernommen: ' + Object.keys(p).join(', ') + '. ' + (prop_('LEHRER_CODE') ? 'LEHRER_CODE ist schon da.' : 'Jetzt noch LEHRER_CODE in den Skripteigenschaften von Hand eintragen.'));
}
