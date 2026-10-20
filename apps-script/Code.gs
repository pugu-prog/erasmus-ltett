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
 * Die Schüler lesen nur (GET). Ändern (POST) geht nur mit dem Lehrer-Code.
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

function doGet() {
  return json_(readAll_());
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

  if (req.action === 'upload') return upload_(req);
  if (req.action === 'fiche') return saveFiche_(req);
  if (req.action === 'wunsch') return saveWunsch_(req);

  var expected = PropertiesService.getScriptProperties().getProperty('LEHRER_CODE');
  if (!expected || String(req.code || '') !== expected) return json_({ ok: false, error: 'code' });
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
  var own = codeOf_(name);
  if (!own || String(req.scode || '').trim().toUpperCase() !== own.toUpperCase()) return json_({ ok: false, error: 'scode' });

  var bytes = Utilities.base64Decode(req.data);
  if (bytes.length > MAX_BYTES) return json_({ ok: false, error: 'size' });

  var ext = filename.replace(/^.*\./, '').toLowerCase();
  var stamp = Utilities.formatDate(new Date(), 'Europe/Luxembourg', 'yyyy-MM-dd HH-mm');
  var blob = Utilities.newBlob(bytes, req.mime || 'application/octet-stream', typ + ' - ' + name + ' - ' + stamp + '.' + ext);
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
  var neu = [];
  names.forEach(function (n) {
    if (have[n.toLowerCase()]) return;
    var c = '';
    for (var i = 0; i < 6; i++) c += CODE_ZEICHEN.charAt(Math.floor(Math.random() * CODE_ZEICHEN.length));
    neu.push([n, c]);
  });
  if (neu.length) cs.getRange(cs.getLastRow() + 1, 1, neu.length, 2).setValues(neu);
  return neu.length;
}

// ---------- Fiche données élève ----------
var FICHE_FELDER = [["classe", "Classe actuelle"], ["nom", "Nom"], ["prenom", "Prénom"], ["tel_mobile", "Téléphone mobile"], ["tel_fixe", "Téléphone fixe"], ["nationalite", "Nationalité"], ["matricule", "Matricule (13 chiffres)"], ["adresse", "Adresse postale"], ["email", "Adresse électronique"], ["iban", "Numéro de compte (IBAN)"], ["banque", "Nom de la banque ou CCP"], ["bic", "Code BIC"], ["titulaire", "Titulaire du compte"], ["rep_nom", "Représentant légal (nom et prénom)"], ["rep_adresse", "Adresse du représentant légal"], ["rep_tel", "Téléphone du représentant légal"], ["rep_email", "Adresse électronique du représentant légal"]];

function ficheSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Fiche données');
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
  var own = codeOf_(name);
  if (!own || String(req.scode || '').trim().toUpperCase() !== own.toUpperCase()) return json_({ ok: false, error: 'scode' });

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
      if (String(cell.getValue()).trim().toLowerCase() !== 'x') cell.setValue('i');
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
  var own = codeOf_(name);
  if (!own || String(req.scode || '').trim().toUpperCase() !== own.toUpperCase()) return json_({ ok: false, error: 'scode' });
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
