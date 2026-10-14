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
  var req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_request' }); }

  var expected = PropertiesService.getScriptProperties().getProperty('LEHRER_CODE');
  if (!expected || String(req.code || '') !== expected) return json_({ ok: false, error: 'code' });
  if (req.action === 'check') return json_({ ok: true });
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
    return json_({ ok: true, data: readAll_() });
  } finally {
    lock.releaseLock();
  }
}
