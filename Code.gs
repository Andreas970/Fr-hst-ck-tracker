// ===== Einstellungen =====
var TELEGRAM_TOKEN = "8826774200:AAHTRdLGAh5_8eXChQfmfRoMqupI-UnvtUA";
var TELEGRAM_CHAT_ID = "-1004487393812";
var ADMIN_CODE = "1907"; // Ändere dies auf einen eigenen Code
// Tab-Namen werden jetzt automatisch aus dem Sheet gelesen (kein manuelles Eintragen mehr nötig).

// ===== Web-App-Einstiegspunkt =====
function doGet(e) {
  e = e || {};
  var p = e.parameter || {};

  if (p.api === 'data') {
    return jsonOut({ ok: true, data: getAllData() });
  }

  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Frühstück-Tracker · Olive Inn')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ===== API-Einstiegspunkt für Aktionen (von der extern gehosteten App aufgerufen) =====
function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { body = {}; }
  var action = body.action;

  try {
    var result;
    switch (action) {
      case 'addEntry':
        result = addEntry(body.group, body.datum, body.anzahl, body.bemerkung, body.mitarbeiter);
        break;
      case 'updateEntry':
        result = updateEntry(body.group, body.row, body.datum, body.anzahl, body.bemerkung, body.mitarbeiter, body.code);
        break;
      case 'deleteEntry':
        result = deleteEntry(body.group, body.row, body.code);
        break;
      case 'verifyAdminCode':
        result = verifyAdminCode(body.code);
        break;
      case 'exportPdf':
        result = exportPdf(body.code);
        break;
      default:
        throw new Error('Unbekannte Aktion: ' + action);
    }
    return jsonOut({ ok: true, data: result });
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== Vom Frontend aufgerufen: alle Gruppen + Einträge laden =====
function getAllData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = {};
  ss.getSheets().forEach(function (sheet) {
    var name = sheet.getName();
    var lastRow = sheet.getLastRow();
    var entries = [];
    if (lastRow > 1) {
      var values = sheet.getRange(2, 1, lastRow - 1, 5).getDisplayValues();
      values.forEach(function (r, i) {
        if (r[0]) entries.push({ row: i + 2, datum: r[0], anzahl: r[1], bemerkung: r[2], mitarbeiter: r[3], erstelltAm: r[4] });
      });
    }
    data[name] = entries;
  });
  return data;
}

// ===== Tägliche Erinnerung: Seminarraum-Termine, die morgen stattfinden =====
var SEMINAR_SHEET_NAMES = ['seminarraum', 'seminar raum'];

function sendSeminarReminders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  var tomorrowStr = Utilities.formatDate(tomorrow, Session.getScriptTimeZone(), "yyyy-MM-dd");

  ss.getSheets().forEach(function (sheet) {
    if (SEMINAR_SHEET_NAMES.indexOf(sheet.getName().toLowerCase()) === -1) return;

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;
    var values = sheet.getRange(2, 1, lastRow - 1, 2).getDisplayValues();

    values.forEach(function (r) {
      var datum = r[0];
      var anzahl = r[1];
      if (datum === tomorrowStr) {
        var text = "Erinnerung: Morgen (" + datum + ") ist ein Termin im " + sheet.getName() +
                   (anzahl ? " – Anzahl: " + anzahl : "");
        sendTelegram(text);
      }
    });
  });
}

// ===== Vom Frontend aufgerufen: Admin-Code prüfen (ohne ihn preiszugeben) =====
function verifyAdminCode(code) {
  return code === ADMIN_CODE;
}

// ===== Vom Frontend aufgerufen: Eintrag bearbeiten (nur mit Admin-Code) =====
function updateEntry(groupName, rowNumber, datum, anzahl, bemerkung, mitarbeiter, code) {
  if (code !== ADMIN_CODE) throw new Error('Falscher Admin-Code.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(groupName);
  if (!sheet) throw new Error('Gruppe nicht gefunden: ' + groupName);
  sheet.getRange(rowNumber, 1, 1, 4).setValues([[datum, anzahl, bemerkung, mitarbeiter]]);
  return true;
}

// ===== Vom Frontend aufgerufen: Eintrag löschen (nur mit Admin-Code) =====
function deleteEntry(groupName, rowNumber, code) {
  if (code !== ADMIN_CODE) throw new Error('Falscher Admin-Code.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(groupName);
  if (!sheet) throw new Error('Gruppe nicht gefunden: ' + groupName);
  sheet.deleteRow(rowNumber);
  return true;
}

// ===== Vom Frontend aufgerufen: gesamtes Sheet als PDF exportieren (nur mit Admin-Code) =====
function exportPdf(code) {
  if (code !== ADMIN_CODE) throw new Error('Falscher Admin-Code.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var url = "https://docs.google.com/spreadsheets/d/" + ss.getId() +
            "/export?format=pdf&size=A4&portrait=true&fitw=true" +
            "&gridlines=true&printtitle=false&sheetnames=true&pagenumbers=false";
  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(url, { headers: { Authorization: "Bearer " + token } });
  return Utilities.base64Encode(response.getBlob().getBytes());
}

// ===== Vom Frontend aufgerufen: neuen Eintrag speichern + automatisch Telegram senden =====
function addEntry(groupName, datum, anzahl, bemerkung, mitarbeiter) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(groupName);
  if (!sheet) throw new Error('Gruppe nicht gefunden: ' + groupName);

  var erstelltAm = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm");
  sheet.appendRow([datum, anzahl, bemerkung, mitarbeiter, erstelltAm]);

  var text = "Neuer Eintrag – " + groupName + "\n" +
             datum + ": " + anzahl +
             (bemerkung ? " (" + bemerkung + ")" : "") +
             (mitarbeiter ? "\nErfasst von: " + mitarbeiter : "") +
             "\nErstellt am: " + erstelltAm;
  sendTelegram(text);

  return true;
}

function sendTelegram(text) {
  var url = "https://api.telegram.org/bot" + TELEGRAM_TOKEN + "/sendMessage";
  UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/x-www-form-urlencoded",
    payload: {
      chat_id: TELEGRAM_CHAT_ID,
      text: text
    }
  });
}
